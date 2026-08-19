#!/usr/bin/env python3
"""Gera o arquivo de atualização para quem administra o banco (se_pac).

Compara a planilha local (data/base_completa_atualizada_*.xlsx) com a tabela
se_pac.tab_selecao_novopac_previsto e produz atualizacao_banco_<data>.xlsx com:

- Leia-me: resumo das ações e a conciliação dos totais;
- Incluir_no_banco: propostas que existem na planilha e não existem no banco,
  com todas as colunas da tabela (data_atualizacao preenchida com a data da base);
- Trocar_status: propostas em comum que estão "enquadrada" no banco mas já
  constam "selecionada" na planilha (com a portaria antiga e a nova);
- Excluir_duplicadas: propostas com linhas repetidas no banco (mais linhas que
  na planilha) — manter uma linha e excluir as excedentes;
- Ajustar_valores: propostas com o mesmo nº de linhas mas valor divergente
  (valor correto = planilha);
- Revisar_divergencias: casos inversos de status (banco "selecionada" /
  planilha "enquadrada"), apenas para conferência manual — sem ação automática.

Uso:
  uv run --with pandas --with openpyxl --with sqlalchemy --with psycopg2-binary \
         --with python-dotenv python/gerar_atualizacao_banco.py

Requer python/config.env (credenciais — arquivo gitignorado).
"""
import os
from datetime import date

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL

HERE = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(HERE)
XLSX_LOCAL = os.path.join(RAIZ, 'data', 'base_completa_atualizada_20260818_1126.xlsx')
DATA_BASE_LOCAL = '2026-08-18'
SAIDA = os.path.join(RAIZ, f'atualizacao_banco_{date.today():%Y%m%d}.xlsx')

load_dotenv(os.path.join(HERE, 'config.env'))
url = URL.create(
    drivername='postgresql+psycopg2',
    username=os.getenv('DB_USER'),
    password=os.getenv('DB_PASSWORD'),
    host=os.getenv('DB_HOST'),
    port=int(os.getenv('DB_PORT', '5432')),
    database=os.getenv('DB_NAME'),
)
eng = create_engine(url, connect_args={'connect_timeout': 15})

banco = pd.read_sql(text('SELECT * FROM se_pac.tab_selecao_novopac_previsto'), eng)
local = pd.read_excel(XLSX_LOCAL, header=1)
print(f'banco: {len(banco)} linhas | planilha local: {len(local)} linhas')


def chave(df):
    return list(zip(df.num_proposta_selecao.astype(str).str.strip(),
                    df.identificador.astype(str).str.strip()))


local = local.assign(_k=chave(local))
banco = banco.assign(_k=chave(banco))
kb = set(banco._k)

# ---- 1) Incluir_no_banco: propostas da planilha ausentes do banco -------------
incluir = local[~local._k.isin(kb)].drop(columns=['_k']).copy()
incluir['data_atualizacao'] = DATA_BASE_LOCAL

# ---- 2/3) status por proposta (conjuntos, robusto a linhas duplicadas) --------
st_banco = banco.groupby('_k').agg(
    status_banco=('status_selecao', lambda s: ' | '.join(sorted(set(s)))),
    portaria_banco=('num_portaria', lambda s: ' | '.join(sorted({str(v) for v in s}))),
)
st_local = local.groupby('_k').agg(
    status_local=('status_selecao', lambda s: ' | '.join(sorted(set(s)))),
    portaria_local=('num_portaria', lambda s: ' | '.join(sorted({str(v) for v in s}))),
    modalidade=('modalidade', 'first'),
    uf=('uf', 'first'),
    municipio=('municipio', 'first'),
    proponente=('proponente', 'first'),
    ano_selecao=('ano_selecao', 'first'),
    vlr_portaria_total=('vlr_portaria_total', lambda s: pd.to_numeric(s, errors='coerce').fillna(0).sum()),
)
comum = st_local.join(st_banco, how='inner')

trocar = comum[(comum.status_banco == 'enquadrada') & comum.status_local.str.contains('selecionada')].copy()
revisar = comum[(comum.status_banco == 'selecionada') & (comum.status_local == 'enquadrada')].copy()

for df in (trocar, revisar):
    df.reset_index(inplace=True)
    if len(df):
        df[['num_proposta_selecao', 'identificador']] = pd.DataFrame(df['_k'].tolist(), index=df.index)
    else:
        df['num_proposta_selecao'] = df['identificador'] = pd.Series(dtype=str)
    df.drop(columns=['_k'], inplace=True)

ordem = ['num_proposta_selecao', 'identificador', 'modalidade', 'uf', 'municipio', 'proponente',
         'ano_selecao', 'status_banco', 'status_local', 'portaria_banco', 'portaria_local',
         'vlr_portaria_total']
trocar = trocar[ordem].rename(columns={'status_local': 'status_novo', 'portaria_local': 'portaria_nova'})
revisar = revisar[ordem]

# ---- 4/5) duplicidades e valores divergentes nas propostas em comum ----------
banco['_vlr'] = pd.to_numeric(banco.vlr_portaria_total, errors='coerce').fillna(0)
local['_vlr'] = pd.to_numeric(local.vlr_portaria_total, errors='coerce').fillna(0)
gb = banco[banco._k.isin(set(local._k))].groupby('_k').agg(
    linhas_banco=('_vlr', 'size'), valor_banco=('_vlr', 'sum'),
    portarias_banco=('num_portaria', lambda s: ' | '.join(sorted({str(v) for v in s}))))
gl = local.groupby('_k').agg(
    linhas_planilha=('_vlr', 'size'), valor_planilha=('_vlr', 'sum'),
    modalidade=('modalidade', 'first'), uf=('uf', 'first'), proponente=('proponente', 'first'))
cmp_ = gb.join(gl, how='inner')
cmp_['linhas_a_excluir'] = cmp_.linhas_banco - cmp_.linhas_planilha
cmp_['diferenca_valor'] = (cmp_.valor_banco - cmp_.valor_planilha).round(2)

duplicadas = cmp_[cmp_.linhas_a_excluir > 0].copy()
ajustar = cmp_[(cmp_.linhas_a_excluir == 0) & (cmp_.diferenca_valor.abs() > 0.01)].copy()
for df in (duplicadas, ajustar):
    df.reset_index(inplace=True)
    if len(df):
        df[['num_proposta_selecao', 'identificador']] = pd.DataFrame(df['_k'].tolist(), index=df.index)
    df.drop(columns=['_k'], inplace=True)
cols_cmp = ['num_proposta_selecao', 'identificador', 'modalidade', 'uf', 'proponente',
            'portarias_banco', 'linhas_banco', 'linhas_planilha', 'linhas_a_excluir',
            'valor_banco', 'valor_planilha', 'diferenca_valor']
duplicadas = duplicadas[cols_cmp]
ajustar = ajustar[cols_cmp]

# ---- resumo + gravação --------------------------------------------------------
mi = lambda v: round(v / 1e6, 2)  # noqa: E731
vlr_incluir = pd.to_numeric(incluir.vlr_portaria_total, errors='coerce').fillna(0).sum()
vlr_banco = banco._vlr.sum()
vlr_local = local._vlr.sum()
vlr_dup = duplicadas.diferenca_valor.sum() if len(duplicadas) else 0.0
vlr_aj = ajustar.diferenca_valor.sum() if len(ajustar) else 0.0

resumo = pd.DataFrame([
    ['Total hoje no banco (tab_selecao_novopac_previsto)', len(banco), mi(vlr_banco)],
    ['Total correto (planilha de referência)', len(local), mi(vlr_local)],
    ['1. Incluir_no_banco — inserir estas propostas', len(incluir), mi(vlr_incluir)],
    ['2. Trocar_status — enquadrada -> selecionada (portaria nova na aba)', len(trocar), None],
    ['3. Excluir_duplicadas — remover as linhas excedentes', int(duplicadas.linhas_a_excluir.sum()) if len(duplicadas) else 0, mi(-vlr_dup)],
    ['4. Ajustar_valores — corrigir para o valor da planilha', len(ajustar), mi(-vlr_aj)],
    ['Após as ações: banco deve fechar com a planilha', len(local), mi(vlr_banco + vlr_incluir - vlr_dup - vlr_aj)],
], columns=['Ação / conferência', 'Propostas ou linhas', 'Impacto no valor (R$ mi)'])

print(resumo.to_string(index=False))
print(f'Revisar_divergencias (banco selecionada / planilha enquadrada): {len(revisar)}')

with pd.ExcelWriter(SAIDA, engine='openpyxl') as w:
    resumo.to_excel(w, sheet_name='Leia-me', index=False)
    incluir.to_excel(w, sheet_name='Incluir_no_banco', index=False)
    trocar.to_excel(w, sheet_name='Trocar_status', index=False)
    duplicadas.to_excel(w, sheet_name='Excluir_duplicadas', index=False)
    if len(ajustar):
        ajustar.to_excel(w, sheet_name='Ajustar_valores', index=False)
    if len(revisar):
        revisar.to_excel(w, sheet_name='Revisar_divergencias', index=False)

print('gerado:', SAIDA)
