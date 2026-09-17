#!/usr/bin/env python3
"""Tira uma foto da tabela se_cgpac.tab_base_unica_gm para data/.

A apresentação (build.py e build_detalhada.py) lê essa foto — não conecta no banco —
para o index.html continuar autocontido e funcionando offline.

Gera data/base_unica_gm_<DDMMYYYY>_<HHMM>.xlsx, com a data/hora da carga mais recente
(max(dte_carga)) — uma versão por extração completa do banco.

A foto vai para o repositório (público) e é o arquivo do botão "XLSX Base" do painel, por isso:
- leva só as colunas que a planilha pública anterior já trazia + dte_carga (as colunas de
  controle interno da tabela ficam de fora);
- é gravada protegida por senha (SENHA_BASE_XLSX no config.env), com a criptografia padrão
  do Excel. A senha é distribuída fora do repositório.

Uso:
  uv run --with pandas --with openpyxl --with sqlalchemy --with psycopg2-binary \
         --with python-dotenv --with msoffcrypto-tool python/extrair_base_unica.py

Requer python/config.env (credenciais e senha — arquivo gitignorado).
"""
import io
import os

import msoffcrypto
import pandas as pd
from dotenv import load_dotenv
from msoffcrypto.format.ooxml import OOXMLFile
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL

HERE = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(HERE)
TABELA = 'se_cgpac.tab_base_unica_gm'
COLUNAS = [
    'num_proposta_selecao', 'fonte', 'subfonte', 'regiao', 'uf', 'municipio', 'proponente', 'descricao',
    'vlr_portaria_total', 'vlr_portaria_ogu', 'vlr_portaria_fin', 'eixo', 'subeixo', 'modalidade',
    'grupo_modalidade', 'sigla_modalidade', 'sigla_grupo_modalidade', 'identificador', 'num_portaria',
    'link', 'observacao', 'bln_alteracao', 'status_alteracao', 'data_envio_cc', 'data_resposta_cc',
    'link_alteracao', 'ano_selecao', 'status_selecao', 'tipo_financiamento', 'dte_carga',
]

load_dotenv(os.path.join(HERE, 'config.env'))
SENHA = os.getenv('SENHA_BASE_XLSX')
if not SENHA:
    raise SystemExit('Defina SENHA_BASE_XLSX em python/config.env — a base publicada vai com senha.')

url = URL.create(
    drivername='postgresql+psycopg2',
    username=os.getenv('DB_USER'),
    password=os.getenv('DB_PASSWORD'),
    host=os.getenv('DB_HOST'),
    port=int(os.getenv('DB_PORT', '5432')),
    database=os.getenv('DB_NAME'),
)
# sessão somente leitura: o script nunca altera o banco
eng = create_engine(url, connect_args={'connect_timeout': 15,
                                       'options': '-c default_transaction_read_only=on'})

df = pd.read_sql(text(f'SELECT {", ".join(COLUNAS)} FROM {TABELA} ORDER BY id_linha'), eng)
carga = pd.to_datetime(df.dte_carga).max()
saida = os.path.join(RAIZ, 'data', f'base_unica_gm_{carga:%d%m%Y_%H%M}.xlsx')

# grava em memória e só escreve em disco já cifrado — o arquivo aberto nunca chega a existir
buf = io.BytesIO()
df.to_excel(buf, index=False, sheet_name='base_unica_gm')
buf.seek(0)
with open(saida, 'wb') as f:
    OOXMLFile(buf).encrypt(SENHA, f)

# confere que o arquivo publicado abre com a senha (e só com ela)
with open(saida, 'rb') as f:
    protegido = msoffcrypto.OfficeFile(f)
    if not protegido.is_encrypted():
        raise SystemExit('ERRO: a foto foi gravada sem senha — não publique este arquivo.')
    protegido.load_key(password=SENHA)
    aberto = io.BytesIO()
    protegido.decrypt(aberto)
if len(pd.read_excel(aberto)) != len(df):
    raise SystemExit('ERRO: a foto protegida não reabriu com o mesmo conteúdo.')

print(f'{TABELA}: {len(df)} linhas | última carga {carga:%d/%m/%Y %H:%M}')
print(df.status_selecao.value_counts(dropna=False).to_string())
print(f'foto (protegida por senha): {os.path.relpath(saida, RAIZ)}')
