#!/usr/bin/env python3
"""Relatório antes × depois dos números da apresentação (index.html).

"Antes" e "depois" são lidos das linhas embutidas (const ROWS) de duas versões do
index.html e agregados exatamente como o navegador faz (sem filtro de UF, status ou ano),
então cada quadro repete o que o slide mostra.

- antes:  index.html de um commit git (padrão: HEAD)
- depois: index.html do diretório de trabalho

Também compara proposta a proposta as seleções (chave num_proposta_selecao +
identificador + fonte) entre a planilha antiga e a foto do banco — as migradas não têm
chave comum entre as duas fontes e são comparadas só nos agregados.

Saídas (raiz do projeto):
- relatorio_antes_depois_<AAAAMMDD>.xlsx
- relatorio_antes_depois_<AAAAMMDD>.json (insumo do relatorio_docx.js, que gera o .docx)

Uso: python3 python/relatorio_antes_depois.py [ref_git_antes]
"""
import glob
import json
import os
import re
import subprocess
import sys
from datetime import date

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(HERE)
sys.path.insert(0, RAIZ)
import dados  # noqa: E402

REF = sys.argv[1] if len(sys.argv) > 1 else 'HEAD'
HOJE = f'{date.today():%Y%m%d}'
SAIDA = os.path.join(RAIZ, f'relatorio_antes_depois_{HOJE}')
MCMV = ('MCMV FNHIS', 'MCMV FNHIS SUB50')
XLSX_ANTIGO = os.path.join(RAIZ, 'data', 'base_completa_18082026_1126.xlsx')


def base_antes():
    """Base do 'antes' na comparação proposta a proposta: a foto anterior à mais recente.

    Sem foto anterior (primeira extração do banco), cai na planilha de 18/08, que era a fonte
    da apresentação antes do banco. Pode ser passada no 2º argumento da linha de comando.
    """
    if len(sys.argv) > 2:
        return os.path.abspath(sys.argv[2])
    fotos = sorted(glob.glob(f'{RAIZ}/data/base_unica_gm_[0-9]*.xlsx'), key=dados.quando)
    return fotos[-2] if len(fotos) > 1 else XLSX_ANTIGO


# ------------------------------------------------------------------ leitura do index.html
def extrair(html):
    def const(nome):
        m = re.search(rf'^const {nome} = (.*);$', html, re.M)
        return json.loads(m.group(1))
    data = re.search(r'Atualizado em <b>(\d\d/\d\d/\d{4})</b>', html).group(1)
    versao = re.search(r'versão (\d+\.\d+)', html).group(1)
    return const('MODS'), const('ROWS'), data, versao


html_antes = subprocess.run(['git', '-C', RAIZ, 'show', f'{REF}:index.html'],
                            capture_output=True, text=True, check=True).stdout
html_depois = open(os.path.join(RAIZ, 'index.html'), encoding='utf-8').read()
MODS, ROWS_A, DATA_A, VER_A = extrair(html_antes)
MODS_D, ROWS_D, DATA_D, VER_D = extrair(html_depois)
assert MODS == MODS_D, 'lista de modalidades mudou entre as versões'

# linha: [modIdx, fonteIdx, uf, kind, gov, ogu, fin, ano] · fonte 0=FIN 1=OGU · kind 0=mig 1=sel 2=enq
DATASETS = {
    'all': lambda r: True,
    'mig': lambda r: r[3] == 0,
    'plan': lambda r: r[3] >= 1,
    'enq': lambda r: r[3] == 2,
    'gov': lambda r: r[4] == 1,
}
PREV = [MODS.index('Drenagem Urbana'), MODS.index('Contenção de Encostas')]
# (slide, título, dataset, modalidades, colunas) — espelha TABLES do build.py
SLIDES = [
    (2, 'Investimentos totais', 'all', None, ['n', 'tot']),
    (3, 'Investimentos por quantidade', 'all', None, ['qf', 'qo', 'n']),
    (4, 'Investimento por fonte', 'all', None, ['qf', 'qo', 'n', 'ogu', 'fin', 'tot']),
    (5, 'Novo PAC Migrado', 'mig', None, ['qf', 'qo', 'fin', 'ogu']),
    (6, 'Novas seleções', 'plan', None, ['qf', 'qo', 'fin', 'ogu']),
    (7, 'Propostas enquadradas/habilitadas', 'enq', None, ['qf', 'fin']),
    (8, 'Prevenção a Desastres', 'all', PREV, ['qf', 'qo', 'n', 'ogu', 'fin', 'tot']),
    (9, 'Governadores', 'gov', None, ['qf', 'qo', 'fin', 'ogu']),
]
NOME_COL = {'qf': 'Qtd. FIN', 'qo': 'Qtd. OGU', 'n': 'Qtd. total',
            'ogu': 'Valor OGU (R$ mi)', 'fin': 'Valor FIN (R$ mi)', 'tot': 'Valor total (R$ mi)'}
QTD = {'qf', 'qo', 'n'}


def zero():
    return {'qf': 0, 'qo': 0, 'n': 0, 'ogu': 0.0, 'fin': 0.0, 'tot': 0.0}


def somar(a, r):
    a['n'] += 1
    if r[1] == 0:
        a['qf'] += 1
    elif r[1] == 1:
        a['qo'] += 1
    a['ogu'] += r[5] / 1e6
    a['fin'] += r[6] / 1e6
    a['tot'] += (r[5] + r[6]) / 1e6


def agregar(rows, chave, filtro=lambda r: True):
    acc = {}
    for r in rows:
        if filtro(r):
            somar(acc.setdefault(chave(r), zero()), r)
    return acc


def comparar(nome_linha, acc_a, acc_d, cols, ordem):
    """Linhas: [rótulo, (antes, depois, dif) por coluna...] + TOTAL."""
    linhas, tot_a, tot_d = [], zero(), zero()
    for k in ordem:
        a, d = acc_a.get(k), acc_d.get(k)
        if not a and not d:
            continue
        a, d = a or zero(), d or zero()
        for c in tot_a:
            tot_a[c] += a[c]
            tot_d[c] += d[c]
        linhas.append([nome_linha(k)] + [(a[c], d[c], d[c] - a[c]) for c in cols])
    linhas.append(['TOTAL'] + [(tot_a[c], tot_d[c], tot_d[c] - tot_a[c]) for c in cols])
    return linhas


quadros = []  # {id, titulo, nota, colunas, linhas, grupos}

# ---- resumo geral
def recorte(rows, f):
    a = zero()
    for r in rows:
        if f(r):
            somar(a, r)
    return a


RECORTES = [
    ('Migradas', lambda r: r[3] == 0),
    ('Selecionadas', lambda r: r[3] == 1),
    ('Enquadradas/habilitadas', lambda r: r[3] == 2),
    ('Total da apresentação', lambda r: True),
]
res_linhas = []
for nome, f in RECORTES:
    a, d = recorte(ROWS_A, f), recorte(ROWS_D, f)
    res_linhas.append([nome] + [(a[c], d[c], d[c] - a[c]) for c in ['n', 'ogu', 'fin', 'tot']])
quadros.append({'id': 'Resumo', 'titulo': 'Resumo geral', 'rotulo': 'Recorte',
                'nota': 'Total da apresentação = migradas + selecionadas + enquadradas/habilitadas (FIN).',
                'colunas': ['n', 'ogu', 'fin', 'tot'], 'linhas': res_linhas, 'pct': True})

# ---- um quadro por slide
for num, titulo, ds, mods, cols in SLIDES:
    filtro = (lambda r, f=DATASETS[ds], m=mods: f(r) and (m is None or r[0] in m))
    acc_a = agregar(ROWS_A, lambda r: r[0], filtro)
    acc_d = agregar(ROWS_D, lambda r: r[0], filtro)
    quadros.append({'id': f'Slide {num}', 'titulo': f'Slide {num} — {titulo}', 'rotulo': 'Modalidade',
                    'nota': '', 'colunas': cols, 'pct': False,
                    'linhas': comparar(lambda k: MODS[k], acc_a, acc_d, cols, range(len(MODS)))})

# ---- por ano de seleção × status (novas seleções)
ANO_ORDEM = [(a, k) for a in ['2023', '2025', '2026'] for k in (1, 2)]
NOME_KIND = {1: 'selecionadas', 2: 'enquadradas/habilitadas'}
chave_ano = lambda r: (r[7], r[3])  # noqa: E731
acc_a = agregar(ROWS_A, chave_ano, DATASETS['plan'])
acc_d = agregar(ROWS_D, chave_ano, DATASETS['plan'])
quadros.append({'id': 'Por ano', 'titulo': 'Novas seleções por ano (botões do recorte) e status',
                'rotulo': 'Ano · status',
                'nota': 'Rótulos dos botões: 2023 = ano_selecao 2024 (e, a partir de agora, 2023); 2025; 2026.',
                'colunas': ['n', 'ogu', 'fin', 'tot'], 'pct': True,
                'linhas': comparar(lambda k: f'{k[0]} · {NOME_KIND[k[1]]}', acc_a, acc_d,
                                   ['n', 'ogu', 'fin', 'tot'], ANO_ORDEM)})

# ---- por UF (total da apresentação)
ufs = sorted({r[2] for r in ROWS_A} | {r[2] for r in ROWS_D})
acc_a = agregar(ROWS_A, lambda r: r[2])
acc_d = agregar(ROWS_D, lambda r: r[2])
quadros.append({'id': 'Por UF', 'titulo': 'Por UF — total da apresentação', 'rotulo': 'UF', 'nota': '',
                'colunas': ['n', 'ogu', 'fin', 'tot'], 'pct': True,
                'linhas': comparar(lambda k: k, acc_a, acc_d, ['n', 'ogu', 'fin', 'tot'], ufs)})


# ------------------------------------------------------------------ proposta a proposta (seleções)
def por_proposta(df):
    df = df.copy()
    df['num_proposta_selecao'] = df.num_proposta_selecao.astype(str).str.replace('﻿', '').str.strip()
    df['identificador'] = df.identificador.astype(str).str.strip()
    for c in ('vlr_portaria_ogu', 'vlr_portaria_fin'):
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    return df.groupby(['num_proposta_selecao', 'identificador', 'fonte']).agg(
        uf=('uf', 'first'), modalidade=('modalidade', 'first'), linhas=('uf', 'size'),
        status=('status_selecao', lambda s: ' + '.join(sorted(set(s)))),
        ano=('ano_selecao', lambda s: ' + '.join(str(int(v)) for v in sorted(set(s)))),
        ogu=('vlr_portaria_ogu', 'sum'), fin=('vlr_portaria_fin', 'sum'))


ANTES = base_antes()
FOTO_ANTES = os.path.basename(ANTES).startswith('base_unica_gm')
if FOTO_ANTES:  # foto do banco: mesmo recorte de seleções da apresentação
    antigo, _, DATA_ANTES, _ = dados.load(ANTES)
else:  # planilha antiga (cabeçalho na 2ª linha), que já vinha só com as seleções
    antigo, DATA_ANTES = pd.read_excel(dados.abrir(ANTES), header=1), None
antigo = antigo[~antigo.modalidade.isin(MCMV)]
x_novo, _, _, foto = dados.load()
x_novo = x_novo[~x_novo.modalidade.isin(MCMV)]
pa, pd_ = por_proposta(antigo), por_proposta(x_novo)
j = pa.join(pd_, how='outer', lsuffix='_antes', rsuffix='_depois')
TOL = 0.01


def situacao(r):
    if pd.isna(r.linhas_antes):
        return 'Entrou'
    if pd.isna(r.linhas_depois):
        return 'Saiu'
    mud = []
    if r.status_antes != r.status_depois:
        mud.append('status')
    if r.ano_antes != r.ano_depois:
        mud.append('ano')
    if abs(r.ogu_antes - r.ogu_depois) > TOL or abs(r.fin_antes - r.fin_depois) > TOL:
        mud.append('valor')
    if r.linhas_antes != r.linhas_depois:
        mud.append('nº de linhas')
    return 'Mudou: ' + ', '.join(mud) if mud else ''


j['situacao'] = j.apply(situacao, axis=1)
mud = j[j.situacao != ''].reset_index()
mud['uf'] = mud.uf_depois.fillna(mud.uf_antes)
mud['modalidade'] = mud.modalidade_depois.fillna(mud.modalidade_antes)
mud['dif_total_mi'] = ((mud.ogu_depois.fillna(0) + mud.fin_depois.fillna(0))
                       - (mud.ogu_antes.fillna(0) + mud.fin_antes.fillna(0))) / 1e6
mud = mud.sort_values(['situacao', 'dif_total_mi'], key=lambda s: s.abs() if s.name == 'dif_total_mi' else s,
                      ascending=[True, False])
COLS_MUD = ['situacao', 'num_proposta_selecao', 'identificador', 'fonte', 'uf', 'modalidade',
            'status_antes', 'status_depois', 'ano_antes', 'ano_depois',
            'ogu_antes', 'ogu_depois', 'fin_antes', 'fin_depois', 'dif_total_mi']
mud = mud[COLS_MUD]


def tipo(s):
    return 'Entrou' if s == 'Entrou' else 'Saiu' if s == 'Saiu' else 'Mudou'


contagem = mud.situacao.map(tipo).value_counts().to_dict()
det = {}
for s in mud.situacao:
    if s.startswith('Mudou: '):
        for m in s[7:].split(', '):
            det[m] = det.get(m, 0) + 1

foto_df = pd.read_excel(dados.abrir(foto))
n_2023 = int((foto_df[foto_df.status_selecao.isin(dados.SELECOES)].ano_selecao == 2023).sum())
n_subst = int((foto_df.status_selecao == 'substituída').sum())
ORIGEM_ANTES = (f'foto da tabela se_cgpac.tab_base_unica_gm ({os.path.basename(ANTES)})' if FOTO_ANTES
                else 'planilha base_completa_18082026_1126.xlsx e view_sis_novopac_previsto_18082026_0817.xlsx')
NOTAS = [
    f'Antes: index.html versão {VER_A} (dados de {DATA_A}) — {ORIGEM_ANTES}.',
    f'Depois: index.html versão {VER_D} (dados de {DATA_D}) — foto da tabela se_cgpac.tab_base_unica_gm '
    f'({os.path.basename(foto)}).',
    'Migradas: status_selecao = "retomada" no banco; as 2 propostas MCMV seguem fora da contagem.',
    f'Substituídas: {n_subst} {"linha" if n_subst == 1 else "linhas"} com status_selecao = "substituída" '
    'no banco (seleções trocadas por outras, guardadas só como histórico) ficam fora de toda a contagem.'
    if n_subst else
    'Substituídas: nenhuma linha com status_selecao = "substituída" na foto.',
    'OGU/FIN: as 5 propostas de Mobilidade que tinham uma linha "OGU/FIN" vêm em duas linhas no banco '
    '(uma OGU e uma FIN); cada linha conta 1 na sua fonte, então a quantidade total sobe 5 por esse motivo.'
    if not FOTO_ANTES else
    'OGU/FIN: cada linha conta 1 na sua fonte — propostas com OGU e FIN vêm em duas linhas.',
    f'Ano 2023: {n_2023} {"seleção" if n_2023 == 1 else "seleções"} com ano_selecao 2023 no banco '
    'conta como 2024 (botão "2023").',
    f'Proposta a proposta (seleções, chave nº da proposta + identificador + fonte): '
    f'{contagem.get("Entrou", 0)} entraram, {contagem.get("Saiu", 0)} saíram e {contagem.get("Mudou", 0)} mudaram '
    f'({", ".join(f"{k}: {v}" for k, v in sorted(det.items(), key=lambda t: -t[1]))}). '
    'Parte das entradas/saídas é só correção de chave (ex.: caractere invisível, número da proposta reescrito). '
    'Lista completa na planilha, aba "Propostas".',
]

# ------------------------------------------------------------------ Excel
AZUL, CINZA = '1351B4', 'EDF1F7'
fino = Side(style='thin', color='C5CCD6')
borda = Border(bottom=fino)
FMT_Q, FMT_V = '#,##0;[Red]-#,##0', '#,##0.00;[Red]-#,##0.00'
FMT_P = '0.0%;[Red]-0.0%'

with pd.ExcelWriter(SAIDA + '.xlsx', engine='openpyxl') as xw:
    wb = xw.book
    ws = wb.create_sheet('Leia-me')
    ws['A1'] = 'Novo PAC — MCid · Investimentos: relatório antes × depois'
    ws['A1'].font = Font(bold=True, size=14, color=AZUL)
    for i, n in enumerate(NOTAS, start=3):
        ws.cell(i, 1, n).alignment = Alignment(wrap_text=True, vertical='top')
    ws.column_dimensions['A'].width = 140

    for q in quadros:
        ws = wb.create_sheet(q['id'])
        ws['A1'] = q['titulo']
        ws['A1'].font = Font(bold=True, size=13, color=AZUL)
        if q['nota']:
            ws['A2'] = q['nota']
            ws['A2'].font = Font(italic=True, color='555555')
        cab1, cab2 = 4, 5
        ws.cell(cab1, 1, q['rotulo'])
        ws.merge_cells(start_row=cab1, start_column=1, end_row=cab2, end_column=1)
        col = 2
        sub = ['Antes', 'Depois', 'Diferença'] + (['Dif. %'] if q['pct'] else [])
        for c in q['colunas']:
            ws.cell(cab1, col, NOME_COL[c])
            ws.merge_cells(start_row=cab1, start_column=col, end_row=cab1, end_column=col + len(sub) - 1)
            for i, s in enumerate(sub):
                ws.cell(cab2, col + i, s)
            col += len(sub)
        for r in (cab1, cab2):
            for c in range(1, col):
                cel = ws.cell(r, c)
                cel.font = Font(bold=True, color='FFFFFF')
                cel.fill = PatternFill('solid', fgColor=AZUL)
                cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for li, linha in enumerate(q['linhas'], start=cab2 + 1):
            ws.cell(li, 1, linha[0])
            col = 2
            for c, (a, d, df) in zip(q['colunas'], linha[1:]):
                fmt = FMT_Q if c in QTD else FMT_V
                vals = [a, d, df]
                if q['pct']:
                    vals.append(df / a if a else None)
                for i, v in enumerate(vals):
                    cel = ws.cell(li, col + i, v)
                    cel.number_format = FMT_P if i == 3 else fmt
                    cel.border = borda
                col += len(sub)
            if linha[0] == 'TOTAL':
                for c in range(1, col):
                    ws.cell(li, c).font = Font(bold=True)
                    ws.cell(li, c).fill = PatternFill('solid', fgColor=CINZA)
        ws.column_dimensions['A'].width = 38
        for c in range(2, col):
            ws.column_dimensions[get_column_letter(c)].width = 14
        ws.freeze_panes = ws.cell(cab2 + 1, 2)

    mud_x = mud.copy()
    for c in ('ogu_antes', 'ogu_depois', 'fin_antes', 'fin_depois'):
        mud_x[c] = mud_x[c] / 1e6
    mud_x.columns = ['Situação', 'Nº proposta', 'Identificador', 'Fonte', 'UF', 'Modalidade',
                     'Status antes', 'Status depois', 'Ano antes', 'Ano depois',
                     'OGU antes (R$ mi)', 'OGU depois (R$ mi)', 'FIN antes (R$ mi)', 'FIN depois (R$ mi)',
                     'Dif. total (R$ mi)']
    mud_x.to_excel(xw, sheet_name='Propostas', index=False)
    ws = xw.sheets['Propostas']
    for cel in ws[1]:
        cel.font = Font(bold=True, color='FFFFFF')
        cel.fill = PatternFill('solid', fgColor=AZUL)
        cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for row in ws.iter_rows(min_row=2, min_col=11, max_col=15):
        for cel in row:
            cel.number_format = FMT_V
    for c, w in zip('ABCDEFGHIJKLMNO', [26, 24, 22, 8, 6, 32, 22, 22, 12, 12, 14, 14, 14, 14, 14]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = ws.dimensions
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

# ------------------------------------------------------------------ JSON para o Word
maiores = mud[mud.situacao != ''].reindex(mud.dif_total_mi.abs().sort_values(ascending=False).index).head(15)
json.dump({
    'data_antes': DATA_A, 'data_depois': DATA_D, 'versao_antes': VER_A, 'versao_depois': VER_D,
    'nome_col': NOME_COL, 'quadros': quadros, 'notas': NOTAS,
    'contagem': contagem, 'detalhe': det,
    'maiores': [[r.situacao, r.num_proposta_selecao, r.uf, r.modalidade, r.fonte, round(r.dif_total_mi, 2)]
                for r in maiores.itertuples()],
}, open(SAIDA + '.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

print('escrito:', os.path.relpath(SAIDA + '.xlsx', RAIZ), '|', os.path.relpath(SAIDA + '.json', RAIZ))
for linha in res_linhas:
    print(f'  {linha[0]:25} qtd {linha[1][0]:>6,.0f} → {linha[1][1]:>6,.0f} | total R$ mi {linha[4][0]:>12,.2f} → {linha[4][1]:>12,.2f}')
print('  propostas:', contagem, det)
