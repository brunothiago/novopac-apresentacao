#!/usr/bin/env python3
"""Gera versao-detalhada.html — deck com painéis de destaque, barras proporcionais e
slides extras (Calamidade RS e Seleções 2026), sem filtro de UF.

Fontes de dados (única fonte de verdade — números computados aqui, nunca transcritos):
- data/base_completa_atualizada_20260818_1126.xlsx (seleções 2024-2026)
- data/view_sis_novopac_previsto_unificado_202608180817.csv (migradas)
Convenção da apresentação antiga: totais = migradas + selecionadas + enquadradas FIN.
"""
import os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/versao-detalhada.html'
VAL = ['vlr_portaria_total', 'vlr_portaria_ogu', 'vlr_portaria_fin']

LABELS = {
    'Médias e Grandes Cidades': 'Mobilidade: Médias e Grandes Cidades',
    'MCMV FNHIS': 'MCMV (Calamidade RS)',
    'MCMV FNHIS SUB50': 'MCMV (Calamidade RS)',
}
ORDER = [
    'Abastecimento de Água - Rural', 'Abastecimento de Água - Urbano',
    'Contenção de Encostas', 'Drenagem Urbana', 'Esgotamento Sanitário',
    'MCMV (Calamidade RS)', 'Mobilidade: Médias e Grandes Cidades',
    'Regularização Fundiária', 'Renovação de Frota', 'Resíduos Sólidos',
    'Urbanização de Favelas',
]


def ptbr(v, dec=2):
    s = f'{v:,.{dec}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def ptint(v):
    return f'{int(v):,}'.replace(',', '.')


def load():
    x = pd.read_excel(f'{HERE}/data/base_completa_atualizada_20260818_1126.xlsx', header=1)
    v = pd.read_csv(f'{HERE}/data/view_sis_novopac_previsto_unificado_202608180817.csv')
    mig = v[v.origem_dado == 'Novo PAC - Retomada'].copy()
    for d in (x, mig):
        for c in VAL:
            d[c] = pd.to_numeric(d[c], errors='coerce').fillna(0)
        d['mod'] = d.modalidade.replace(LABELS)
    return x, mig


def agg(d):
    g = d.groupby('mod').agg(
        qFIN=('fonte', lambda s: int((s == 'FIN').sum())),
        qOGU=('fonte', lambda s: int((s == 'OGU').sum())),
        qMIX=('fonte', lambda s: int((s == 'OGU/FIN').sum())),
        n=('fonte', 'size'),
        ogu=('vlr_portaria_ogu', lambda s: s.sum() / 1e6),
        fin=('vlr_portaria_fin', lambda s: s.sum() / 1e6),
        tot=('vlr_portaria_total', lambda s: s.sum() / 1e6),
    ).reindex([m for m in ORDER if m in set(d['mod'])])
    return g


x, mig = load()
tot = agg(pd.concat([x[['mod', 'fonte'] + VAL], mig[['mod', 'fonte'] + VAL]]))
migr = agg(mig)
novas = agg(x)
enq = agg(x[x.status_selecao == 'enquadrada'])
gov = agg(x[x.grupo_modalidade == 'Governadores'])
cal = agg(x[x.grupo_modalidade == 'Calamidade RS'])
s26 = x[x.ano_selecao == 2026]
a26 = agg(s26)
g26 = s26.groupby(['mod', 'status_selecao']).size().unstack(fill_value=0)

T = lambda g, c: g[c].sum()  # noqa: E731

# ---------------------------------------------------------------- helpers HTML
def row_qtd_bar(g, m, maxn):
    """linha: modalidade | FIN | OGU | total | barra empilhada"""
    r = g.loc[m]
    wf, wo = r.qFIN / maxn * 100, r.qOGU / maxn * 100
    segs = ''
    if r.qFIN:
        segs += f'<i class="sg fin" style="width:{wf:.2f}%"></i>'
    if r.qOGU:
        segs += f'<i class="sg ogu" style="width:{wo:.2f}%"></i>'
    return (f'<tr><td>{m}</td><td>{ptint(r.qFIN)}</td><td>{ptint(r.qOGU)}</td>'
            f'<td class="tt">{ptint(r.n)}</td><td class="barcell"><span class="tbar">{segs}</span></td></tr>')


def row_val_share(g, m):
    """linha: modalidade | OGU | FIN | total | barra de participação OGU vs FIN"""
    r = g.loc[m]
    share_ogu = r.ogu / r.tot * 100 if r.tot else 0
    share_fin = 100 - share_ogu
    segs = ''
    if r.ogu > 0:
        segs += f'<i class="sg ogu" style="width:{share_ogu:.2f}%"></i>'
    if r.fin > 0:
        segs += f'<i class="sg fin" style="width:{share_fin:.2f}%"></i>'
    return (f'<tr><td>{m}</td><td>{ptbr(r.ogu)}</td><td>{ptbr(r.fin)}</td>'
            f'<td class="tt">{ptbr(r.tot)}</td><td class="barcell"><span class="tbar full">{segs}</span></td></tr>')


def row_tot_bar(g, m, maxv):
    r = g.loc[m]
    w = r.tot / maxv * 100
    return (f'<tr><td>{m}</td><td>{ptint(r.n)}</td><td class="tt">{ptbr(r.tot)}</td>'
            f'<td class="barcell"><span class="tbar"><i class="sg fin" style="width:{w:.2f}%"></i></span></td></tr>')


def row_4col(g, m):
    r = g.loc[m]
    return (f'<tr><td>{m}</td><td>{ptint(r.qFIN)}</td><td>{ptint(r.qOGU)}</td>'
            f'<td>{ptbr(r.fin)}</td><td>{ptbr(r.ogu)}</td></tr>')


# ------------------------------------------------------------------- tabelas
max_tot = tot.tot.max()
rows_s2 = '\n'.join(row_tot_bar(tot, m, max_tot) for m in tot.index)
tr_s2 = (f'<tr class="tot"><td>TOTAL</td><td>{ptint(T(tot,"n"))}</td>'
         f'<td class="tt">{ptbr(T(tot,"tot"))}</td><td class="barcell"></td></tr>')

max_n = tot.n.max()
rows_s3 = '\n'.join(row_qtd_bar(tot, m, max_n) for m in tot.index)
tr_s3 = (f'<tr class="tot"><td>TOTAL</td><td>{ptint(T(tot,"qFIN"))}</td><td>{ptint(T(tot,"qOGU"))}</td>'
         f'<td class="tt">{ptint(T(tot,"n"))}</td><td class="barcell"></td></tr>')

rows_s4 = '\n'.join(row_val_share(tot, m) for m in tot.index)
share_ogu_all = T(tot, 'ogu') / T(tot, 'tot') * 100
share_fin_all = 100 - share_ogu_all
tr_s4 = (f'<tr class="tot"><td>TOTAL</td><td>{ptbr(T(tot,"ogu"))}</td><td>{ptbr(T(tot,"fin"))}</td>'
         f'<td class="tt">{ptbr(T(tot,"tot"))}</td><td class="barcell"><span class="tbar full">'
         f'<i class="sg ogu" style="width:{share_ogu_all:.2f}%"></i>'
         f'<i class="sg fin" style="width:{share_fin_all:.2f}%"></i></span></td></tr>')

rows_s5 = '\n'.join(row_4col(migr, m) for m in migr.index)
tr_s5 = (f'<tr class="tot"><td>TOTAL</td><td>{ptint(T(migr,"qFIN"))}</td><td>{ptint(T(migr,"qOGU"))}</td>'
         f'<td>{ptbr(T(migr,"fin"))}</td><td>{ptbr(T(migr,"ogu"))}</td></tr>')

rows_s6 = '\n'.join(row_4col(novas, m) for m in novas.index)
tr_s6 = (f'<tr class="tot"><td>TOTAL</td><td>{ptint(T(novas,"qFIN"))}</td><td>{ptint(T(novas,"qOGU"))}</td>'
         f'<td>{ptbr(T(novas,"fin"))}</td><td>{ptbr(T(novas,"ogu"))}</td></tr>')

max_enq = enq.tot.max()
rows_s7 = '\n'.join(
    (lambda r, w: f'<tr><td>{m}</td><td>{ptint(r.qFIN)}</td><td class="tt">{ptbr(r.fin)}</td>'
     f'<td class="barcell"><span class="tbar"><i class="sg fin" style="width:{w:.2f}%"></i></span></td></tr>')
    (enq.loc[m], enq.loc[m].tot / max_enq * 100) for m in enq.index)
tr_s7 = (f'<tr class="tot"><td>TOTAL</td><td>{ptint(T(enq,"qFIN"))}</td>'
         f'<td class="tt">{ptbr(T(enq,"fin"))}</td><td class="barcell"></td></tr>')

rows_s8 = '\n'.join(row_4col(gov, m) for m in gov.index)
tr_s8 = (f'<tr class="tot"><td>TOTAL</td><td>{ptint(T(gov,"qFIN"))}</td><td>{ptint(T(gov,"qOGU"))}</td>'
         f'<td>{ptbr(T(gov,"fin"))}</td><td>{ptbr(T(gov,"ogu"))}</td></tr>')

# slide 10 — 2026 (tudo FIN): modalidade | selecionadas | enquadradas | valor | barra sel/enq
max26 = a26.n.max()
rows_s10 = []
for m in a26.index:
    r = a26.loc[m]
    nsel = int(g26.loc[m].get('selecionada', 0))
    nenq = int(g26.loc[m].get('enquadrada', 0))
    ws, we = nsel / max26 * 100, nenq / max26 * 100
    segs = ''
    if nsel:
        segs += f'<i class="sg sel" style="width:{ws:.2f}%"></i>'
    if nenq:
        segs += f'<i class="sg enq" style="width:{we:.2f}%"></i>'
    rows_s10.append(f'<tr><td>{m}</td><td>{nsel}</td><td>{nenq}</td><td class="tt">{ptbr(r.tot)}</td>'
                    f'<td class="barcell sm"><span class="tbar">{segs}</span></td></tr>')
rows_s10 = '\n'.join(rows_s10)
n26_sel = int((s26.status_selecao == 'selecionada').sum())
n26_enq = int((s26.status_selecao == 'enquadrada').sum())
v26_sel = s26[s26.status_selecao == 'selecionada'].vlr_portaria_total.sum() / 1e6
v26_enq = s26[s26.status_selecao == 'enquadrada'].vlr_portaria_total.sum() / 1e6
tr_s10 = (f'<tr class="tot"><td>TOTAL</td><td>{n26_sel}</td><td>{n26_enq}</td>'
          f'<td class="tt">{ptbr(T(a26,"tot"))}</td><td class="barcell sm"></td></tr>')

# números de destaque
N_ALL, V_ALL = ptint(T(tot, 'n')), ptbr(T(tot, 'tot'))
V_ALL_BI = ptbr(T(tot, 'tot') / 1000, 1)
V_FIN_BI = ptbr(T(tot, 'fin') / 1000, 1)
V_OGU_BI = ptbr(T(tot, 'ogu') / 1000, 1)
N_MIG, V_MIG_BI = ptint(T(migr, 'n')), ptbr(T(migr, 'tot') / 1000, 1)
N_NOV, V_NOV_BI = ptint(T(novas, 'n')), ptbr(T(novas, 'tot') / 1000, 1)
N_ENQ, V_ENQ_BI = ptint(T(enq, 'n')), ptbr(T(enq, 'fin') / 1000, 1)
N_GOV, V_GOV_BI = ptint(T(gov, 'n')), ptbr(T(gov, 'tot') / 1000, 1)
N_CAL, V_CAL = ptint(T(cal, 'n')), ptbr(T(cal, 'tot'))
N_26, V_26_BI = ptint(T(a26, 'n')), ptbr(T(a26, 'tot') / 1000, 1)
cal_dre = cal.loc['Drenagem Urbana']
cal_mcmv = cal.loc['MCMV (Calamidade RS)']

# ------------------------------------------------------- assets do deck (gov.br)
DECK_CSS = '<style>' + open(f'{HERE}/assets/deck.css', encoding='utf-8').read() + '</style>'
DECK_TAIL = open(f'{HERE}/assets/deck-chrome.html', encoding='utf-8').read()
# escapa "</script" dentro do JS (ocorre em comentário de uso) para não fechar a tag inline
DECK_JS = open(f'{HERE}/assets/deck-stage.js', encoding='utf-8').read().replace('</script', '<\\/script')

EXTRA_CSS = """
<style>
  body{ margin:0; background:var(--navy); }
  /* ---- deck de dados: tabelas compactas + barras ---- */
  table.data{ width:100%; border-collapse:collapse; }
  table.data th{ text-align:right; font-size:21px; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--ink-soft); padding:0 20px 14px; border-bottom:2px solid var(--line-strong); white-space:nowrap; }
  table.data th:first-child{ text-align:left; padding-left:0; }
  table.data th.bh{ text-align:left; }
  table.data td{ padding:11px 20px; border-bottom:1px solid var(--line); font-size:25px; color:var(--ink); font-variant-numeric:tabular-nums; text-align:right; white-space:nowrap; vertical-align:middle; }
  table.data td:first-child{ text-align:left; padding-left:0; font-weight:600; }
  table.data td.tt{ font-weight:800; color:var(--navy); }
  table.data tr.tot td{ border-top:2px solid var(--line-strong); border-bottom:0; font-weight:800; color:var(--navy); }
  table.data td.barcell{ width:380px; padding-right:0; }
  table.data td.barcell.sm{ width:260px; }
  /* variante densa — tabelas de 12 linhas */
  table.data.dense th{ font-size:19px; padding:0 16px 12px; }
  table.data.dense td{ padding:7px 16px; font-size:23px; }
  .body.tight{ padding-top:26px; }
  .title-block.tight{ margin-bottom:26px; }
  .title-row.tight{ margin-bottom:26px; }
  .tbar{ display:flex; width:100%; height:23px; background:rgba(7,29,65,.05); }
  .tbar.full{ background:transparent; }
  .tbar .sg{ display:block; height:100%; flex:none; }
  .tbar .sg + .sg{ margin-left:2px; }
  .sg.fin{ background:var(--blue); }
  .sg.ogu{ background:var(--green); }
  .sg.sel{ background:var(--blue); }
  .sg.enq{ background:var(--blue-soft); }
  @media (prefers-reduced-motion: no-preference){
    section[data-deck-active] .tbar .sg{ animation:barGrow .8s .25s cubic-bezier(.25,.8,.3,1) backwards; transform-origin:left center; }
  }
  /* legenda */
  .leg{ display:flex; gap:34px; align-items:center; }
  .leg .lg{ display:flex; align-items:center; gap:11px; font-size:22px; font-weight:600; color:var(--ink-soft); }
  .leg .sw{ width:16px; height:16px; flex:none; }
  /* header de slide com legenda à direita */
  .title-row{ display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:var(--gap); }
  .title-row .title-block{ margin-bottom:0; }
  /* painel lateral de destaque */
  .sidepanel{ flex:none; width:27%; display:flex; flex-direction:column; justify-content:center; gap:8px; padding-right:56px; border-right:1px solid var(--line); margin-right:56px; }
  .sidepanel .bignum{ font-size:120px; }
  .sidepanel .stat2{ font-size:56px; font-weight:800; letter-spacing:-.02em; color:var(--navy); font-variant-numeric:tabular-nums; margin-top:26px; }
  .sidepanel .lbl{ font-size:23px; font-weight:700; letter-spacing:.13em; text-transform:uppercase; color:var(--ink-soft); margin-top:8px; }
  /* barra de participação total (rodapé slide fonte) */
  .sharebar{ margin-top:34px; }
  .sharebar .stack{ height:52px; }
  .sharebar .stack .seg{ font-size:26px; }
</style>
"""

# ---------------------------------------------------------------------- head
HEAD = """<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Novo PAC — MCid · Investimentos</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin="">
<!-- Tipografia oficial do Governo Federal (Padrão Digital de Governo — gov.br/ds) -->
<link rel="stylesheet" href="https://cdngovbr-ds.estaleiro.serpro.gov.br/design-system/fonts/rawline/css/rawline.css">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,300;1,400;1,500;1,600;1,700&display=swap">
"""

CHROME_TOP = """    <div class="top">
      <div class="eyebrow">Novo PAC — MCid · Investimentos</div>
      <div class="brand"><span class="govbr">gov<b>.br</b></span><span class="div"></span><span class="min">Ministério das Cidades</span></div>
    </div>
    <hr class="hr">
"""


def foot(src='Fonte: MCid · Portarias de seleção do Novo PAC · dados de 18/08/2026'):
    return f"""    <hr class="hr">
    <div class="bot">
      <span class="org">{src}</span>
      <span class="pageno"></span>
    </div>
"""


LEG_FIN_OGU = """<div class="leg"><span class="lg"><span class="sw" style="background:var(--blue)"></span>FIN · Financiamento</span><span class="lg"><span class="sw" style="background:var(--green)"></span>OGU · Orçamento Geral da União</span></div>"""
LEG_SEL_ENQ = """<div class="leg"><span class="lg"><span class="sw" style="background:var(--blue)"></span>Selecionadas</span><span class="lg"><span class="sw" style="background:var(--blue-soft)"></span>Enquadradas</span></div>"""

# --------------------------------------------------------------------- slides
SLIDES = f"""
<deck-stage width="1920" height="1080" no-rail="">

  <!-- ============ 01 · CAPA ============ -->
  <section class="p2 dark cover" data-label="Capa" style="position:relative; overflow:hidden; z-index:0;">
    <svg class="covernet" viewBox="0 0 460 900" fill="none" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
      <g stroke="var(--blue-soft)" stroke-width="1.6" stroke-linecap="round">
        <line x1="120" y1="150" x2="300" y2="100"></line><line x1="300" y1="100" x2="390" y2="270"></line>
        <line x1="120" y1="150" x2="230" y2="310"></line><line x1="230" y1="310" x2="390" y2="270"></line>
        <line x1="230" y1="310" x2="95" y2="370"></line><line x1="230" y1="310" x2="330" y2="480"></line>
        <line x1="330" y1="480" x2="390" y2="270"></line><line x1="330" y1="480" x2="300" y2="670"></line>
        <line x1="230" y1="310" x2="185" y2="550"></line><line x1="185" y1="550" x2="110" y2="650"></line>
        <line x1="185" y1="550" x2="300" y2="670"></line><line x1="300" y1="670" x2="250" y2="810"></line>
        <line x1="110" y1="650" x2="250" y2="810"></line>
      </g>
      <g fill="var(--blue-soft)">
        <circle cx="120" cy="150" r="5"></circle><circle cx="390" cy="270" r="5"></circle><circle cx="95" cy="370" r="5"></circle>
        <circle cx="330" cy="480" r="5"></circle><circle cx="185" cy="550" r="5"></circle><circle cx="110" cy="650" r="5"></circle><circle cx="250" cy="810" r="5"></circle>
      </g>
      <g fill="var(--amber-bright)">
        <circle cx="300" cy="100" r="6.5"></circle><circle cx="230" cy="310" r="8"></circle><circle cx="300" cy="670" r="6.5"></circle>
      </g>
    </svg>
    <div class="top">
      <div class="eyebrow" style="letter-spacing:.2em;">Programa de Aceleração do Crescimento</div>
      <div class="brand">
        <span class="govbr" style="color:#fff;">gov<b style="color:var(--amber-bright);">.br</b></span>
        <span class="div"></span>
        <span class="min">Ministério das Cidades</span>
      </div>
    </div>
    <hr class="hr">

    <div class="body" style="justify-content:center; padding-top:0;">
      <div class="label" style="color:var(--amber-bright); margin-bottom:30px;">Investimentos</div>
      <h1 style="font-size:150px; font-weight:800; line-height:.98; letter-spacing:-.03em; margin:0; max-width:20ch;">Novo PAC<span style="color:var(--blue-soft); font-weight:500;"> — MCid</span></h1>
      <p style="font-family:'Raleway',sans-serif; font-style:italic; font-weight:500; font-size:38px; line-height:1.4; color:var(--paper); margin:190px 0 0; max-width:44ch;">Balanço da carteira de seleções: migradas, novas seleções, propostas enquadradas e destaques — {N_ALL} propostas e R$ {V_ALL_BI} bilhões em investimentos.</p>
    </div>

    <hr class="hr">
    <div class="bot" style="border-top:0; padding-top:18px;">
      <span class="org" style="color:var(--paper);">Ministério das Cidades</span>
      <div style="flex: 1;"></div>
      <span class="org" style="color:var(--paper);">Agosto de 2026</span>
    </div>
  </section>

  <!-- ============ 02 · INVESTIMENTOS TOTAIS ============ -->
  <section class="p2 content" data-label="Investimentos Totais">
{CHROME_TOP}    <div class="body tight">
      <div class="title-block tight">
        <h2 class="title">Investimentos totais</h2>
        <p class="dek" style="max-width:80ch;">Migradas, novas seleções e propostas enquadradas.</p>
      </div>
      <div class="row grow" style="align-items:stretch;">
        <div class="sidepanel">
          <div class="label">Valor total</div>
          <div class="bignum"><span class="u">R$ </span>{V_ALL_BI}<span class="u"> bi</span></div>
          <div class="stat2">{N_ALL}</div>
          <div class="lbl">Propostas</div>
        </div>
        <div class="col grow" style="justify-content:center;">
          <table class="data dense">
            <thead><tr><th>Modalidade</th><th>Propostas</th><th>Valor total (R$ mi)</th><th class="bh" style="padding-left:24px;">Proporção do valor</th></tr></thead>
            <tbody>
{rows_s2}
{tr_s2}
            </tbody>
          </table>
        </div>
      </div>
      <div class="take"><span class="tick"></span><p>151 propostas e R$ 13,2 bi a mais que no balanço anterior.</p></div>
    </div>
{foot()}  </section>

  <!-- ============ 03 · INVESTIMENTOS POR QUANTIDADE ============ -->
  <section class="p2 content" data-label="Por quantidade">
{CHROME_TOP}    <div class="body tight">
      <div class="title-row tight">
        <div class="title-block tight" style="margin-bottom:0;">
          <h2 class="title">Investimentos por quantidade</h2>
          <p class="dek" style="max-width:80ch;">Quantidade de propostas por modalidade e fonte de recursos.</p>
        </div>
        {LEG_FIN_OGU}
      </div>
      <div class="col grow" style="justify-content:center;">
        <table class="data dense">
          <thead><tr><th>Modalidade</th><th>Qtd. FIN</th><th>Qtd. OGU</th><th>Total</th><th class="bh" style="padding-left:24px;">Propostas por fonte</th></tr></thead>
          <tbody>
{rows_s3}
{tr_s3}
          </tbody>
        </table>
      </div>
      <div class="take"><span class="tick"></span><p>O OGU lidera em quantidade; o financiamento concentra o valor.</p></div>
    </div>
{foot('Fonte: MCid · Novo PAC · 18/08/2026 · Cinco propostas de Mobilidade combinam OGU e FIN e contam apenas no total')}  </section>

  <!-- ============ 04 · INVESTIMENTO POR FONTE ============ -->
  <section class="p2 content" data-label="Por fonte">
{CHROME_TOP}    <div class="body tight">
      <div class="title-row tight">
        <div class="title-block tight" style="margin-bottom:0;">
          <h2 class="title">Investimento por fonte</h2>
          <p class="dek" style="max-width:80ch;">Valores por modalidade, em R$ milhões.</p>
        </div>
        {LEG_FIN_OGU}
      </div>
      <div class="col grow" style="justify-content:center;">
        <table class="data dense">
          <thead><tr><th>Modalidade</th><th>Valor OGU</th><th>Valor FIN</th><th>Valor total</th><th class="bh" style="padding-left:24px;">Participação por fonte</th></tr></thead>
          <tbody>
{rows_s4}
{tr_s4}
          </tbody>
        </table>
      </div>
      <div class="take"><span class="tick"></span><p>O financiamento responde por {share_fin_all:.0f}% do valor: R$ {V_FIN_BI} bi, ante R$ {V_OGU_BI} bi do OGU.</p></div>
    </div>
{foot()}  </section>

  <!-- ============ 05 · NOVO PAC MIGRADO ============ -->
  <section class="p2 content" data-label="Novo PAC Migrado">
{CHROME_TOP}    <div class="body">
      <div class="title-block">
        <h2 class="title">Novo PAC Migrado</h2>
        <p class="dek" style="font-style:italic;">(Valores: Empenho de OGU e Pago de FIN) &gt; dez/2022</p>
      </div>
      <div class="row grow" style="align-items:stretch;">
        <div class="sidepanel">
          <div class="label">Carteira migrada</div>
          <div class="bignum">{N_MIG}</div>
          <div class="lbl">Propostas</div>
          <div class="stat2"><span style="font-size:.55em; font-weight:700; color:var(--ink-soft);">R$ </span>{V_MIG_BI}<span style="font-size:.55em; font-weight:700; color:var(--ink-soft);"> bi</span></div>
          <div class="lbl">Valor total</div>
        </div>
        <div class="col grow" style="justify-content:center;">
          <table class="data">
            <thead><tr><th>Modalidade</th><th>Qtd. FIN</th><th>Qtd. OGU</th><th>Valor FIN (R$ mi)</th><th>Valor OGU (R$ mi)</th></tr></thead>
            <tbody>
{rows_s5}
{tr_s5}
            </tbody>
          </table>
        </div>
      </div>
      <div class="take"><span class="tick"></span><p>A carteira herdada de dezembro de 2022 segue idêntica ao balanço anterior: {N_MIG} contratos migrados.</p></div>
    </div>
{foot('Fonte: MCid · Novo PAC — carteira migrada (retomadas) · dados de 18/08/2026')}  </section>

  <!-- ============ 06 · NOVAS SELEÇÕES ============ -->
  <section class="p2 content" data-label="Novas Seleções">
{CHROME_TOP}    <div class="body tight">
      <div class="title-block tight">
        <h2 class="title">Novas seleções</h2>
        <p class="dek" style="max-width:80ch;">2024, 2025 e 2026 (sem migradas) — inclui propostas enquadradas FIN, como no balanço anterior.</p>
      </div>
      <div class="col grow" style="justify-content:center;">
        <table class="data dense">
          <thead><tr><th>Modalidade</th><th>Qtd. FIN</th><th>Qtd. OGU</th><th>Valor FIN (R$ mi)</th><th>Valor OGU (R$ mi)</th></tr></thead>
          <tbody>
{rows_s6}
{tr_s6}
        </tbody>
        </table>
      </div>
      <div class="take"><span class="tick"></span><p>{N_NOV} propostas e R$ {V_NOV_BI} bilhões desde 2024.</p></div>
    </div>
{foot('Fonte: MCid · Novo PAC · 18/08/2026 · Cinco propostas de Mobilidade combinam OGU e FIN e ficam fora das colunas por fonte')}  </section>

  <!-- ============ 07 · PROPOSTAS ENQUADRADAS ============ -->
  <section class="p2 content" data-label="Enquadradas">
{CHROME_TOP}    <div class="body">
      <div class="title-block">
        <h2 class="title">Propostas enquadradas</h2>
        <p class="dek">Operações de financiamento habilitadas, aguardando seleção.</p>
      </div>
      <div class="row grow" style="align-items:stretch;">
        <div class="sidepanel">
          <div class="label">Enquadradas FIN</div>
          <div class="bignum">{N_ENQ}</div>
          <div class="lbl">Propostas</div>
          <div class="stat2"><span style="font-size:.55em; font-weight:700; color:var(--ink-soft);">R$ </span>{V_ENQ_BI}<span style="font-size:.55em; font-weight:700; color:var(--ink-soft);"> bi</span></div>
          <div class="lbl">Valor FIN</div>
        </div>
        <div class="col grow" style="justify-content:center;">
          <table class="data">
            <thead><tr><th>Modalidade</th><th>Qtd. FIN</th><th>Valor FIN (R$ mi)</th><th class="bh" style="padding-left:24px;">Proporção do valor</th></tr></thead>
            <tbody>
{rows_s7}
{tr_s7}
            </tbody>
          </table>
        </div>
      </div>
      <div class="take"><span class="tick"></span><p>As enquadradas cresceram de 319 para {N_ENQ} propostas desde o balanço anterior — R$ {V_ENQ_BI} bilhões prontos para avançar à seleção.</p></div>
    </div>
{foot()}  </section>

  <!-- ============ 08 · GOVERNADORES ============ -->
  <section class="p2 content" data-label="Governadores">
{CHROME_TOP}    <div class="body">
      <div class="title-block">
        <h2 class="title">Governadores</h2>
        <p class="dek">Seleções pactuadas diretamente com os governos estaduais.</p>
      </div>
      <div class="row grow" style="align-items:stretch;">
        <div class="sidepanel">
          <div class="label">Carteira</div>
          <div class="bignum">{N_GOV}</div>
          <div class="lbl">Propostas</div>
          <div class="stat2"><span style="font-size:.55em; font-weight:700; color:var(--ink-soft);">R$ </span>{V_GOV_BI}<span style="font-size:.55em; font-weight:700; color:var(--ink-soft);"> bi</span></div>
          <div class="lbl">Valor total</div>
        </div>
        <div class="col grow" style="justify-content:center;">
          <table class="data">
            <thead><tr><th>Modalidade</th><th>Qtd. FIN</th><th>Qtd. OGU</th><th>Valor FIN (R$ mi)</th><th>Valor OGU (R$ mi)</th></tr></thead>
            <tbody>
{rows_s8}
{tr_s8}
            </tbody>
          </table>
          <div class="pills" style="margin-top:38px;">
            <span class="pill">AC</span><span class="pill">BA</span><span class="pill">DF</span><span class="pill">MA</span><span class="pill">RJ</span><span class="pill">SP</span>
          </div>
        </div>
      </div>
      <div class="take"><span class="tick"></span><p>R$ {V_GOV_BI} bilhões pactuados com seis unidades da federação, com foco em mobilidade urbana.</p></div>
    </div>
{foot()}  </section>

  <!-- ============ 09 · CALAMIDADE RS ============ -->
  <section class="p2 content" data-label="Calamidade RS">
{CHROME_TOP}    <div class="body">
      <div class="title-block">
        <h2 class="title">Calamidade — Rio Grande do Sul</h2>
        <p class="dek">Seleções dedicadas à reconstrução após a calamidade de 2024. Grupo novo na carteira, sem equivalente no balanço anterior.</p>
      </div>
      <div class="trio grow">
        <div>
          <div class="tl">Propostas</div>
          <div class="tv">{N_CAL}</div>
          <div class="td">{ptint(cal.qOGU.sum())} via OGU e {ptint(cal.qFIN.sum())} via financiamento, todas no Rio Grande do Sul.</div>
        </div>
        <div>
          <div class="tl">Drenagem Urbana</div>
          <div class="tv" style="font-size:84px;">R$ {ptbr(cal_dre.tot, 0)}<span style="font-size:.4em; color:var(--ink-soft);"> mi</span></div>
          <div class="td">{ptint(cal_dre.n)} propostas de drenagem e contenção de cheias.</div>
        </div>
        <div>
          <div class="tl">MCMV — Habitação</div>
          <div class="tv" style="font-size:84px;">R$ {ptbr(cal_mcmv.tot, 1)}<span style="font-size:.4em; color:var(--ink-soft);"> mi</span></div>
          <div class="td">{ptint(cal_mcmv.n)} proposta de provisão habitacional (MCMV FNHIS).</div>
        </div>
      </div>
      <div class="take"><span class="tick"></span><p>R$ {V_CAL} milhões dedicados à reconstrução do Rio Grande do Sul.</p></div>
    </div>
{foot()}  </section>

  <!-- ============ 10 · SELEÇÕES 2026 ============ -->
  <section class="p2 content" data-label="Seleções 2026">
{CHROME_TOP}    <div class="body">
      <div class="title-row">
        <div class="title-block">
          <h2 class="title">Seleções de 2026</h2>
          <p class="dek">Propostas que entraram na carteira em 2026 — todas via financiamento.</p>
        </div>
        {LEG_SEL_ENQ}
      </div>
      <div class="row grow" style="align-items:stretch;">
        <div class="sidepanel">
          <div class="label">Em 2026</div>
          <div class="bignum">{N_26}</div>
          <div class="lbl">Propostas</div>
          <div class="stat2"><span style="font-size:.55em; font-weight:700; color:var(--ink-soft);">R$ </span>{V_26_BI}<span style="font-size:.55em; font-weight:700; color:var(--ink-soft);"> bi</span></div>
          <div class="lbl">Valor FIN</div>
        </div>
        <div class="col grow" style="justify-content:center;">
          <table class="data dense">
            <thead><tr><th>Modalidade</th><th>Selecionadas</th><th>Enquadradas</th><th>Valor (R$ mi)</th><th class="bh" style="padding-left:24px;">Por status</th></tr></thead>
            <tbody>
{rows_s10}
{tr_s10}
            </tbody>
          </table>
        </div>
      </div>
      <div class="take"><span class="tick"></span><p>Em 2026, {n26_sel} propostas já selecionadas (R$ {ptbr(v26_sel/1000,1)} bi) e {n26_enq} enquadradas (R$ {ptbr(v26_enq/1000,1)} bi) reforçam a carteira de financiamento.</p></div>
    </div>
{foot()}  </section>

  <!-- ============ 11 · ENCERRAMENTO ============ -->
  <section class="p2 dark" data-label="Encerramento">
    <div class="top">
      <div class="eyebrow">Novo PAC — MCid · Investimentos</div>
      <div class="brand"><span class="govbr">gov<b>.br</b></span><span class="div"></span><span class="min">Ministério das Cidades</span></div>
    </div>
    <hr class="hr">
    <div class="body" style="justify-content:center; padding-top:0;">
      <h2 style="font-size:120px; font-weight:800; letter-spacing:-.03em; margin:0; line-height:.95;">Obrigado.</h2>
      <p style="font-family:'Raleway',sans-serif; font-style:italic; font-weight:500; font-size:38px; line-height:1.42; color:var(--paper); margin:140px 0 0; max-width:42ch;">Novo PAC — MCid: {N_ALL} propostas e R$ {V_ALL_BI} bilhões em investimentos selecionados em todo o país.</p>
    </div>
    <hr class="hr">
    <div class="bot" style="border-top:0; padding-top:18px;">
      <span class="org" style="color:var(--paper);">Ministério das Cidades</span>
      <div style="flex: 1;"></div>
      <span class="org" style="color:var(--paper);">Agosto de 2026</span>
    </div>
  </section>

</deck-stage>
"""

html = (HEAD + DECK_CSS + EXTRA_CSS + '\n</head>\n<body>\n'
        + SLIDES
        + '\n<script>\n' + DECK_JS + '\n</script>\n'
        + DECK_TAIL + '\n</body></html>\n')

open(OUT, 'w', encoding='utf-8').write(html)
print('escrito:', OUT, f'({len(html)/1024:.0f} KB)')
print('conferência: total', N_ALL, '/', V_ALL, 'mi | mig', N_MIG, '| novas', N_NOV, '| enq', N_ENQ,
      '| gov', N_GOV, '| cal', N_CAL, '| 2026', N_26)
