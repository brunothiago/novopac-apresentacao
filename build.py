#!/usr/bin/env python3
"""Gera index.html — apresentação "Novo PAC — MCid · Investimentos" (versão simples).

Espelha os agrupamentos da apresentação original (tabelas por modalidade), acrescentando:
- data de atualização e legenda de contagem em cada tabela;
- seletor de UF que refiltra todas as tabelas (também aceita ?uf=XX na URL).

Fontes (em data/):
- base_completa_atualizada_20260818_1126.xlsx — seleções 2024-2026 (selecionadas + enquadradas)
- view_sis_novopac_previsto_unificado_202608180817.csv — migradas (origem_dado == "Novo PAC - Retomada")

Convenção (igual à apresentação original): totais = migradas + selecionadas + enquadradas FIN.
"""
import json
import os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = f'{HERE}/data/base_completa_atualizada_20260818_1126.xlsx'
CSV = f'{HERE}/data/view_sis_novopac_previsto_unificado_202608180817.csv'
OUT = f'{HERE}/index.html'
DATA_ATUALIZACAO = '18/08/2026'

LABELS = {
    'Médias e Grandes Cidades': 'Mobilidade: Médias e Grandes Cidades',
    'MCMV FNHIS': 'MCMV (Calamidade RS)',
    'MCMV FNHIS SUB50': 'MCMV (Calamidade RS)',
}
MODS = [
    'Abastecimento de Água - Rural',
    'Abastecimento de Água - Urbano',
    'Contenção de Encostas',
    'Drenagem Urbana',
    'Esgotamento Sanitário',
    'MCMV (Calamidade RS)',
    'Mobilidade: Médias e Grandes Cidades',
    'Regularização Fundiária',
    'Renovação de Frota',
    'Resíduos Sólidos',
    'Urbanização de Favelas',
]
FONTES = {'FIN': 0, 'OGU': 1, 'OGU/FIN': 2}

# ---------------------------------------------------------------- carga
x = pd.read_excel(XLSX, header=1)
v = pd.read_csv(CSV)
mig = v[v.origem_dado == 'Novo PAC - Retomada'].copy()
for d in (x, mig):
    for c in ('vlr_portaria_ogu', 'vlr_portaria_fin'):
        d[c] = pd.to_numeric(d[c], errors='coerce').fillna(0)
    d['mod'] = d.modalidade.replace(LABELS)
    d['uf'] = d.uf.astype(str).str.strip().str.upper()

# linhas compactas: [modIdx, fonteIdx, uf, kind, gov, ogu, fin]
# kind: 0 = migrada · 1 = selecionada · 2 = enquadrada  |  gov: 1 = grupo Governadores
rows = []
for _, r in mig.iterrows():
    rows.append([MODS.index(r['mod']), FONTES[r.fonte], r.uf, 0, 0,
                 round(float(r.vlr_portaria_ogu), 2), round(float(r.vlr_portaria_fin), 2)])
for _, r in x.iterrows():
    kind = 2 if r.status_selecao == 'enquadrada' else 1
    gov = 1 if r.grupo_modalidade == 'Governadores' else 0
    rows.append([MODS.index(r['mod']), FONTES[r.fonte], r.uf, kind, gov,
                 round(float(r.vlr_portaria_ogu), 2), round(float(r.vlr_portaria_fin), 2)])

# 'BR' marca 1 proposta MCMV de abrangência nacional: conta no Brasil, mas não é opção de estado
ufs = sorted({r[2] for r in rows} - {'BR'})
uf_opts = '<option value="">Brasil — todas as UFs</option>' + ''.join(
    f'<option value="{u}">{u}</option>' for u in ufs)

print(f'linhas: {len(rows)} (migradas {len(mig)}, planilha {len(x)}) | UFs: {len(ufs)}')

# ---------------------------------------------------------------- assets
DECK_CSS = open(f'{HERE}/assets/deck.css', encoding='utf-8').read()
DECK_JS = open(f'{HERE}/assets/deck-stage.js', encoding='utf-8').read().replace('</script', '<\\/script')
DECK_CHROME = open(f'{HERE}/assets/deck-chrome.html', encoding='utf-8').read()

EXTRA_CSS = """
  /* ---- tabelas de dados (versão simples) ---- */
  table.data{ width:100%; border-collapse:collapse; }
  table.data th{ text-align:right; font-size:21px; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--ink-soft); padding:0 22px 13px; border-bottom:2px solid var(--line-strong); white-space:nowrap; }
  table.data th:first-child{ text-align:left; padding-left:0; }
  table.data td{ padding:11px 22px; border-bottom:1px solid var(--line); font-size:25px; color:var(--ink); font-variant-numeric:tabular-nums; text-align:right; white-space:nowrap; vertical-align:middle; }
  table.data td:first-child{ text-align:left; padding-left:0; font-weight:600; }
  table.data tr.tot td{ border-top:2px solid var(--line-strong); border-bottom:0; font-weight:800; color:var(--navy); }
  table.data td.empty{ text-align:center; font-weight:400; color:var(--ink-soft); padding:60px 0; }
  table.data.dense th{ font-size:19px; padding:0 18px 12px; }
  table.data.dense td{ padding:8px 18px; font-size:23px; }
  .body.tight{ padding-top:26px; }
  .title-block.tight{ margin-bottom:22px; }
  /* linha de metadados da tabela: atualização · contagem · recorte */
  .tmeta{ font-size:21px; color:var(--ink-soft); margin:0 0 16px; }
  .tmeta b{ color:var(--navy); font-weight:700; }
  .tmeta .sep{ margin:0 12px; opacity:.5; }
  /* seletor de UF */
  .ufctl{ display:flex; align-items:center; gap:14px; }
  .ufctl label{ font-size:21px; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:var(--ink-soft); }
  .ufctl select{ font-family:'rawline','Raleway',sans-serif; font-size:23px; font-weight:600; color:var(--navy); background:#fff; border:1px solid var(--line-strong); border-radius:2px; padding:10px 16px; min-width:300px; cursor:pointer; }
  .ufctl select:focus{ outline:2px solid var(--blue); outline-offset:1px; }
  .title-row{ display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:26px; }
  .title-row .title-block{ margin-bottom:0; }
  @media print{ .ufctl select{ border:0; padding:0; } }
"""

HEAD = f"""<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Novo PAC — MCid · Investimentos</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin="">
<!-- Tipografia oficial do Governo Federal (Padrão Digital de Governo — gov.br/ds) -->
<link rel="stylesheet" href="https://cdngovbr-ds.estaleiro.serpro.gov.br/design-system/fonts/rawline/css/rawline.css">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,300;1,400;1,500;1,600;1,700&display=swap">
<style>{DECK_CSS}
{EXTRA_CSS}</style>
</head>
<body>
"""

CHROME_TOP = """    <div class="top">
      <div class="eyebrow">Novo PAC — MCid · Investimentos</div>
      <div class="brand"><span class="govbr">gov<b>.br</b></span><span class="div"></span><span class="min">Ministério das Cidades</span></div>
    </div>
    <hr class="hr">
"""

FOOT = f"""    <hr class="hr">
    <div class="bot">
      <span class="org">Fonte: MCid · Portarias de seleção do Novo PAC · atualizado em {DATA_ATUALIZACAO}</span>
      <span class="pageno"></span>
    </div>
"""


def slide(label, titulo, tabela_id, legenda, colunas, dek='', dense=False):
    ths = ''.join(f'<th>{c}</th>' for c in colunas)
    dek_html = f'\n        <p class="dek" style="max-width:80ch; font-style:italic;">{dek}</p>' if dek else ''
    return f"""
  <section class="p2 content" data-label="{label}">
{CHROME_TOP}    <div class="body tight">
      <div class="title-row">
        <div class="title-block tight">
          <h2 class="title">{titulo}</h2>{dek_html}
        </div>
        <div class="ufctl"><label for="{tabela_id}-uf">Recorte</label><select id="{tabela_id}-uf" class="ufsel">{uf_opts}</select></div>
      </div>
      <div class="col grow" style="justify-content:center;">
        <p class="tmeta">Atualizado em <b>{DATA_ATUALIZACAO}</b><span class="sep">·</span>{legenda}<span class="sep">·</span>Recorte: <b class="ufname">Brasil</b></p>
        <table class="data{' dense' if dense else ''}" id="{tabela_id}">
          <thead><tr><th>Modalidade</th>{ths}</tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
{FOOT}  </section>
"""


SLIDES = f"""
<deck-stage width="1920" height="1080" no-rail="">

  <!-- ============ CAPA ============ -->
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
      <p style="font-family:'Raleway',sans-serif; font-style:italic; font-weight:500; font-size:38px; line-height:1.4; color:var(--paper); margin:190px 0 0; max-width:44ch;">Balanço da carteira de seleções — migradas, novas seleções e propostas enquadradas.</p>
    </div>
    <hr class="hr">
    <div class="bot" style="border-top:0; padding-top:18px;">
      <span class="org" style="color:var(--paper);">Ministério das Cidades</span>
      <div style="flex: 1;"></div>
      <span class="org" style="color:var(--paper);">Atualizado em {DATA_ATUALIZACAO}</span>
    </div>
  </section>

{slide('Investimentos Totais', 'Investimentos totais', 't-totais',
       'Contagem: migradas + selecionadas + enquadradas (FIN)',
       ['Total de propostas', 'Valor total (R$ mi)'], dense=True)}
{slide('Por quantidade', 'Investimentos por quantidade', 't-qtd',
       'Contagem: migradas + selecionadas + enquadradas (FIN)',
       ['Qtd. FIN', 'Qtd. OGU', 'Total'], dense=True)}
{slide('Por fonte', 'Investimento por fonte', 't-fonte',
       'Contagem: migradas + selecionadas + enquadradas (FIN)',
       ['Qtd. FIN', 'Qtd. OGU', 'Total', 'Valor OGU (R$ mi)', 'Valor FIN (R$ mi)', 'Valor total (R$ mi)'], dense=True)}
{slide('Novo PAC Migrado', 'Novo PAC Migrado', 't-mig',
       'Contagem: somente carteira migrada',
       ['Qtd. FIN', 'Qtd. OGU', 'Valor FIN (R$ mi)', 'Valor OGU (R$ mi)'],
       dek='(Valores: Empenho de OGU e Pago de FIN) &gt; dez/2022')}
{slide('Novas Seleções', 'Novas seleções', 't-novas',
       'Contagem: selecionadas + enquadradas (FIN), sem migradas',
       ['Qtd. FIN', 'Qtd. OGU', 'Valor FIN (R$ mi)', 'Valor OGU (R$ mi)'],
       dek='2024, 2025 e 2026 (sem migradas)', dense=True)}
{slide('Enquadradas', 'Novo PAC — propostas enquadradas', 't-enq',
       'Contagem: somente propostas enquadradas (FIN)',
       ['Qtd. FIN', 'Valor FIN (R$ mi)'])}
{slide('Governadores', 'Governadores', 't-gov',
       'Contagem: seleções do grupo Governadores',
       ['Qtd. FIN', 'Qtd. OGU', 'Valor FIN (R$ mi)', 'Valor OGU (R$ mi)'])}

  <!-- ============ ENCERRAMENTO ============ -->
  <section class="p2 dark" data-label="Encerramento">
    <div class="top">
      <div class="eyebrow">Novo PAC — MCid · Investimentos</div>
      <div class="brand"><span class="govbr">gov<b>.br</b></span><span class="div"></span><span class="min">Ministério das Cidades</span></div>
    </div>
    <hr class="hr">
    <div class="body" style="justify-content:center; padding-top:0;">
      <h2 style="font-size:120px; font-weight:800; letter-spacing:-.03em; margin:0; line-height:.95;">Obrigado.</h2>
      <p style="font-family:'Raleway',sans-serif; font-style:italic; font-weight:500; font-size:38px; line-height:1.42; color:var(--paper); margin:140px 0 0; max-width:42ch;">Novo PAC — Ministério das Cidades.</p>
    </div>
    <hr class="hr">
    <div class="bot" style="border-top:0; padding-top:18px;">
      <span class="org" style="color:var(--paper);">Ministério das Cidades</span>
      <div style="flex: 1;"></div>
      <span class="org" style="color:var(--paper);">Atualizado em {DATA_ATUALIZACAO}</span>
    </div>
  </section>

</deck-stage>
"""

APP_JS = """
// ===== dados e filtro por UF =====
// linha: [modIdx, fonteIdx, uf, kind, gov, vlrOGU, vlrFIN]
// fonte: 0=FIN 1=OGU 2=OGU/FIN · kind: 0=migrada 1=selecionada 2=enquadrada · gov: grupo Governadores
const MODS = __MODS__;
const ROWS = __ROWS__;

const fmtInt = new Intl.NumberFormat('pt-BR');
const fmtMi = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const mi = (v) => fmtMi.format(v / 1e6);

const DATASETS = {
  all:   (r) => true,
  mig:   (r) => r[3] === 0,
  plan:  (r) => r[3] >= 1,
  enq:   (r) => r[3] === 2,
  gov:   (r) => r[4] === 1,
};
// colunas: qf/qo = contagem FIN/OGU · n = total (inclui OGU/FIN) · ogu/fin/tot = valores
const TABLES = [
  { id: 't-totais', ds: 'all',  cols: ['n', 'tot'] },
  { id: 't-qtd',    ds: 'all',  cols: ['qf', 'qo', 'n'] },
  { id: 't-fonte',  ds: 'all',  cols: ['qf', 'qo', 'n', 'ogu', 'fin', 'tot'] },
  { id: 't-mig',    ds: 'mig',  cols: ['qf', 'qo', 'fin', 'ogu'] },
  { id: 't-novas',  ds: 'plan', cols: ['qf', 'qo', 'fin', 'ogu'] },
  { id: 't-enq',    ds: 'enq',  cols: ['qf', 'fin'] },
  { id: 't-gov',    ds: 'gov',  cols: ['qf', 'qo', 'fin', 'ogu'] },
];

function aggregate(ds, uf) {
  const keep = DATASETS[ds];
  const acc = MODS.map(() => ({ qf: 0, qo: 0, n: 0, ogu: 0, fin: 0, tot: 0 }));
  for (const r of ROWS) {
    if (!keep(r)) continue;
    if (uf && r[2] !== uf) continue;
    const a = acc[r[0]];
    a.n += 1;
    if (r[1] === 0) a.qf += 1; else if (r[1] === 1) a.qo += 1;
    a.ogu += r[5]; a.fin += r[6]; a.tot += r[5] + r[6];
  }
  return acc;
}

function cell(col, a) {
  const v = (col === 'qf' || col === 'qo' || col === 'n') ? fmtInt.format(a[col]) : mi(a[col]);
  return '<td>' + v + '</td>';
}

function render(uf) {
  for (const t of TABLES) {
    const acc = aggregate(t.ds, uf);
    const body = [];
    const total = { qf: 0, qo: 0, n: 0, ogu: 0, fin: 0, tot: 0 };
    acc.forEach((a, i) => {
      if (!a.n) return;
      for (const k in total) total[k] += a[k];
      body.push('<tr><td>' + MODS[i] + '</td>' + t.cols.map((c) => cell(c, a)).join('') + '</tr>');
    });
    const tb = document.querySelector('#' + t.id + ' tbody');
    if (!body.length) {
      tb.innerHTML = '<tr><td class="empty" colspan="' + (t.cols.length + 1) + '">Sem propostas neste recorte.</td></tr>';
    } else {
      body.push('<tr class="tot"><td>TOTAL</td>' + t.cols.map((c) => cell(c, total)).join('') + '</tr>');
      tb.innerHTML = body.join('');
    }
  }
  const nome = uf || 'Brasil';
  document.querySelectorAll('.ufname').forEach((el) => { el.textContent = nome; });
  document.querySelectorAll('select.ufsel').forEach((el) => { el.value = uf; });
}

(function () {
  let uf = (new URLSearchParams(location.search).get('uf') || '').toUpperCase().trim();
  if (uf && !ROWS.some((r) => r[2] === uf)) uf = '';
  render(uf);
  document.querySelectorAll('select.ufsel').forEach((el) => {
    el.addEventListener('change', function () { render(this.value); });
  });
})();
"""

app_js = APP_JS.replace('__MODS__', json.dumps(MODS, ensure_ascii=False)) \
               .replace('__ROWS__', json.dumps(rows, ensure_ascii=False, separators=(',', ':')))

html = (HEAD + SLIDES
        + '\n<script>\n' + DECK_JS + '\n</script>\n'
        + '<script>\n' + app_js + '\n</script>\n'
        + DECK_CHROME + '\n</body></html>\n')

open(OUT, 'w', encoding='utf-8').write(html)
print('escrito:', OUT, f'({len(html)/1024:.0f} KB)')
