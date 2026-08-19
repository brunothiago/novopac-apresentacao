#!/usr/bin/env python3
"""Gera index.html — apresentação "Novo PAC — MCid · Investimentos" (versão simples).

Espelha os agrupamentos da apresentação original (tabelas por modalidade), acrescentando:
- data de atualização e legenda de contagem em cada tabela;
- seletor de UF que refiltra todas as tabelas (também aceita ?uf=XX na URL);
- botões de status (Selecionadas / Enquadradas) nas tabelas que somam os dois
  status (também aceita ?status=sel ou ?status=enq na URL).

Exclui as 2 propostas MCMV (FNHIS / FNHIS SUB50) da contagem.

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
PDF_NAME = 'novopac-mcid-investimentos.pdf'
DATA_ATUALIZACAO = '18/08/2026'

LABELS = {
    'Médias e Grandes Cidades': 'Mobilidade: Médias e Grandes Cidades',
}
MCMV = ('MCMV FNHIS', 'MCMV FNHIS SUB50')  # fora da contagem, a pedido
MODS = [
    'Abastecimento de Água - Rural',
    'Abastecimento de Água - Urbano',
    'Contenção de Encostas',
    'Drenagem Urbana',
    'Esgotamento Sanitário',
    'Mobilidade: Médias e Grandes Cidades',
    'Regularização Fundiária',
    'Renovação de Frota',
    'Resíduos Sólidos',
    'Urbanização de Favelas',
]
FONTES = {'FIN': 0, 'OGU': 1, 'OGU/FIN': 2}

# ---------------------------------------------------------------- carga
x = pd.read_excel(XLSX, header=1)
x = x[~x.modalidade.isin(MCMV)]
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

# 'BR' (abrangência nacional) contaria no Brasil, mas nunca é opção de estado
ufs = sorted({r[2] for r in rows} - {'BR'})
uf_opts = '<option value="">Brasil — todas as UFs</option>' + ''.join(
    f'<option value="{u}">{u}</option>' for u in ufs)

print(f'linhas: {len(rows)} (migradas {len(mig)}, planilha {len(x)}) | UFs: {len(ufs)}')

# ---------------------------------------------------------------- assets
DECK_CSS = open(f'{HERE}/assets/deck.css', encoding='utf-8').read()
DECK_JS = open(f'{HERE}/assets/deck-stage.js', encoding='utf-8').read().replace('</script', '<\\/script')
DECK_CHROME = open(f'{HERE}/assets/deck-chrome.html', encoding='utf-8').read()

EXTRA_CSS = """
  body{ margin:0; background:var(--navy); }
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
  /* controles de filtro (UF + status) */
  .ctrls{ display:flex; flex-direction:column; align-items:flex-end; gap:12px; }
  .ufctl, .stctl{ display:flex; align-items:center; gap:14px; }
  .ufctl label, .stctl label{ font-size:21px; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:var(--ink-soft); }
  .ufctl select{ font-family:'rawline','Raleway',sans-serif; font-size:23px; font-weight:600; color:var(--navy); background:#fff; border:1px solid var(--line-strong); border-radius:2px; padding:10px 16px; min-width:300px; cursor:pointer; }
  .ufctl select:focus{ outline:2px solid var(--blue); outline-offset:1px; }
  .stgrp{ display:flex; border:1px solid var(--line-strong); border-radius:2px; overflow:hidden; }
  .stbtn{ font-family:'rawline','Raleway',sans-serif; font-size:21px; font-weight:700; letter-spacing:.02em; padding:11px 22px; background:#fff; color:var(--ink-soft); border:0; cursor:pointer; }
  .stbtn + .stbtn{ border-left:1px solid var(--line-strong); }
  .stbtn[aria-pressed="true"]{ background:var(--blue); color:#fff; }
  .stbtn:focus-visible{ outline:2px solid var(--blue); outline-offset:-2px; }
  .title-row{ display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:26px; }
  .title-row .title-block{ margin-bottom:0; }
  @media print{ .ufctl select{ border:0; padding:0; } }
  /* botões do rodapé de encerramento: exportar PDF e baixar a base em Excel */
  .xlsxlink{ display:inline-flex; align-items:center; gap:10px; margin-left:24px;
    font-family:'rawline','Raleway',sans-serif; font-size:var(--t-foot); font-weight:700;
    letter-spacing:.08em; text-transform:uppercase; color:var(--amber-bright);
    text-decoration:none; background:transparent; cursor:pointer;
    border:1px solid rgba(255,255,255,.32); border-radius:2px; padding:9px 18px;
    transition:background .15s, border-color .15s; }
  .xlsxlink:hover{ background:rgba(255,255,255,.10); border-color:var(--amber-bright); }
  .xlsxlink svg{ width:20px; height:20px; fill:none; stroke:currentColor; stroke-width:2; stroke-linecap:round; stroke-linejoin:round; }
  @media print{ .xlsxlink, .rotate-hint{ display:none !important; } }
  /* celular/tablet: navegação só por toque (metades esquerda/direita da tela);
     a barra inferior sai para não cobrir o conteúdo dos slides */
  @media (hover: none), (pointer: coarse){ .deck-bar-zone{ display:none !important; } }
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


def slide(label, titulo, tabela_id, legenda, colunas, dek='', dense=False, status=False):
    ths = ''.join(f'<th>{c}</th>' for c in colunas)
    dek_html = f'\n        <p class="dek" style="max-width:80ch; font-style:italic;">{dek}</p>' if dek else ''
    stctl = ('\n          <div class="stctl"><label>Status</label><div class="stgrp">'
             '<button type="button" class="stbtn" data-st="sel" aria-pressed="true">Selecionadas</button>'
             '<button type="button" class="stbtn" data-st="enq" aria-pressed="true">Enquadradas</button>'
             '</div></div>') if status else ''
    stmeta = '<span class="sep">·</span>Status: <b class="stname">selecionadas + enquadradas</b>' if status else ''
    return f"""
  <section class="p2 content" data-label="{label}">
{CHROME_TOP}    <div class="body tight">
      <div class="title-row">
        <div class="title-block tight">
          <h2 class="title">{titulo}</h2>{dek_html}
        </div>
        <div class="ctrls">
          <div class="ufctl"><label for="{tabela_id}-uf">Recorte</label><select id="{tabela_id}-uf" class="ufsel">{uf_opts}</select></div>{stctl}
        </div>
      </div>
      <div class="col grow" style="justify-content:center;">
        <p class="tmeta">Atualizado em <b>{DATA_ATUALIZACAO}</b><span class="sep">·</span>{legenda}{stmeta}<span class="sep">·</span>Recorte: <b class="ufname">Brasil</b></p>
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
       'Contagem: migradas + novas seleções',
       ['Total de propostas', 'Valor total (R$ mi)'], dense=True, status=True)}
{slide('Por quantidade', 'Investimentos por quantidade', 't-qtd',
       'Contagem: migradas + novas seleções',
       ['Qtd. FIN', 'Qtd. OGU', 'Total'], dense=True, status=True)}
{slide('Por fonte', 'Investimento por fonte', 't-fonte',
       'Contagem: migradas + novas seleções',
       ['Qtd. FIN', 'Qtd. OGU', 'Total', 'Valor OGU (R$ mi)', 'Valor FIN (R$ mi)', 'Valor total (R$ mi)'], dense=True, status=True)}
{slide('Novo PAC Migrado', 'Novo PAC Migrado', 't-mig',
       'Contagem: somente carteira migrada',
       ['Qtd. FIN', 'Qtd. OGU', 'Valor FIN (R$ mi)', 'Valor OGU (R$ mi)'],
       dek='(Valores: Empenho de OGU e Pago de FIN) &gt; dez/2022')}
{slide('Novas Seleções', 'Novas seleções', 't-novas',
       'Contagem: novas seleções, sem migradas',
       ['Qtd. FIN', 'Qtd. OGU', 'Valor FIN (R$ mi)', 'Valor OGU (R$ mi)'],
       dek='2024, 2025 e 2026 (sem migradas)', dense=True, status=True)}
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
      <a class="xlsxlink" href="{PDF_NAME}" download title="Baixar o PDF da apresentação (um slide por página)">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 21V3h9l5 5v13z"></path><path d="M14 3v6h6"></path></svg>
        Exportar PDF
      </a>
      <a class="xlsxlink" href="data/{os.path.basename(XLSX)}" download title="Baixar a base de dados em Excel">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12"></path><path d="M7 10l5 5 5-5"></path><path d="M4 19h16"></path></svg>
        XLSX Base
      </a>
    </div>
  </section>

</deck-stage>
"""

APP_JS = """
// ===== dados e filtros (UF e status) =====
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
// status: true = a tabela soma selecionadas e enquadradas e obedece aos botões de status
const TABLES = [
  { id: 't-totais', ds: 'all',  status: true,  cols: ['n', 'tot'] },
  { id: 't-qtd',    ds: 'all',  status: true,  cols: ['qf', 'qo', 'n'] },
  { id: 't-fonte',  ds: 'all',  status: true,  cols: ['qf', 'qo', 'n', 'ogu', 'fin', 'tot'] },
  { id: 't-mig',    ds: 'mig',  status: false, cols: ['qf', 'qo', 'fin', 'ogu'] },
  { id: 't-novas',  ds: 'plan', status: true,  cols: ['qf', 'qo', 'fin', 'ogu'] },
  { id: 't-enq',    ds: 'enq',  status: false, cols: ['qf', 'fin'] },
  { id: 't-gov',    ds: 'gov',  status: false, cols: ['qf', 'qo', 'fin', 'ogu'] },
];

const state = { uf: '', sel: true, enq: true };
const isSel = (r) => r[3] <= 1; // migradas (kind 0) contam como selecionadas

function aggregate(t) {
  const keep = DATASETS[t.ds];
  const acc = MODS.map(() => ({ qf: 0, qo: 0, n: 0, ogu: 0, fin: 0, tot: 0 }));
  for (const r of ROWS) {
    if (!keep(r)) continue;
    if (state.uf && r[2] !== state.uf) continue;
    if (t.status && !(isSel(r) ? state.sel : state.enq)) continue;
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

function render() {
  for (const t of TABLES) {
    const acc = aggregate(t);
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
  const nome = state.uf || 'Brasil';
  const stname = state.sel && state.enq ? 'selecionadas + enquadradas'
    : (state.sel ? 'somente selecionadas' : 'somente enquadradas');
  document.querySelectorAll('.ufname').forEach((el) => { el.textContent = nome; });
  document.querySelectorAll('.stname').forEach((el) => { el.textContent = stname; });
  document.querySelectorAll('select.ufsel').forEach((el) => { el.value = state.uf; });
  document.querySelectorAll('.stbtn').forEach((el) => {
    el.setAttribute('aria-pressed', String(state[el.dataset.st]));
  });
}

(function () {
  const params = new URLSearchParams(location.search);
  let uf = (params.get('uf') || '').toUpperCase().trim();
  if (uf && !ROWS.some((r) => r[2] === uf)) uf = '';
  state.uf = uf;
  const st = (params.get('status') || '').toLowerCase().trim();
  if (st === 'sel') state.enq = false; else if (st === 'enq') state.sel = false;
  render();
  document.querySelectorAll('select.ufsel').forEach((el) => {
    el.addEventListener('change', function () { state.uf = this.value; render(); });
  });
  document.querySelectorAll('.stbtn').forEach((el) => {
    el.addEventListener('click', function () {
      const k = this.dataset.st, other = k === 'sel' ? 'enq' : 'sel';
      if (state[k] && !state[other]) return; // pelo menos um status ativo
      state[k] = !state[k];
      render();
    });
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

# ---------------- PDF pré-gerado (um slide por página, paisagem 1920×1080) ----------------
# O Safari do iPhone ignora o @page do CSS de impressão; servir um PDF pronto garante o
# formato correto e abre direto no visualizador do celular, com o botão de compartilhar.
CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
pdf_out = f'{HERE}/{PDF_NAME}'
if os.path.exists(CHROME):
    import subprocess
    r = subprocess.run(
        [CHROME, '--headless=new', '--disable-gpu', '--virtual-time-budget=8000',
         '--no-pdf-header-footer', f'--print-to-pdf={pdf_out}', f'file://{OUT}'],
        capture_output=True, timeout=180)
    if r.returncode == 0 and os.path.exists(pdf_out):
        print('PDF gerado:', pdf_out, f'({os.path.getsize(pdf_out)//1024} KB)')
    else:
        print('AVISO: geração do PDF falhou (código', r.returncode, ') — link do slide ficará quebrado até regenerar')
else:
    aviso = 'mantido o PDF anterior' if os.path.exists(pdf_out) else 'PDF não gerado; link ficará quebrado'
    print(f'AVISO: Chrome não encontrado — {aviso}')
