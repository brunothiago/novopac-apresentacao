#!/usr/bin/env python3
"""Gera index.html — apresentação "Novo PAC — MCid · Investimentos" (versão simples).

Espelha os agrupamentos da apresentação original (tabelas por modalidade), acrescentando:
- data de atualização e legenda de contagem em cada tabela;
- seletor de UF que refiltra todas as tabelas (também aceita ?uf=XX na URL);
- botões de status (Selecionadas / Enquadradas) nas tabelas que somam os dois
  status (também aceita ?status=sel ou ?status=enq na URL).

Exclui as 2 propostas MCMV (FNHIS / FNHIS SUB50) da contagem.

Fonte: foto da tabela se_cgpac.tab_base_unica_gm em data/base_unica_gm_<AAAAMMDD>.xlsx
(gerada por python/extrair_base_unica.py; carga e regras em dados.py):
- migradas: status_selecao == "retomada"
- seleções 2024-2026: status_selecao "selecionada" ou "enquadrada"

Convenção (igual à apresentação original): totais = migradas + selecionadas + enquadradas FIN.
"""
import json
import os
import pandas as pd

import dados

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/index.html'
PDF_NAME = 'novopac-mcid-investimentos.pdf'
VERSAO = '1.17'  # subir a cada commit que altere a apresentação
# FIRECE — financiamento da mesma temática FORA do escopo do Novo PAC.
# Valor fixo: atualizar somente quando o usuário indicar (e subir a VERSAO).
FIRECE_VALOR = 'R$ 6,5 bi'
FIRECE_DESC = 'Reconstrução e adaptação após os eventos extremos no RS.'
FIRECE_LINK = 'https://www.in.gov.br/web/dou/-/resolucao-n-1-comite-gestor-do-fundo/cc-de-13-de-dezembro-de-2024-633793584'
# histórico exibido no popup do "versão x.y" (encerramento) — acrescentar uma linha a cada versão
HISTORICO = [
    ('1.17', 'Base em Excel protegida por senha, com versão por data/hora.'),
    ('1.16', 'Dados passam a vir da base única do banco (se_cgpac).'),
    ('1.15', 'Recorte de ano passa a excluir as migradas.'),
    ('1.14', 'Exportar PDF baixa direto, na horizontal, com os filtros.'),
    ('1.13', 'Exportar PDF respeita os filtros aplicados.'),
    ('1.12', 'Status "Enquadradas/habilitadas" e ajuste do slide 7.'),
    ('1.11', 'Recorte por ano corrigido: 2023, 2025 e 2026.'),
    ('1.10', 'Histórico de versões no rodapé.'),
    ('1.9', 'Rótulos com "DMP/SE" nos slides.'),
    ('1.8', 'Seta de destaque para a fonte de dados.'),
    ('1.7', 'Símbolo "+" maior no painel FIRECE.'),
    ('1.6', 'FIRECE ajustado para R$ 6,5 bi, com link do DOU.'),
    ('1.5', 'Novo rótulo do painel FIRECE.'),
    ('1.4', 'Painel FIRECE no slide Prevenção a Desastres.'),
    ('1.3', 'Slide Prevenção a Desastres antes de Governadores.'),
    ('1.2', 'Rolagem entre slides mais ágil.'),
    ('1.1', 'Setas ↑/↓ também navegam os slides.'),
    ('1.0', 'Recorte por ano e versão no rodapé.'),
]
# anos de exibição (somente rótulo — nos dados, ano_selecao segue o ano da portaria)
ANO_LABEL = {2024: '2023', 2025: '2025', 2026: '2026'}

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
x, mig, DATA_ATUALIZACAO, XLSX = dados.load()
x = x[~x.modalidade.isin(MCMV)]
for d in (x, mig):
    for c in ('vlr_portaria_ogu', 'vlr_portaria_fin'):
        d[c] = pd.to_numeric(d[c], errors='coerce').fillna(0)
    d['mod'] = d.modalidade.replace(LABELS)
    d['uf'] = d.uf.astype(str).str.strip().str.upper()
    fora = (set(d['mod']) - set(MODS)) | (set(d.fonte) - set(FONTES))
    if fora:
        raise SystemExit(f'modalidade/fonte fora da lista (MODS/FONTES): {sorted(fora)}')

# linhas compactas: [modIdx, fonteIdx, uf, kind, gov, ogu, fin, ano]
# kind: 0 = migrada · 1 = selecionada · 2 = enquadrada  |  gov: 1 = grupo Governadores
# ano: rótulo de exibição (ANO_LABEL); '' para migradas (fora do recorte por ano)
rows = []
for _, r in mig.iterrows():
    rows.append([MODS.index(r['mod']), FONTES[r.fonte], r.uf, 0, 0,
                 round(float(r.vlr_portaria_ogu), 2), round(float(r.vlr_portaria_fin), 2), ''])
for _, r in x.iterrows():
    kind = 2 if r.status_selecao == 'enquadrada' else 1
    gov = 1 if r.grupo_modalidade == 'Governadores' else 0
    rows.append([MODS.index(r['mod']), FONTES[r.fonte], r.uf, kind, gov,
                 round(float(r.vlr_portaria_ogu), 2), round(float(r.vlr_portaria_fin), 2),
                 ANO_LABEL[int(r.ano_selecao)]])
ANOS = sorted({r[7] for r in rows if r[7]})

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
  /* botão de status com rótulo em duas linhas (Enquadradas/ + habilitadas) */
  .stbtn.st2{ padding-top:4px; padding-bottom:4px; line-height:1.08; }
  .stbtn .stsub{ display:block; }
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
  .xlsxlink .lock{ width:17px; height:17px; margin-left:-2px; opacity:.85; }
  .xlsxlink svg{ width:20px; height:20px; fill:none; stroke:currentColor; stroke-width:2; stroke-linecap:round; stroke-linejoin:round; }
  @media print{ .xlsxlink, .rotate-hint{ display:none !important; } }
  /* celular/tablet: navegação só por toque (metades esquerda/direita da tela);
     a barra inferior sai para não cobrir o conteúdo dos slides */
  @media (hover: none), (pointer: coarse){ .deck-bar-zone{ display:none !important; } }
  /* botão flutuante de tela cheia — só em telas de toque */
  .fsbtn{ display:none; position:fixed; z-index:2147483000;
    right:max(16px, env(safe-area-inset-right)); bottom:max(16px, env(safe-area-inset-bottom));
    width:54px; height:54px; border-radius:50%; border:1px solid rgba(255,255,255,.28);
    background:rgba(7,29,65,.55); backdrop-filter:blur(8px); cursor:pointer;
    align-items:center; justify-content:center; padding:0; }
  .fsbtn svg{ width:22px; height:22px; stroke:#fff; fill:none; stroke-width:2; stroke-linecap:round; stroke-linejoin:round; }
  .fsbtn:active{ background:rgba(19,81,180,.75); }
  @media (hover: none), (pointer: coarse){ .fsbtn{ display:flex; } }
  @media print{ .fsbtn{ display:none !important; } }
  /* painel de financiamento fora do escopo do Novo PAC (FIRECE) */
  .aside-fin{ margin-top:38px; background:var(--paper-2); border-left:5px solid var(--green); padding:24px 32px; }
  .af-label{ font-size:19px; font-weight:700; letter-spacing:.13em; text-transform:uppercase; color:var(--ink-soft); margin-bottom:12px; }
  .af-plus{ font-size:32px; font-weight:800; color:var(--green); vertical-align:-3px; margin-right:2px; }
  .af-row{ display:flex; align-items:baseline; gap:26px; flex-wrap:wrap; }
  .af-name{ font-size:33px; font-weight:800; color:var(--ink); letter-spacing:-.01em; }
  .af-val{ font-size:44px; font-weight:800; color:var(--green); font-variant-numeric:tabular-nums; letter-spacing:-.02em; }
  .af-desc{ font-size:24px; color:var(--ink-soft); }
  .af-link{ color:var(--green); font-weight:600; text-decoration:underline; text-underline-offset:3px; }
  .af-link:hover{ opacity:.8; }
  /* versão da apresentação (rodapés) */
  .ver{ font-size:18px; font-weight:600; letter-spacing:.06em; opacity:.72; margin-left:14px; white-space:nowrap; }
  /* histórico de atualizações: popup do "versão x.y" no rodapé do encerramento */
  .verwrap{ position:relative; display:inline-block; }
  .verlink{ background:none; border:0; padding:0; font-family:inherit; color:inherit; cursor:pointer; text-decoration:underline dotted; text-underline-offset:4px; }
  .verlink:hover{ opacity:1; color:var(--amber-bright); }
  .verlog{ position:absolute; bottom:calc(100% + 16px); right:0; width:640px; background:var(--navy-2); border:1px solid var(--line-dark); border-left:4px solid var(--amber-bright); padding:24px 28px; text-align:left; box-shadow:0 12px 40px rgba(0,0,0,.45); }
  .verlog h3{ margin:0 0 14px; font-size:16px; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:var(--amber-bright); }
  .verlog ul{ margin:0; padding:0; list-style:none; }
  .verlog li{ display:flex; gap:16px; font-size:18px; line-height:1.75; color:var(--paper); }
  .verlog li b{ flex:none; width:52px; color:var(--blue-soft); font-variant-numeric:tabular-nums; }
  @media print{ .verlog{ display:none !important; } }
  /* chamada com seta apontando para o link de fonte dos dados (encerramento) */
  .srchint{ display:flex; align-items:center; gap:14px; margin-bottom:26px; }
  .srchint-txt{ font-family:'Raleway',sans-serif; font-style:italic; font-weight:600; font-size:26px; color:var(--amber-bright); }
  .srchint-arrow{ width:34px; height:34px; flex:none; fill:none; stroke:var(--amber-bright); stroke-width:2.4; stroke-linecap:round; stroke-linejoin:round; animation:srchint-bounce 1.4s ease-in-out infinite; }
  @keyframes srchint-bounce{ 0%,100%{ transform:translateY(0); } 50%{ transform:translateY(8px); } }
  @media print{ .srchint-arrow{ animation:none; } }
  /* link de fonte dos dados (encerramento) */
  .srclink{ display:inline-flex; align-items:center; gap:16px; font-size:29px; font-weight:600;
    color:#fff; text-decoration:none; border:1px solid rgba(255,255,255,.32); border-radius:2px;
    padding:18px 28px; transition:background .15s, border-color .15s; }
  .srclink:hover{ background:rgba(255,255,255,.10); border-color:var(--amber-bright); }
  .srclink svg{ width:28px; height:28px; flex:none; fill:none; stroke:var(--amber-bright); stroke-width:2; stroke-linecap:round; stroke-linejoin:round; }
"""

HEAD = f"""<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<title>Novo PAC — MCid · Investimentos</title>
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Novo PAC">
<meta name="theme-color" content="#071D41">
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
      <div class="brand"><span class="govbr">gov<b>.br</b></span><span class="div"></span><span class="min">Ministério das Cidades - DMP/SE</span></div>
    </div>
    <hr class="hr">
"""

FOOT = f"""    <hr class="hr">
    <div class="bot">
      <span class="org">Fonte: MCid · Portarias de seleção do Novo PAC · atualizado em {DATA_ATUALIZACAO}</span>
      <span style="display:inline-flex; align-items:baseline;"><span class="pageno"></span><span class="ver">versão {VERSAO}</span></span>
    </div>
"""


def slide(label, titulo, tabela_id, legenda, colunas, dek='', dense=False, status=False, extra='', empilhado=False):
    ths = ''.join(f'<th>{c}</th>' for c in colunas)
    dek_html = f'\n        <p class="dek" style="max-width:80ch; font-style:italic;">{dek}</p>' if dek else ''
    stctl = ('\n          <div class="stctl"><label>Status</label><div class="stgrp">'
             '<button type="button" class="stbtn" data-st="sel" aria-pressed="true">Selecionadas</button>'
             '<button type="button" class="stbtn st2" data-st="enq" aria-pressed="true">Enquadradas/<span class="stsub">habilitadas</span></button>'
             '</div></div>') if status else ''
    anos_btns = ''.join(f'<button type="button" class="stbtn anobtn" data-ano="{a}" aria-pressed="true">{a}</button>' for a in ANOS)
    anoctl = (f'\n          <div class="stctl anoctl"><label>Ano</label><div class="stgrp">{anos_btns}</div></div>') if status else ''
    stmeta = ('<span class="sep">·</span>Status: <b class="stname">selecionadas + enquadradas/habilitadas</b>'
              '<span class="sep">·</span>Anos: <b class="anoname">todos</b>') if status else ''
    return f"""
  <section class="p2 content" data-label="{label}">
{CHROME_TOP}    <div class="body tight">
      <div class="title-row"{' style="flex-direction:column; align-items:stretch; gap:16px;"' if empilhado else ''}>
        <div class="title-block tight">
          <h2 class="title">{titulo}</h2>{dek_html}
        </div>
        <div class="ctrls"{' style="align-self:flex-end;"' if empilhado else ''}>
          <div class="ufctl"><label for="{tabela_id}-uf">Recorte</label><select id="{tabela_id}-uf" class="ufsel">{uf_opts}</select></div>{stctl}{anoctl}
        </div>
      </div>
      <div class="col grow" style="justify-content:center;">
        <p class="tmeta">Atualizado em <b>{DATA_ATUALIZACAO}</b><span class="sep">·</span>{legenda}{stmeta}<span class="sep">·</span>Recorte: <b class="ufname">Brasil</b></p>
        <table class="data{' dense' if dense else ''}" id="{tabela_id}">
          <thead><tr><th>Modalidade</th>{ths}</tr></thead>
          <tbody></tbody>
        </table>{extra}
      </div>
    </div>
{FOOT}  </section>
"""


HISTORICO_LI = '\n'.join(f'            <li><b>{v}</b>{t}</li>' for v, t in HISTORICO)

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
        <span class="min">Ministério das Cidades - DMP/SE</span>
      </div>
    </div>
    <hr class="hr">
    <div class="body" style="justify-content:center; padding-top:0;">
      <div class="label" style="color:var(--amber-bright); margin-bottom:30px;">Investimentos</div>
      <h1 style="font-size:150px; font-weight:800; line-height:.98; letter-spacing:-.03em; margin:0; max-width:20ch;">Novo PAC<span style="color:var(--blue-soft); font-weight:500;"> — MCid</span></h1>
      <p style="font-family:'Raleway',sans-serif; font-style:italic; font-weight:500; font-size:38px; line-height:1.4; color:var(--paper); margin:190px 0 0; max-width:44ch;">Balanço da carteira de seleções — migradas, novas seleções e propostas enquadradas/habilitadas.</p>
    </div>
    <hr class="hr">
    <div class="bot" style="border-top:0; padding-top:18px;">
      <span class="org" style="color:var(--paper);">Ministério das Cidades - DMP/SE</span>
      <div style="flex: 1;"></div>
      <span class="org" style="color:var(--paper);">Atualizado em {DATA_ATUALIZACAO}<span class="ver">versão {VERSAO}</span></span>
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
       dek='2023, 2025 e 2026 (sem migradas)', dense=True, status=True)}
{slide('Enquadradas/Habilitadas', 'Novo PAC — propostas enquadradas/habilitadas', 't-enq',
       'Contagem: somente propostas enquadradas/habilitadas (FIN)',
       ['Qtd. FIN', 'Valor FIN (R$ mi)'], empilhado=True)}
{slide('Prevenção a Desastres', 'Prevenção a Desastres', 't-prev',
       'Contagem: migradas + novas seleções — subeixo Prevenção a Desastres',
       ['Qtd. FIN', 'Qtd. OGU', 'Total', 'Valor OGU (R$ mi)', 'Valor FIN (R$ mi)', 'Valor total (R$ mi)'],
       dek='Drenagem Urbana e Contenção de Encostas, por fonte', status=True,
       extra=f'''
        <div class="aside-fin">
          <div class="af-label"><span class="af-plus">+</span> Investimento de Prevenção a Desastres</div>
          <div class="af-row">
            <span class="af-name">FIRECE</span>
            <span class="af-val">{FIRECE_VALOR}</span>
            <span class="af-desc">{FIRECE_DESC} <a class="af-link" href="{FIRECE_LINK}" target="_blank" rel="noopener">Conforme link DOU</a></span>
          </div>
        </div>''')}
{slide('Governadores', 'Governadores', 't-gov',
       'Contagem: seleções do grupo Governadores',
       ['Qtd. FIN', 'Qtd. OGU', 'Valor FIN (R$ mi)', 'Valor OGU (R$ mi)'])}

  <!-- ============ ENCERRAMENTO ============ -->
  <section class="p2 dark" data-label="Encerramento">
    <div class="top">
      <div class="eyebrow">Novo PAC — MCid · Investimentos</div>
      <div class="brand"><span class="govbr">gov<b>.br</b></span><span class="div"></span><span class="min">Ministério das Cidades - DMP/SE</span></div>
    </div>
    <hr class="hr">
    <div class="body" style="justify-content:center; padding-top:0;">
      <h2 style="font-size:120px; font-weight:800; letter-spacing:-.03em; margin:0; line-height:.95;">Obrigado.</h2>
      <p style="font-family:'Raleway',sans-serif; font-style:italic; font-weight:500; font-size:38px; line-height:1.42; color:var(--paper); margin:100px 0 0; max-width:42ch;">Novo PAC — Ministério das Cidades.</p>
      <div style="margin-top:90px;">
        <div class="srchint">
          <span class="srchint-txt">Acesse a fonte dos dados no clique abaixo</span>
          <svg class="srchint-arrow" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v14"></path><path d="M6 12l6 6 6-6"></path></svg>
        </div>
        <div class="label" style="color:var(--amber-bright); margin-bottom:20px;">Fonte dos dados</div>
        <a class="srclink" href="https://www.gov.br/cidades/pt-br/acesso-a-informacao/acoes-e-programas/pac/selecoes-novo-pac/investimentos-selecionados" target="_blank" rel="noopener">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6"></path><path d="M20 4 10 14"></path><path d="M18 13v6a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h6"></path></svg>
          gov.br/cidades — Seleções Novo PAC · Investimentos Selecionados
        </a>
        <p style="font-size:22px; color:var(--blue-soft); margin:14px 0 0; font-variant-numeric:tabular-nums;">www.gov.br/cidades/pt-br/acesso-a-informacao/acoes-e-programas/pac/selecoes-novo-pac/investimentos-selecionados</p>
      </div>
    </div>
    <hr class="hr">
    <div class="bot" style="border-top:0; padding-top:18px;">
      <span class="org" style="color:var(--paper);">Ministério das Cidades - DMP/SE</span>
      <div style="flex: 1;"></div>
      <span class="org" style="color:var(--paper);">Atualizado em {DATA_ATUALIZACAO}<span class="verwrap"><button type="button" class="ver verlink" id="verBtn" aria-haspopup="dialog" aria-expanded="false" title="Ver histórico de atualizações">versão {VERSAO}</button>
        <div class="verlog" id="verLog" role="dialog" aria-label="Histórico de atualizações" hidden>
          <h3>Atualizações da apresentação</h3>
          <ul>
{HISTORICO_LI}
          </ul>
        </div></span></span>
      <a class="xlsxlink" href="{PDF_NAME}" download title="Baixar o PDF da apresentação (um slide por página) — com filtros aplicados, o PDF é gerado refletindo os filtros">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 21V3h9l5 5v13z"></path><path d="M14 3v6h6"></path></svg>
        <span class="pdftxt">Exportar PDF</span>
      </a>
      <a class="xlsxlink" href="data/{os.path.basename(XLSX)}" download title="Baixar a base de dados em Excel — arquivo protegido por senha">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12"></path><path d="M7 10l5 5 5-5"></path><path d="M4 19h16"></path></svg>
        XLSX Base
        <svg class="lock" viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="10" width="16" height="11" rx="2"></rect><path d="M8 10V7a4 4 0 0 1 8 0v3"></path></svg>
      </a>
    </div>
  </section>

</deck-stage>

<!-- popup do histórico de atualizações (versão no rodapé do encerramento) -->
<script>
(function () {{
  var btn = document.getElementById('verBtn');
  var log = document.getElementById('verLog');
  if (!btn || !log) return;
  function setOpen(open) {{
    log.hidden = !open;
    btn.setAttribute('aria-expanded', String(open));
  }}
  btn.addEventListener('click', function (e) {{
    e.stopPropagation();
    setOpen(log.hidden);
  }});
  document.addEventListener('click', function (e) {{
    if (!log.hidden && !log.contains(e.target)) setOpen(false);
  }});
  document.addEventListener('keydown', function (e) {{
    if (e.key === 'Escape' && !log.hidden) setOpen(false);
  }});
}})();
</script>

<!-- tela cheia no toque (iPhone/iPad): esconde a interface do navegador -->
<button type="button" class="fsbtn" id="fsBtn" aria-label="Tela cheia" title="Tela cheia">
  <svg id="fsBtnIco" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 9V5a1 1 0 0 1 1-1h4M20 9V5a1 1 0 0 0-1-1h-4M4 15v4a1 1 0 0 0 1 1h4M20 15v4a1 1 0 0 1-1 1h-4"></path></svg>
</button>
<script>
(function () {{
  var b = document.getElementById('fsBtn');
  var ico = document.getElementById('fsBtnIco');
  var el = document.documentElement;
  var req = el.requestFullscreen || el.webkitRequestFullscreen;
  if (!req) {{ b.style.display = 'none'; return; }}
  function fsEl() {{ return document.fullscreenElement || document.webkitFullscreenElement; }}
  function sync() {{
    ico.innerHTML = fsEl()
      ? '<path d="M9 4v3a2 2 0 0 1-2 2H4M20 9h-3a2 2 0 0 1-2-2V4M9 20v-3a2 2 0 0 0-2-2H4M15 20v-3a2 2 0 0 1 2-2h3"/>'
      : '<path d="M4 9V5a1 1 0 0 1 1-1h4M20 9V5a1 1 0 0 0-1-1h-4M4 15v4a1 1 0 0 0 1 1h4M20 15v4a1 1 0 0 1-1 1h-4"/>';
  }}
  b.addEventListener('click', function () {{
    if (fsEl()) (document.exitFullscreen || document.webkitExitFullscreen).call(document);
    else req.call(el);
  }});
  document.addEventListener('fullscreenchange', sync);
  document.addEventListener('webkitfullscreenchange', sync);
}})();
</script>
"""

APP_JS = """
// ===== dados e filtros (UF, status e ano) =====
// linha: [modIdx, fonteIdx, uf, kind, gov, vlrOGU, vlrFIN, ano]
// fonte: 0=FIN 1=OGU 2=OGU/FIN · kind: 0=migrada 1=selecionada 2=enquadrada · gov: grupo Governadores
// ano: rótulo de exibição ('' = migrada, fora do recorte por ano)
const MODS = __MODS__;
const ROWS = __ROWS__;
const ANOS = __ANOS__;

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
  { id: 't-prev',   ds: 'all',  status: true,
    mods: [MODS.indexOf('Drenagem Urbana'), MODS.indexOf('Contenção de Encostas')],
    cols: ['qf', 'qo', 'n', 'ogu', 'fin', 'tot'] },
  { id: 't-gov',    ds: 'gov',  status: false, cols: ['qf', 'qo', 'fin', 'ogu'] },
];

const state = { uf: '', sel: true, enq: true, anos: {} };
ANOS.forEach((a) => { state.anos[a] = true; });
const isSel = (r) => r[3] <= 1; // migradas (kind 0) contam como selecionadas
const anoFiltrado = () => ANOS.some((a) => !state.anos[a]); // algum ano desligado

function aggregate(t) {
  const keep = DATASETS[t.ds];
  const acc = MODS.map(() => ({ qf: 0, qo: 0, n: 0, ogu: 0, fin: 0, tot: 0 }));
  for (const r of ROWS) {
    if (!keep(r)) continue;
    if (state.uf && r[2] !== state.uf) continue;
    if (t.mods && t.mods.indexOf(r[0]) < 0) continue;
    if (t.status && !(isSel(r) ? state.sel : state.enq)) continue;
    // com recorte de ano ativo, as migradas (sem ano de seleção, r[7]==='')
    // também saem da conta — senão o OGU delas aparece como se fosse do ano filtrado
    if (t.status && anoFiltrado() && !state.anos[r[7]]) continue;
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
  const stname = state.sel && state.enq ? 'selecionadas + enquadradas/habilitadas'
    : (state.sel ? 'somente selecionadas' : 'somente enquadradas/habilitadas');
  const anosOn = ANOS.filter((a) => state.anos[a]);
  const anoname = anosOn.length === ANOS.length ? 'todos' : anosOn.join(' + ') + ' (sem migradas)';
  document.querySelectorAll('.ufname').forEach((el) => { el.textContent = nome; });
  document.querySelectorAll('.stname').forEach((el) => { el.textContent = stname; });
  document.querySelectorAll('.anoname').forEach((el) => { el.textContent = anoname; });
  document.querySelectorAll('select.ufsel').forEach((el) => { el.value = state.uf; });
  document.querySelectorAll('.stbtn:not(.anobtn)').forEach((el) => {
    el.setAttribute('aria-pressed', String(state[el.dataset.st]));
  });
  document.querySelectorAll('.anobtn').forEach((el) => {
    el.setAttribute('aria-pressed', String(!!state.anos[el.dataset.ano]));
  });
}

(function () {
  const params = new URLSearchParams(location.search);
  let uf = (params.get('uf') || '').toUpperCase().trim();
  if (uf && !ROWS.some((r) => r[2] === uf)) uf = '';
  state.uf = uf;
  const st = (params.get('status') || '').toLowerCase().trim();
  if (st === 'sel') state.enq = false; else if (st === 'enq') state.sel = false;
  const anoParam = (params.get('ano') || '').trim();
  if (anoParam) {
    const pedidos = anoParam.split(',').map((a) => a.trim()).filter((a) => ANOS.includes(a));
    if (pedidos.length) ANOS.forEach((a) => { state.anos[a] = pedidos.includes(a); });
  }
  render();
  document.querySelectorAll('select.ufsel').forEach((el) => {
    el.addEventListener('change', function () { state.uf = this.value; render(); });
  });
  document.querySelectorAll('.stbtn:not(.anobtn)').forEach((el) => {
    el.addEventListener('click', function () {
      const k = this.dataset.st, other = k === 'sel' ? 'enq' : 'sel';
      if (state[k] && !state[other]) return; // pelo menos um status ativo
      state[k] = !state[k];
      render();
    });
  });
  document.querySelectorAll('.anobtn').forEach((el) => {
    el.addEventListener('click', function () {
      const a = this.dataset.ano;
      const ligados = ANOS.filter((y) => state.anos[y]);
      if (state.anos[a] && ligados.length === 1) return; // pelo menos um ano ativo
      state.anos[a] = !state.anos[a];
      render();
    });
  });

  // ---- Exportar PDF: com filtros aplicados, gera e baixa o PDF refletindo os filtros ----
  // Sem filtros, baixa o PDF pré-gerado (vetorial, leve). Com qualquer recorte ativo
  // (UF, status ou ano), monta o PDF no navegador — cada slide é rasterizado em
  // 1920×1080 (horizontal) com html2canvas e adicionado a um jsPDF — sem diálogo
  // de impressão. Se algo falhar, cai para window.print() (@page já é paisagem).
  function carregarScript(src) {
    return new Promise(function (res, rej) {
      const s = document.createElement('script');
      s.src = src; s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }
  async function capturarSlide(secao) {
    const wrap = document.createElement('div');
    wrap.style.cssText = 'position:fixed; left:-20000px; top:0; width:1920px; height:1080px; overflow:hidden; background:#fff; z-index:-1;';
    const clone = secao.cloneNode(true);
    clone.removeAttribute('data-deck-active');
    clone.style.position = 'relative';
    clone.style.inset = 'auto';
    clone.style.width = '1920px';
    clone.style.height = '1080px';
    clone.style.boxSizing = 'border-box';
    clone.style.visibility = 'visible';
    clone.style.opacity = '1';
    // cloneNode não copia a seleção feita via JS nos <select>
    const selO = secao.querySelectorAll('select');
    clone.querySelectorAll('select').forEach((s, i) => { s.value = selO[i].value; });
    // atributos SVG com var() não resolvem na rasterização; fixa a cor computada
    const svgO = secao.querySelectorAll('svg, svg *');
    clone.querySelectorAll('svg, svg *').forEach((el, i) => {
      ['stroke', 'fill'].forEach((p) => {
        if ((el.getAttribute(p) || '').indexOf('var(') !== -1) {
          el.setAttribute(p, getComputedStyle(svgO[i])[p]);
        }
      });
    });
    // como na impressão, os botões de download ficam fora do PDF
    clone.querySelectorAll('.xlsxlink').forEach((el) => el.remove());
    // html2canvas não desenha a arte da capa (filho com z-index negativo);
    // troca o SVG por um <img> rasterizado, que ele pinta normalmente
    const net = clone.querySelector('.covernet');
    if (net) {
      net.setAttribute('width', 884);
      net.setAttribute('height', 1080);
      const img = document.createElement('img');
      img.style.cssText = 'position:absolute; top:0; right:0; height:100%; width:46%; opacity:.42; z-index:0;';
      const carregou = new Promise((res) => { img.onload = res; img.onerror = res; });
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(net));
      net.replaceWith(img);
      Array.from(clone.children).forEach((ch) => {
        if (ch !== img) { if (!ch.style.position) ch.style.position = 'relative'; ch.style.zIndex = '1'; }
      });
      await carregou;
    }
    wrap.appendChild(clone);
    document.body.appendChild(wrap);
    try {
      return await html2canvas(clone, { scale: 1, width: 1920, height: 1080, logging: false, backgroundColor: null });
    } finally {
      wrap.remove();
    }
  }
  let gerandoPdf = false;
  async function gerarPdfFiltrado(link) {
    if (gerandoPdf) return;
    gerandoPdf = true;
    const txt = link.querySelector('.pdftxt');
    const antes = txt ? txt.textContent : '';
    try {
      if (!window.html2canvas) await carregarScript('assets/html2canvas.min.js');
      if (!window.jspdf) await carregarScript('assets/jspdf.umd.min.js');
      const secoes = Array.from(document.querySelectorAll('deck-stage > section'))
        .filter((s) => !s.hasAttribute('data-deck-skip'));
      const pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'px', format: [1920, 1080], compress: true, hotfixes: ['px_scaling'] });
      for (let i = 0; i < secoes.length; i++) {
        if (txt) txt.textContent = 'Gerando PDF… ' + (i + 1) + '/' + secoes.length;
        const canvas = await capturarSlide(secoes[i]);
        if (i) pdf.addPage([1920, 1080], 'landscape');
        pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, 1920, 1080);
      }
      const partes = ['novopac-mcid-investimentos', (state.uf || 'brasil').toLowerCase()];
      if (!(state.sel && state.enq)) partes.push(state.sel ? 'selecionadas' : 'enquadradas-habilitadas');
      const anosOn = ANOS.filter((a) => state.anos[a]);
      if (anosOn.length !== ANOS.length) partes.push(anosOn.join('-'));
      pdf.save(partes.join('_') + '.pdf');
    } catch (err) {
      window.print();
    } finally {
      if (txt) txt.textContent = antes;
      gerandoPdf = false;
    }
  }
  const pdfLink = document.querySelector('a.xlsxlink[href$=".pdf"]');
  if (pdfLink) {
    pdfLink.addEventListener('click', function (e) {
      const filtrado = state.uf !== '' || !state.sel || !state.enq || ANOS.some((a) => !state.anos[a]);
      if (!filtrado) return;
      e.preventDefault();
      gerarPdfFiltrado(pdfLink);
    });
  }

  // ---- navegação por rolagem do mouse: para baixo avança, para cima volta ----
  const stage = document.querySelector('deck-stage');
  const menu = document.getElementById('deckMenu');

  // setas ↑/↓ também passam os slides (↓ avança, ↑ volta) — ←/→ já são nativas do deck
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    if (!stage || !stage.next) return;
    if (menu && menu.hasAttribute('data-open')) return;
    const tag = (e.target && e.target.tagName) || '';
    if (tag === 'SELECT' || tag === 'INPUT' || tag === 'TEXTAREA') return;
    e.preventDefault();
    if (e.key === 'ArrowDown') stage.next(); else stage.prev();
  });
  let ultimoAvanco = 0, acumulado = 0;
  window.addEventListener('wheel', function (e) {
    if (!stage || !stage.next) return;
    if (menu && menu.hasAttribute('data-open')) return;
    const agora = Date.now();
    if (agora - ultimoAvanco < 400) { acumulado = 0; return; }
    if ((acumulado > 0) !== (e.deltaY > 0)) acumulado = 0;
    acumulado += e.deltaY;
    if (Math.abs(acumulado) < 60) return;
    if (acumulado > 0) stage.next(); else stage.prev();
    acumulado = 0; ultimoAvanco = agora;
  }, { passive: true });
})();
"""

app_js = APP_JS.replace('__MODS__', json.dumps(MODS, ensure_ascii=False)) \
               .replace('__ANOS__', json.dumps(ANOS)) \
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
