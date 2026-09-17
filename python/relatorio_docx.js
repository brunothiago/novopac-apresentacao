// Gera relatorio_antes_depois_<AAAAMMDD>.docx a partir do .json de python/relatorio_antes_depois.py
// Uso: node python/relatorio_docx.js relatorio_antes_depois_<AAAAMMDD>.json   (requer o pacote npm "docx")
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, PageOrientation, Footer, PageNumber, LevelFormat, BorderStyle,
  VerticalAlign,
} = require('docx');

const entrada = process.argv[2];
const R = JSON.parse(fs.readFileSync(entrada, 'utf8'));
const saida = entrada.replace(/\.json$/, '.docx');

const AZUL = '1351B4', NAVY = '071D41', CINZA = 'EDF1F7', LINHA = 'C5CCD6', VERMELHO = 'C0392B', VERDE = '168821';
const FONTE = 'Calibri';
const QTD = new Set(['qf', 'qo', 'n']);
const LARGURA = 15398; // A4 deitado com margens de 720 DXA

const fQ = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });
const fV = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fP = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const num = (v, qtd) => (qtd ? fQ : fV).format(v);
const dif = (v, qtd) => {
  const zero = Math.abs(v) < (qtd ? 0.5 : 0.005);
  return { txt: zero ? '—' : (v > 0 ? '+' : '') + num(v, qtd), cor: zero ? '777777' : (v < 0 ? VERMELHO : VERDE) };
};

const borda = { style: BorderStyle.SINGLE, size: 4, color: LINHA };
const semBorda = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const bordas = { top: semBorda, left: semBorda, right: semBorda, bottom: borda };

function celula(texto, largura, o = {}) {
  return new TableCell({
    width: { size: largura, type: WidthType.DXA },
    borders: bordas,
    verticalAlign: VerticalAlign.CENTER,
    shading: o.fundo ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fundo } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    columnSpan: o.span,
    children: [new Paragraph({
      alignment: o.alinha || AlignmentType.RIGHT,
      children: [new TextRun({ text: texto, font: FONTE, size: o.tam || 17, bold: o.negrito, color: o.cor })],
    })],
  });
}

// tabela antes × depois; cols = métricas; pct = coluna Dif. %
function tabela(q, cols, pct) {
  const sub = pct ? ['Antes', 'Depois', 'Diferença', 'Dif. %'] : ['Antes', 'Depois', 'Diferença'];
  const wLabel = q.rotulo === 'UF' ? 1400 : 3300;
  const nDados = cols.length * sub.length;
  const wDado = Math.floor((LARGURA - wLabel) / nDados);
  const larguras = [wLabel, ...Array(nDados).fill(wDado)];
  const total = larguras.reduce((a, b) => a + b, 0);
  const cab = { fundo: AZUL, cor: 'FFFFFF', negrito: true, alinha: AlignmentType.CENTER, tam: 16 };
  const linhas = [
    new TableRow({ tableHeader: true, children: [
      celula(q.rotulo, wLabel, { ...cab, alinha: AlignmentType.LEFT }),
      ...cols.map((c) => celula(R.nome_col[c], wDado * sub.length, { ...cab, span: sub.length })),
    ] }),
    new TableRow({ tableHeader: true, children: [
      celula('', wLabel, cab),
      ...cols.flatMap(() => sub.map((s) => celula(s, wDado, { ...cab, fundo: NAVY }))),
    ] }),
  ];
  const idx = cols.map((c) => q.colunas.indexOf(c) + 1);
  for (const l of q.linhas) {
    const tot = l[0] === 'TOTAL';
    const base = { negrito: tot, fundo: tot ? CINZA : undefined };
    const cels = [celula(l[0], wLabel, { ...base, alinha: AlignmentType.LEFT })];
    cols.forEach((c, i) => {
      const [a, d, df] = l[idx[i]];
      const qtd = QTD.has(c);
      const x = dif(df, qtd);
      cels.push(celula(num(a, qtd), wDado, base), celula(num(d, qtd), wDado, base),
                celula(x.txt, wDado, { ...base, cor: x.cor }));
      if (pct) {
        const p = a ? df / a * 100 : null;
        const zero = p === null || Math.abs(p) < 0.05;
        cels.push(celula(p === null ? '—' : (zero ? '—' : (p > 0 ? '+' : '') + fP.format(p) + '%'), wDado,
                         { ...base, cor: zero ? '777777' : (p < 0 ? VERMELHO : VERDE) }));
      }
    });
    linhas.push(new TableRow({ cantSplit: true, children: cels }));
  }
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: larguras, rows: linhas });
}

// quebra as métricas em blocos que caibam na largura da página
function tabelas(q, cols) {
  cols = cols || q.colunas;
  const porBloco = q.pct ? 2 : 3;
  const qtd = cols.filter((c) => QTD.has(c)), val = cols.filter((c) => !QTD.has(c));
  const grupos = cols.length <= porBloco ? [cols] : [qtd, val].filter((g) => g.length);
  const out = [];
  grupos.forEach((g, i) => {
    for (let k = 0; k < g.length; k += porBloco) {
      if (out.length) out.push(espaco(120));
      out.push(tabela(q, g.slice(k, k + porBloco), q.pct));
    }
  });
  return out;
}

const p = (texto, o = {}) => new Paragraph({
  spacing: { after: o.depois ?? 120, before: o.antes ?? 0 },
  children: [new TextRun({ text: texto, font: FONTE, size: o.tam || 21, italics: o.italico, color: o.cor, bold: o.negrito })],
});
const espaco = (n) => new Paragraph({ spacing: { after: n }, children: [] });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 140 }, children: [new TextRun({ text: t })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 }, keepNext: true, children: [new TextRun({ text: t })] });
const item = (t) => new Paragraph({ numbering: { reference: 'marcadores', level: 0 }, spacing: { after: 80 },
  children: [new TextRun({ text: t, font: FONTE, size: 20 })] });

const Q = Object.fromEntries(R.quadros.map((q) => [q.id, q]));
const corpo = [
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'Novo PAC — MCid · Investimentos', font: FONTE, size: 22, bold: true, color: AZUL })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 120 }, children: [new TextRun({ text: 'Relatório antes × depois da troca da fonte de dados' })] }),
  p(`Antes: apresentação versão ${R.versao_antes}, dados de ${R.data_antes} (planilha e CSV exportados). ` +
    `Depois: versão ${R.versao_depois}, dados de ${R.data_depois} (tabela se_cgpac.tab_base_unica_gm do banco).`),
  p('Valores em R$ milhões. Diferença = depois − antes. Cada quadro repete a tabela do slide sem filtros ' +
    '(Brasil, todos os status e anos).', { italico: true, cor: '555555', tam: 19, depois: 200 }),

  h1('Principais números'),
  tabela(Q['Resumo'], ['n', 'tot'], true), espaco(120), tabela(Q['Resumo'], ['ogu', 'fin'], true),
  p(Q['Resumo'].nota, { italico: true, cor: '555555', tam: 18, antes: 80 }),

  h1('O que muda na base'),
  ...R.notas.slice(2).map(item),
];

corpo.push(h1('Tabelas por slide'));
for (const q of R.quadros.filter((x) => x.id.startsWith('Slide'))) {
  corpo.push(h2(q.titulo), ...tabelas(q));
}

corpo.push(h1('Novas seleções por ano e status'), ...tabelas(Q['Por ano']),
  p(Q['Por ano'].nota, { italico: true, cor: '555555', tam: 18, antes: 80 }));
corpo.push(h1('Por UF'), p('Total da apresentação por UF (quantidade e valor total).', { tam: 19, cor: '555555' }),
  ...tabelas(Q['Por UF'], ['n', 'tot']));

// maiores variações de valor por proposta
corpo.push(h1('Maiores variações de valor por proposta (seleções)'),
  p(`Entraram ${R.contagem.Entrou || 0}, saíram ${R.contagem.Saiu || 0} e mudaram ${R.contagem.Mudou || 0} propostas. ` +
    'As 15 maiores variações de valor total estão abaixo; a lista completa está na planilha (aba "Propostas").', { tam: 20 }));
const wM = [3400, 3000, 900, 4398, 1200, 2500];
const cabM = { fundo: AZUL, cor: 'FFFFFF', negrito: true, alinha: AlignmentType.LEFT, tam: 16 };
corpo.push(new Table({
  width: { size: wM.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: wM,
  rows: [
    new TableRow({ tableHeader: true, children: ['Situação', 'Nº proposta', 'UF', 'Modalidade', 'Fonte', 'Dif. total (R$ mi)']
      .map((t, i) => celula(t, wM[i], { ...cabM, alinha: i === 5 ? AlignmentType.RIGHT : AlignmentType.LEFT })) }),
    ...R.maiores.map((m) => {
      const x = dif(m[5], false);
      return new TableRow({ cantSplit: true, children: [
        ...m.slice(0, 5).map((t, i) => celula(String(t), wM[i], { alinha: AlignmentType.LEFT })),
        celula(x.txt, wM[5], { cor: x.cor }),
      ] });
    }),
  ],
}));

const doc = new Document({
  creator: 'Ministério das Cidades - DMP/SE',
  title: 'Relatório antes × depois — Novo PAC MCid',
  styles: {
    default: { document: { run: { font: FONTE, size: 21 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', next: 'Normal', run: { font: FONTE, size: 44, bold: true, color: NAVY } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONTE, size: 30, bold: true, color: NAVY }, paragraph: { outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONTE, size: 24, bold: true, color: AZUL }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: 'marcadores', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
      margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: `Novo PAC — MCid · antes (${R.data_antes}) × depois (${R.data_depois}) · página `, font: FONTE, size: 16, color: '777777' }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONTE, size: 16, color: '777777' }),
    ] })] }) },
    children: corpo,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(saida, buf);
  console.log('escrito:', path.basename(saida));
});
