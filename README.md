# Novo PAC — MCid · Investimentos

Apresentação em HTML (deck 16:9, estilo Padrão Digital de Governo — gov.br) com o balanço
da carteira de seleções do Novo PAC no Ministério das Cidades.

## Como usar

Abra **`index.html`** em qualquer navegador (duplo clique). Não precisa de servidor.

Há também a **`versao-detalhada.html`**: mesma base de dados, com painéis de destaque,
barras proporcionais e dois slides extras (Calamidade RS e Seleções 2026), sem filtro de UF.

- **← / →** navegam entre os slides · **M** abre o menu de slides · **F** tela cheia · **R** volta ao início
- **Seletor "Recorte"** (canto superior direito de cada tabela): filtra todas as tabelas por UF.
  Também é possível abrir direto num estado com `index.html?uf=SP`.
- **Imprimir / PDF**: Ctrl/Cmd+P gera um PDF com um slide por página, no recorte selecionado.

Cada tabela traz a **data de atualização**, a **legenda do que entra na contagem**
(migradas, selecionadas e/ou enquadradas) e o **recorte** ativo.

## Slides

1. Capa
2. Investimentos totais — migradas + selecionadas + enquadradas (FIN)
3. Investimentos por quantidade — qtd. FIN × OGU por modalidade
4. Investimento por fonte — quantidades e valores OGU/FIN/total
5. Novo PAC Migrado — carteira herdada (> dez/2022)
6. Novas seleções — 2024–2026, sem migradas
7. Propostas enquadradas — somente FIN
8. Governadores — seleções pactuadas com governos estaduais
9. Encerramento

## Dados e convenções

| Recorte | Fonte | Filtro |
|---|---|---|
| Migradas (557) | `data/view_sis_novopac_previsto_unificado_202608180817.csv` | `origem_dado == "Novo PAC - Retomada"` |
| Seleções (2.865) | `data/base_completa_atualizada_20260818_1126.xlsx` (header na 2ª linha) | tudo (selecionadas + enquadradas) |

- Convenção de totais (igual à apresentação original): **migradas + selecionadas + enquadradas FIN**.
- 5 propostas de Mobilidade têm fonte OGU/FIN e contam apenas na coluna de total.
- 1 proposta MCMV tem abrangência nacional (UF "BR"): conta no Brasil, não aparece no seletor de estados.
- Rótulos: `Médias e Grandes Cidades` → "Mobilidade: Médias e Grandes Cidades";
  `MCMV FNHIS`/`MCMV FNHIS SUB50` → "MCMV (Calamidade RS)".

## Como atualizar os dados

1. Substitua os arquivos em `data/` pelas novas extrações (ajuste os nomes/data em `build.py`).
2. Rode `python3 build.py` (requer `pandas` e `openpyxl`).
3. O `index.html` é regenerado com os dados embutidos — as tabelas são recalculadas
   no navegador a partir das propostas individuais, por isso o filtro de UF funciona offline.

## Estrutura

```
index.html              apresentação simples (autocontida, com filtro de UF)
versao-detalhada.html   apresentação detalhada (painéis, barras e slides extras)
build.py                gerador da versão simples — lê data/ e monta o index.html
build_detalhada.py      gerador da versão detalhada
data/                   extrações-fonte (xlsx + csv)
assets/deck.css         design system do deck (gov.br)
assets/deck-stage.js    web component <deck-stage> (navegação, escala, impressão)
assets/deck-chrome.html barra de controles, menu de slides e dica de rotação
```
