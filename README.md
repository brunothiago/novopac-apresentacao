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
- **Botões "Status"** (Selecionadas / Enquadradas): nas tabelas que somam os dois status
  (totais, quantidade, fonte e novas seleções), ligam/desligam cada status na contagem.
  Via URL: `index.html?status=sel` ou `?status=enq`. As tabelas de Migrado, Enquadradas e
  Governadores já são recortes de status por natureza e não mudam.
- **Botões "Ano"** (2023 / 2025 / 2026): recorte por ano de seleção nas mesmas tabelas.
  Os rótulos são de exibição (nos dados, `ano_selecao` 2024→"2023", 2025→"2025",
  2026→"2026" — ver `ANO_LABEL` no build.py; `ano_selecao` 2023 conta como 2024). Com algum
  ano desligado, as migradas saem da conta. Via URL: `?ano=2026` ou `?ano=2023,2025`.
- **Rolagem do mouse**: scroll para baixo avança o slide, para cima volta
  (além de ← / → e do toque).
- **Exportar PDF** (slide final): baixa `novopac-mcid-investimentos.pdf`, pré-gerado no build
  (um slide por página, paisagem 16:9) — no celular abre direto no visualizador com o botão
  de compartilhar. Ctrl/Cmd+P no desktop também funciona, no recorte selecionado.
- **No celular/tablet**: navegação por toque nas metades esquerda/direita da tela;
  a barra de controles inferior fica oculta para não cobrir o conteúdo. O botão redondo
  no canto inferior direito entra em **tela cheia** (esconde a interface do Safari;
  iOS 16.4+). Para uma experiência de app sem nenhuma barra, use
  Compartilhar → **Adicionar à Tela de Início** e abra pelo ícone.

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

## Versão

O rodapé exibe "versão X.Y" (constante `VERSAO` no topo do `build.py`).
Convenção: **subir o número a cada commit** que altere a apresentação e rodar o build.

## Dados e convenções

Fonte única: tabela **`se_cgpac.tab_base_unica_gm`** (banco `corporativo`), lida a partir de uma
foto em `data/base_unica_gm_<DDMMYYYY>_<HHMM>.xlsx` — o build não conecta no banco. A carga e as
regras ficam em `dados.py`, usado pelos dois builds. Uma versão por extração completa do banco; a
data/hora no nome é a da carga mais recente da tabela, e o build usa sempre a foto mais nova.

**A foto vai protegida por senha.** O repositório é público e ela é o arquivo do botão "XLSX Base"
do painel, então é gravada com a criptografia padrão do Excel, usando `SENHA_BASE_XLSX` do
`python/config.env` (arquivo gitignorado). A senha é repassada fora do repositório; sem ela, nem o
build nem quem baixar pelo painel abre a planilha. A foto leva só as colunas que a planilha pública
anterior já trazia + `dte_carga` — as colunas de controle interno da tabela ficam de fora.

| Recorte (17/09/2026) | Filtro na tabela |
|---|---|
| Migradas (557) | `status_selecao == "retomada"` |
| Seleções (2.890, sem MCMV) | `status_selecao` `"selecionada"` ou `"enquadrada"` |

- Convenção de totais (igual à apresentação original): **migradas + selecionadas + enquadradas FIN**.
- Cada linha conta 1. Propostas com OGU e FIN vêm em duas linhas (uma por fonte) e contam uma
  vez em cada fonte — não existe mais a fonte "OGU/FIN".
- "Atualizado em" é a data da carga mais recente da tabela (`max(dte_carga)`).
- As 2 propostas MCMV (`MCMV FNHIS` / `MCMV FNHIS SUB50`) ficam **fora da contagem** da
  versão simples; a versão detalhada ainda as inclui (como "MCMV (Calamidade RS)").
- Rótulo: `Médias e Grandes Cidades` → "Mobilidade: Médias e Grandes Cidades".

## Como atualizar os dados

1. Tire a foto do banco (requer VPN e `python/config.env` com as credenciais e a `SENHA_BASE_XLSX`):
   `uv run --with pandas --with openpyxl --with sqlalchemy --with psycopg2-binary --with python-dotenv --with msoffcrypto-tool python python/extrair_base_unica.py`
   — grava `data/base_unica_gm_<DDMMYYYY>_<HHMM>.xlsx` já com senha; os builds usam a foto mais recente.
2. Rode `python3 build.py` e `python3 build_detalhada.py` (requerem `pandas`, `openpyxl`,
   `msoffcrypto-tool` e `python-dotenv` — os dois últimos para abrir a foto protegida; com o
   Google Chrome instalado, o PDF `novopac-mcid-investimentos.pdf` é regenerado no mesmo passo).
   Suba a `VERSAO` e acrescente a linha em `HISTORICO` no `build.py`.
3. O `index.html` é regenerado com os dados embutidos — as tabelas são recalculadas
   no navegador a partir das propostas individuais, por isso o filtro de UF funciona offline.
4. (Opcional) Relatório antes × depois em relação ao último commit:
   `python3 python/relatorio_antes_depois.py` (gera o .xlsx e um .json) e
   `node python/relatorio_docx.js relatorio_antes_depois_<AAAAMMDD>.json` (gera o .docx; requer o pacote npm `docx`).

> **Não edite o `index.html` à mão**: todo ajuste vai no `build.py`, senão o próximo build desfaz.

## Estrutura

```
index.html              apresentação simples (autocontida, com filtro de UF)
versao-detalhada.html   apresentação detalhada (painéis, barras e slides extras)
build.py                gerador da versão simples — monta o index.html
build_detalhada.py      gerador da versão detalhada
dados.py                carga da foto do banco (regras de recorte comuns aos dois builds)
data/                   fotos protegidas da tabela se_cgpac.tab_base_unica_gm (e extrações antigas)
python/                 extrator do banco, relatório antes × depois e conciliação se_pac
assets/deck.css         design system do deck (gov.br)
assets/deck-stage.js    web component <deck-stage> (navegação, escala, impressão)
assets/deck-chrome.html barra de controles, menu de slides e dica de rotação
```
