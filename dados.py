"""Carga dos dados da apresentação — usada por build.py e build_detalhada.py.

Fonte: foto da tabela se_cgpac.tab_base_unica_gm em data/base_unica_gm_<DDMMYYYY>_<HHMM>.xlsx
(gerada por python/extrair_base_unica.py). A foto mais recente é usada.

A foto publicada vai protegida por senha (SENHA_BASE_XLSX em python/config.env); aqui ela é
decifrada em memória, sem gravar cópia aberta em disco.

- migradas: status_selecao == 'retomada' (ano_selecao é o ano do contrato — não é usado);
- seleções: status_selecao 'selecionada' ou 'enquadrada';
- ano_selecao 2023 conta como 2024 (mesmo botão "2023" do recorte por ano);
- cada linha conta 1 — propostas com OGU e FIN vêm em duas linhas (uma por fonte).
"""
import glob
import io
import os
from datetime import datetime

import msoffcrypto
import pandas as pd
from dotenv import load_dotenv

HERE = os.path.dirname(os.path.abspath(__file__))
VAL = ['vlr_portaria_total', 'vlr_portaria_ogu', 'vlr_portaria_fin']
STATUS = {'retomada', 'selecionada', 'enquadrada'}
# nome atual e o das primeiras fotos, antes do versionamento por data/hora
FORMATOS = ('%d%m%Y_%H%M', '%Y%m%d')


def quando(caminho):
    """Data/hora da carga a partir do nome do arquivo (o nome DDMMYYYY não ordena sozinho)."""
    marca = os.path.basename(caminho)[len('base_unica_gm_'):-len('.xlsx')]
    for f in FORMATOS:
        try:
            return datetime.strptime(marca, f)
        except ValueError:
            pass
    return datetime.min


def foto_mais_recente():
    fotos = glob.glob(f'{HERE}/data/base_unica_gm_[0-9]*.xlsx')
    if not fotos:
        raise SystemExit('Nenhuma foto em data/base_unica_gm_*.xlsx — rode python/extrair_base_unica.py')
    return max(fotos, key=quando)


def abrir(caminho):
    """Devolve o xlsx pronto para o pandas, decifrando com a senha se estiver protegido."""
    with open(caminho, 'rb') as f:
        arquivo = msoffcrypto.OfficeFile(f)
        if not arquivo.is_encrypted():
            return caminho
        load_dotenv(f'{HERE}/python/config.env')
        senha = os.getenv('SENHA_BASE_XLSX')
        if not senha:
            raise SystemExit(f'{os.path.basename(caminho)} está protegida por senha — '
                             'defina SENHA_BASE_XLSX em python/config.env.')
        aberto = io.BytesIO()
        try:
            arquivo.load_key(password=senha)
            arquivo.decrypt(aberto)
        except Exception as e:
            raise SystemExit(f'Não consegui abrir {os.path.basename(caminho)} com a SENHA_BASE_XLSX '
                             f'de python/config.env: {e}')
    return aberto


def load(caminho=None):
    """Retorna (selecoes, migradas, data_atualizacao 'dd/mm/aaaa', caminho da foto)."""
    caminho = caminho or foto_mais_recente()
    df = pd.read_excel(abrir(caminho))
    desconhecidos = set(df.status_selecao.dropna()) - STATUS
    if desconhecidos or df.status_selecao.isna().any():
        raise SystemExit(f'status_selecao inesperado na foto: {sorted(desconhecidos) or "vazio"}')
    for c in VAL:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    df['uf'] = df.uf.astype(str).str.strip().str.upper()
    data = pd.to_datetime(df.dte_carga).max()

    mig = df[df.status_selecao == 'retomada'].copy()
    x = df[df.status_selecao != 'retomada'].copy()
    x['ano_selecao'] = x.ano_selecao.astype(int).replace({2023: 2024})
    return x, mig, f'{data:%d/%m/%Y}', caminho
