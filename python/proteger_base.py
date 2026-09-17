#!/usr/bin/env python3
"""Protege com senha um arquivo de base em data/ e o renomeia no padrão <nome>_<DDMMYYYY>_<HHMM>.xlsx.

Mesma convenção das fotos do banco (ver python/extrair_base_unica.py): o repositório é público,
então toda base publicada vai cifrada com a SENHA_BASE_XLSX do config.env e com a data/hora da
extração no nome.

- .xlsx: cifrado como está, sem reescrever (preserva cabeçalho fora da 1ª linha, abas, formatação);
- .csv: convertido para .xlsx (CSV não aceita senha) e então cifrado.

A data/hora sai do nome antigo (AAAAMMDD_HHMM, AAAAMMDDHHMM ou AAAAMMDD) ou de --quando.
O arquivo de origem não é apagado — confira o resultado e remova depois.

Uso:
  uv run --with pandas --with openpyxl --with python-dotenv --with msoffcrypto-tool \
         python python/proteger_base.py data/arquivo.xlsx [--nome base_completa] [--quando 18082026_1126]
"""
import argparse
import io
import os
import re

import msoffcrypto
import pandas as pd
from dotenv import load_dotenv
from msoffcrypto.format.ooxml import OOXMLFile

HERE = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(HERE)

ap = argparse.ArgumentParser()
ap.add_argument('arquivo')
ap.add_argument('--nome', help='prefixo do arquivo novo (padrão: o nome atual, sem a data)')
ap.add_argument('--quando', help='data/hora no formato DDMMYYYY_HHMM (padrão: extraída do nome atual)')
args = ap.parse_args()

load_dotenv(os.path.join(HERE, 'config.env'))
SENHA = os.getenv('SENHA_BASE_XLSX')
if not SENHA:
    raise SystemExit('Defina SENHA_BASE_XLSX em python/config.env.')

origem = os.path.abspath(args.arquivo)
base = os.path.basename(origem)


def quando_do_nome(nome):
    """AAAAMMDD_HHMM / AAAAMMDDHHMM / AAAAMMDD no nome antigo → DDMMYYYY_HHMM."""
    m = re.search(r'(\d{4})(\d{2})(\d{2})[_-]?(\d{4})?', nome)
    if not m:
        raise SystemExit(f'Não achei a data no nome "{nome}" — use --quando DDMMYYYY_HHMM.')
    ano, mes, dia, hora = m.groups()
    return f'{dia}{mes}{ano}_{hora or "0000"}'


quando = args.quando or quando_do_nome(base)
prefixo = args.nome or re.sub(r'[_-]?\d{4}\d{2}\d{2}[_-]?\d{0,4}$', '', os.path.splitext(base)[0])
destino = os.path.join(os.path.dirname(origem), f'{prefixo}_{quando}.xlsx')
if os.path.exists(destino):
    raise SystemExit(f'{os.path.relpath(destino, RAIZ)} já existe — apague antes de gerar de novo.')

if origem.lower().endswith('.csv'):
    # CSV não aceita senha: vira xlsx (uma linha de cabeçalho) e só então é cifrado
    buf = io.BytesIO()
    pd.read_csv(origem).to_excel(buf, index=False, sheet_name=prefixo[:31])
    buf.seek(0)
else:
    buf = io.BytesIO(open(origem, 'rb').read())
    if msoffcrypto.OfficeFile(io.BytesIO(buf.getvalue())).is_encrypted():
        raise SystemExit(f'{base} já está protegido.')

with open(destino, 'wb') as f:
    OOXMLFile(buf).encrypt(SENHA, f)

# confere que abre com a senha (e só com ela)
with open(destino, 'rb') as f:
    protegido = msoffcrypto.OfficeFile(f)
    if not protegido.is_encrypted():
        raise SystemExit('ERRO: gravado sem senha — não publique este arquivo.')
    protegido.load_key(password=SENHA)
    aberto = io.BytesIO()
    protegido.decrypt(aberto)
linhas = len(pd.read_excel(aberto, header=None))
print(f'{base} → {os.path.relpath(destino, RAIZ)} (protegido, {linhas} linhas na 1ª aba)')
