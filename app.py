import io
import re
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter
from decimal import Decimal, InvalidOperation

import pandas as pd

NS = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

# Limite mínimo de itens com ao menos uma palavra de marca/produto em comum
LIMITE_MARCA = 0.70

_RE_SCI = re.compile(r'^[0-9.,]+E[+\-]?[0-9]+$')
_RE_DIGITS = re.compile(r'[^0-9]')
_RE_ALPHA = re.compile(r'[^a-z]')
_RE_TOKENS = re.compile(r'[A-Z0-9]+')
_RE_EXT = re.compile(r'\.[A-Za-z0-9]+$')
_RE_NUMS = re.compile(r'\d{3,}')


# ==========================================================
# 1. EAN
# ==========================================================
def _dv_gtin_ok(digitos: str) -> bool:
    """Valida o dígito verificador GS1 (GTIN-8/12/13/14), completando com zeros à esquerda."""
    d = digitos.zfill(14)
    if len(d) != 14 or not d.isdigit():
        return False
    soma = sum(int(c) * (3 if i % 2 == 0 else 1) for i, c in enumerate(d[:13]))
    return (10 - soma % 10) % 10 == int(d[13])


def normalizar_ean(valor):
    """
    Devolve (ean, status):
      ean    -> só dígitos, SEM zeros à esquerda ('' se vazio / sem GTIN)
      status -> 'ok' | 'vazio' | 'ilegivel'
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return '', 'vazio'
    s = str(valor).strip().upper()
    if s in ('', 'NAN', 'NONE', 'SEM GTIN', 'S/GTIN'):
        return '', 'vazio'

    if 'E' in s and _RE_SCI.fullmatch(s):
        try:
            dec = Decimal(s.replace(',', '.'))
        except InvalidOperation:
            return '', 'ilegivel'
        inteiro = format(dec, 'f').split('.')[0]
        mantissa, expo = s.split('E', 1)
        digitos_mantissa = len(_RE_DIGITS.sub('', mantissa))
        total_digitos = int(expo) + 1
        if digitos_mantissa < total_digitos and not _dv_gtin_ok(inteiro):
            return '', 'ilegivel'
        return inteiro.lstrip('0'), 'ok'

    if s.endswith('.0'):
        s = s[:-2]
    apenas = _RE_DIGITS.sub('', s)
    return apenas.lstrip('0'), ('ok' if apenas.strip('0') else 'vazio')


# ==========================================================
# 2. LEITURA DO CSV/XLSX DO WINTHOR + ORDENAÇÃO
# ==========================================================
def _sem_acento(txt: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFKD', str(txt))
        if not unicodedata.combining(c)
    )


def _achar_coluna_seq(colunas):
    for c in colunas:
        n = _RE_ALPHA.sub('', _sem_acento(c).lower())
        if 'seq' in n or n in ('item', 'nitem', 'noitem'):
            return c
    return None


def _ler_bytes(arquivo) -> bytes:
    """Lê bytes de UploadedFile (Streamlit), BytesIO ou caminho local com segurança."""
    if hasattr(arquivo, 'getvalue'):
        return arquivo.getvalue()
    if hasattr(arquivo, 'read'):
        if hasattr(arquivo, 'seek'):
            arquivo.seek(0)
        dados = arquivo.read()
        if hasattr(arquivo, 'seek'):
            arquivo.seek(0)
        return dados
    with open(arquivo, 'rb') as f:
        return f.read()


def ler_csv_winthor(arquivo):
    """Lê o arquivo, ordena por 'Nº seq.' (numérico) e devolve (df, avisos)."""
    avisos = []
    nome = getattr(arquivo, 'name', str(arquivo)).lower()
    bruto = _ler_bytes(arquivo)

    if nome.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(io.BytesIO(bruto), dtype=str, engine='openpyxl')
    else:
        texto = None
        for enc in ('utf-8-sig', 'cp1252', 'latin1'):
            try:
                texto = bruto.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if texto is None:
            texto = bruto.decode('latin1', errors='replace')
        df = pd.read_csv(io.StringIO(texto), sep=';', dtype=str)

    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how='all').reset_index(drop=True)

    col_seq = _achar_coluna_seq(df.columns)
    if col_seq is None:
        avisos.append("Coluna de sequência ('Nº seq.') não encontrada: mantida a ordem original do arquivo.")
        df['_seq'] = range(1, len(df) + 1)
    else:
        seq = pd.to_numeric(df[col_seq].astype(str).str.strip(), errors='coerce')
        invalidas = seq.isna()
        if invalidas.any():
            avisos.append(
                f"{int(invalidas.sum())} linha(s) sem número de sequência válido foram ignoradas (rodapé/total?)."
            )
            df = df[~invalidas].copy()
            seq = seq[~invalidas]
        df['_seq'] = seq.astype(int)
        df = df.sort_values('_seq', kind='stable').reset_index(drop=True)

    return df, avisos


def _to_float(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    s = str(v).strip()
    if not s or s.upper() in ('NAN', 'NONE'):
        return 0.0
    if '.' in s and ',' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def _to_num_str(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return _RE_DIGITS.sub('', s)


def montar_itens_csv(df):
    """Converte o DataFrame ordenado na lista de itens usada pelo cruzamento."""
    col_tipo = next(
        (c for c in df.columns if 'tipo' in c.lower() and ('cálculo' in c.lower() or 'calculo' in c.lower())),
        None
    )
    itens = []
    for _, row in df.iterrows():
        ean, st_ean = normalizar_ean(row.get('Cód. barra', ''))
        itens.append({
            'Linha_WINTHOR': int(row['_seq']),
            'Cod_Produto_WINTHOR': _to_num_str(row.get('Cód. produto', '')),
            'EAN_WINTHOR': ean,
            'EAN_Comp_WINTHOR': ean,
            'EAN_Status_WINTHOR': st_ean,
            'Descricao_WINTHOR': str(row.get('Descrição', '')).strip(),
            'Qtd_WINTHOR': _to_float(row.get('Qtde Pré-Entrada', 0.0)),
            'NCM_WINTHOR': _to_num_str(row.get('NCM', '')),
            'ICMS_WINTHOR': _to_float(row.get('%ICMS', 0.0)),
            '% ICMS red. Whintor': _to_float(row.get('% ICMS red.', 0.0)),
            '%Alíq.Ext. ST Guia_Whintor': _to_float(row.get('%Alíq.Ext.ST Guia', 0.0)),
            'Tipo_Calculo_WINTHOR': str(row.get(col_tipo, '')).strip() if col_tipo else '',
        })
    return itens


# ==========================================================
# 3. LEITURA DO XML
# ==========================================================
def ler_xml_nfe(
    arquivo,
    extrair_fator_embalagem=None,
    obter_descricao_origem=None,
    limpar_numero=None,
    limpar_ean=None,
):
    """Devolve (numero_nf, itens_xml). Compatível com UploadedFile do Streamlit."""
    limpar_numero = limpar_numero or _to_num_str
    limpar_ean = limpar_ean or (lambda v: normalizar_ean(v)[0])
    extrair_fator_embalagem = extrair_fator_embalagem or (lambda u: 1.0)
    obter_descricao_origem = obter_descricao_origem or (lambda o: str(o))

    bruto = _ler_bytes(arquivo)
    root
