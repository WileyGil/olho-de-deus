import io
import os
import re
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime
from decimal import Decimal, InvalidOperation
from itertools import zip_longest

import pandas as pd
import streamlit as st
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E TEMA CLARO
# ==========================================
st.set_page_config(page_title="Auditoria de Notas | Dashboard", page_icon="📊", layout="wide")


@st.cache_resource
def _garantir_tema_claro():
    os.makedirs(".streamlit", exist_ok=True)
    config_path = os.path.join(".streamlit", "config.toml")
    if not os.path.exists(config_path):
        with open(config_path, "w", encoding="utf-8") as f:
            f.write(
                '[theme]\n'
                'base = "light"\n'
                'primaryColor = "#1B2631"\n'
                'backgroundColor = "#F4F6F9"\n'
                'secondaryBackgroundColor = "#FFFFFF"\n'
                'textColor = "#2C3E50"\n'
            )
    return True


_garantir_tema_claro()

st.markdown("""
    <style>
        :root {
            color-scheme: light !important;
            --primary-color: #1B2631 !important;
            --background-color: #F4F6F9 !important;
            --secondary-background-color: #FFFFFF !important;
            --text-color: #2C3E50 !important;
        }
        html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"], .stApp {
            background-color: #F4F6F9 !important;
            color: #2C3E50 !important;
        }
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2rem;
            max-width: 1180px;
        }

        /* Cabeçalho Executivo */
        .dashboard-header {
            background: linear-gradient(135deg, #16202A 0%, #243447 100%);
            border: 1px solid #1B2631;
            border-left: 5px solid #27AE60;
            padding: 16px 22px;
            border-radius: 10px;
            margin-bottom: 18px;
            box-shadow: 0 4px 12px rgba(22, 32, 42, 0.14);
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 12px;
        }
        .header-left {
            display: flex;
            flex-direction: column;
            gap: 3px;
        }
        .dashboard-title {
            font-size: 1.38rem;
            font-weight: 700;
            margin: 0;
            color: #FFFFFF !important;
            letter-spacing: -0.01em;
        }
        .dashboard-subtitle {
            font-size: 0.86rem;
            color: #CBD5E1 !important;
            margin: 0;
        }
        .header-pill {
            background-color: rgba(255, 255, 255, 0.12);
            color: #FFFFFF !important;
            border: 1px solid rgba(255, 255, 255, 0.25);
            font-size: 0.72rem;
            font-weight: 700;
            padding: 5px 12px;
            border-radius: 20px;
            letter-spacing: 0.05em;
            text-transform: uppercase;
        }

        /* Títulos de seção com indicador de etapa */
        .section-label {
            font-size: 0.92rem;
            font-weight: 600;
            color: #1B2631;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .step-num {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 22px;
            height: 22px;
            border-radius: 50%;
            background-color: #1B2631;
            color: #FFFFFF;
            font-size: 0.75rem;
            font-weight: 700;
        }

        /* Uploaders claros e compactos */
        [data-testid="stFileUploader"],
        [data-testid="stFileUploader"] > div,
        [data-testid="stFileUploaderDropzone"] {
            background-color: #FFFFFF !important;
            color: #1B2631 !important;
        }
        [data-testid="stFileUploaderDropzone"] {
            border: 1.5px dashed #CBD5E1 !important;
            border-radius: 8px !important;
            padding: 12px 14px !important;
            transition: border-color 0.2s ease, background-color 0.2s ease;
        }
        [data-testid="stFileUploaderDropzone"]:hover {
            border-color: #2C3E50 !important;
            background-color: #F8FAFC !important;
        }
        [data-testid="stFileUploaderDropzone"] div,
        [data-testid="stFileUploaderDropzone"] span,
        [data-testid="stFileUploaderDropzone"] small,
        [data-testid="stFileUploaderDropzone"] p {
            color: #334155 !important;
        }
        [data-testid="stFileUploaderDropzone"] button {
            background-color: #EFF6FF !important;
            color: #1D4ED8 !important;
            border: 1px solid #BFDBFE !important;
            font-weight: 600 !important;
            border-radius: 6px !important;
        }

        /* Lista de arquivos com rolagem controlada (ideal para 30+ XMLs) */
        [data-testid="stFileUploaderDropzone"] > div:last-child,
        [data-testid="stFileUploaderDropzone"] ul {
            width: 100% !important;
            max-width: 100% !important;
            max-height: 145px !important;
            overflow-y: auto !important;
            flex: 1 1 100% !important;
        }
        [data-testid="stFileUploaderFile"],
        [data-testid="stUploadedFile"] {
            background-color: #EEF2F6 !important;
            border: 1px solid #CBD5E1 !important;
            border-radius: 8px !important;
            color: #1B2631 !important;
            width: 100% !important;
            max-width: 100% !important;
            padding: 7px 10px !important;
            margin-bottom: 4px !important;
            display: flex !important;
            align-items: center !important;
        }
        [data-testid="stFileUploaderFile"] > div,
        [data-testid="stUploadedFile"] > div {
            background-color: transparent !important;
            border: none !important;
            color: #1B2631 !important;
            max-width: 100% !important;
            width: 100% !important;
            flex: 1 1 auto !important;
            overflow: visible !important;
            text-overflow: unset !important;
        }
        [data-testid="stFileUploaderFileName"],
        [data-testid="stFileUploaderFile"] span,
        [data-testid="stFileUploaderFile"] div,
        [data-testid="stUploadedFile"] span,
        [data-testid="stUploadedFile"] div {
            color: #1B2631 !important;
            max-width: none !important;
            width: auto !important;
            white-space: normal !important;
            overflow-wrap: anywhere !important;
            word-break: break-all !important;
            overflow: visible !important;
            text-overflow: unset !important;
            font-weight: 600 !important;
            font-size: 0.82rem !important;
            line-height: 1.3 !important;
        }
        [data-testid="stFileUploaderFile"] small,
        [data-testid="stUploadedFile"] small {
            color: #5D6D7E !important;
            font-size: 0.74rem !important;
            display: block !important;
            margin-top: 2px !important;
        }
        [data-testid="stFileUploaderFile"] svg,
        [data-testid="stUploadedFile"] svg,
        [data-testid="stFileUploaderDeleteBtn"] button,
        [data-testid="stFileUploaderFile"] button,
        [data-testid="stUploadedFile"] button {
            color: #2C3E50 !important;
            fill: #2C3E50 !important;
            background-color: transparent !important;
            border: none !important;
            flex-shrink: 0 !important;
        }

        /* Containers em fundo branco */
        [data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF !important;
            border-color: #E2E8F0 !important;
            border-radius: 8px !important;
        }

        /* Checkboxes compactos */
        div[data-testid="stCheckbox"] {
            margin-bottom: 4px;
        }
        div[data-testid="stCheckbox"] label p {
            font-size: 0.88rem !important;
            color: #2C3E50 !important;
            font-weight: 500 !important;
        }

        /* Barra do lote */
        .lote-pill {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background-color: #EEF2F6;
            color: #1B2631;
            border: 1px solid #CBD5E1;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.82rem;
            font-weight: 600;
        }

        /* Selo de status fixo e separado do botão */
        .temp-notice-wrap {
            display: flex;
            justify-content: center;
            align-items: center;
            width: 100%;
            margin-top: 8px;
            margin-bottom: 16px;
        }
        .temp-notice {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            background-color: #EAFAF1;
            color: #145A32;
            border: 1px solid #A9DFBF;
            padding: 8px 22px;
            border-radius: 8px;
            font-size: 0.88rem;
            font-weight: 600;
            box-shadow: 0 1px 4px rgba(25, 111, 61, 0.08);
        }

        /* Estado vazio discreto */
        .empty-hint {
            text-align: center;
            color: #64748B;
            font-size: 0.86rem;
            padding: 12px;
            background-color: #FFFFFF;
            border: 1px dashed #CBD5E1;
            border-radius: 8px;
            margin-top: 6px;
        }

        /* Alertas de bloqueio/aviso */
        .v-badge {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 0.84rem;
            font-weight: 500;
            padding: 8px 14px;
            border-radius: 6px;
            margin-bottom: 8px;
        }
        .v-badge.warn {
            background-color: #FEF9E7;
            color: #7D6608;
            border: 1px solid #F9E79F;
        }
        .v-badge.err {
            background-color: #FADBD8;
            color: #78281F;
            border: 1px solid #F5B7B1;
        }

        /* Botão de Download em Verde Executivo */
        div[data-testid="stDownloadButton"] > button {
            background-color: #196F3D !important;
            color: #FFFFFF !important;
            border: 1px solid #145A32 !important;
            font-weight: 600 !important;
            font-size: 0.95rem !important;
            padding: 10px 24px !important;
            border-radius: 8px !important;
            box-shadow: 0 3px 8px rgba(25, 111, 61, 0.22) !important;
        }
        div[data-testid="stDownloadButton"] > button:hover {
            background-color: #145A32 !important;
            border-color: #0E3E22 !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONSTANTES E REGEX PRÉ-COMPILADAS
# ==========================================
NS = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
LIMITE_MARCA = 0.70

_RE_SCI = re.compile(r'^[0-9.,]+E[+\-]?[0-9]+$')
_RE_DIGITS = re.compile(r'[^0-9]')
_RE_ALPHA = re.compile(r'[^a-z]')
_RE_TOKENS = re.compile(r'[A-Z0-9]+')
_RE_EXT = re.compile(r'\.[A-Za-z0-9]+$')
_RE_NUMS = re.compile(r'\d{3,}')
_RE_NNF_BYTES = re.compile(rb'<(?:[A-Za-z0-9_]+:)?nNF>\s*(\d+)\s*</(?:[A-Za-z0-9_]+:)?nNF>')
_IGNORAR = {'COM', 'PARA', 'DOS', 'DAS', 'UNID', 'UND', 'CX', 'KIT', 'PCT', 'FR', 'REF', 'ESC', 'ATIV'}


# ==========================================
# 1. FUNÇÕES DE SANEAMENTO, EAN E LEITURA
# ==========================================
def extrair_numero_nf_rapido(bruto: bytes) -> str:
    """Extrai o número da NF instantaneamente sem parsear o XML inteiro."""
    m = _RE_NNF_BYTES.search(bruto[:8192]) or _RE_NNF_BYTES.search(bruto)
    if m:
        return m.group(1).decode('ascii', errors='ignore').strip()
    return 'DESCONHECIDO'


def _dv_gtin_ok(digitos: str) -> bool:
    d = digitos.zfill(14)
    if len(d) != 14 or not d.isdigit():
        return False
    soma = sum(int(c) * (3 if i % 2 == 0 else 1) for i, c in enumerate(d[:13]))
    return (10 - soma % 10) % 10 == int(d[13])


def normalizar_ean(valor):
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


def limpar_numero(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    val_str = str(valor).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return _RE_DIGITS.sub('', val_str)


def converte_float_br(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return 0.0
    val_str = str(valor).strip()
    if val_str == '' or val_str.upper() in ('NAN', 'NONE'):
        return 0.0
    try:
        if '.' in val_str and ',' in val_str:
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '.')
        return float(val_str)
    except ValueError:
        return 0.0


def obter_descricao_origem(orig_cod):
    tabela_origem = {
        '0': '0 - Nacional',
        '1': '1 - Estrangeira (Imp. Direta)',
        '2': '2 - Estrangeira (Mercado Interno)',
        '3': '3 - Nacional (Imp. > 40% e <= 70%)',
        '4': '4 - Nacional (Processo Produtivo Básico)',
        '5': '5 - Nacional (Imp. <= 40%)',
        '6': '6 - Estrangeira (Sem Similar Nacional)',
        '7': '7 - Estrangeira (Mercado Interno, Sem Similar)',
        '8': '8 - Nacional (Imp. > 70%)'
    }
    return tabela_origem.get(str(orig_cod).strip(), f"{orig_cod} - Desconhecida")


def extrair_fator_embalagem(uCom):
    if not uCom:
        return 1.0
    match = re.search(r'(?:DP|CX|PC|UN)x?(\d+)', str(uCom), re.IGNORECASE)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            pass
    match_nums = re.findall(r'\d+', str(uCom))
    if match_nums:
        try:
            return float(match_nums[-1])
        except ValueError:
            pass
    return 1.0


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


@st.cache_data(show_spinner=False)
def ler_csv_winthor_cached(bruto: bytes, nome_arquivo: str):
    avisos = []
    nome = (nome_arquivo or '').lower()

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


def montar_itens_csv(df):
    col_tipo = next(
        (c for c in df.columns if 'tipo' in c.lower() and ('cálculo' in c.lower() or 'calculo' in c.lower())),
        None
    )
    if not col_tipo and len(df.columns) >= 2:
        col_tipo = df.columns[-2]

    itens = []
    for _, row in df.iterrows():
        ean, st_ean = normalizar_ean(row.get('Cód. barra', ''))
        itens.append({
            'Linha_WINTHOR': int(row['_seq']),
            'Cod_Produto_WINTHOR': limpar_numero(row.get('Cód. produto', '')),
            'EAN_WINTHOR': ean,
            'EAN_Comp_WINTHOR': ean,
            'EAN_Status_WINTHOR': st_ean,
            'Descricao_WINTHOR': str(row.get('Descrição', '')).strip(),
            'Qtd_WINTHOR': converte_float_br(row.get('Qtde Pré-Entrada', 0.0)),
            'NCM_WINTHOR': limpar_numero(row.get('NCM', '')),
            'ICMS_WINTHOR': converte_float_br(row.get('%ICMS', 0.0)),
            '% ICMS red. Whintor': converte_float_br(row.get('% ICMS red.', 0.0)),
            '%Alíq.Ext. ST Guia_Whintor': converte_float_br(row.get('%Alíq.Ext.ST Guia', 0.0)),
            'Tipo_Calculo_WINTHOR': str(row.get(col_tipo, '')).strip() if col_tipo else '',
        })
    return itens


@st.cache_data(show_spinner=False)
def ler_xml_nfe_bytes(bruto: bytes):
    root = ET.fromstring(bruto)

    nNF = root.find('.//ns:ide/ns:nNF', NS)
    numero_nf = nNF.text.strip() if nNF is not None and nNF.text else 'DESCONHECIDO'

    def txt(no, tag):
        if no is None:
            return ''
        e = no.find(f'ns:{tag}', NS)
        return e.text.strip() if e is not None and e.text else ''

    itens = []
    for det in root.findall('.//ns:det', NS):
        prod = det.find('ns:prod', NS)
        try:
            n_item = int(det.attrib.get('nItem', '0'))
        except ValueError:
            n_item = det.attrib.get('nItem', '')

        ean_raw, ean_trib_raw = txt(prod, 'cEAN'), txt(prod, 'cEANTrib')
        ean_comp, _ = normalizar_ean(ean_raw)
        ean_trib_comp, _ = normalizar_ean(ean_trib_raw)
        ean_exib = limpar_numero(ean_raw) if ean_comp != '' else limpar_numero(ean_trib_raw)
        u_com = txt(prod, 'uCom') or 'UN'

        icms, origem = 0.0, '0 - Nacional'
        icms_node = det.find('.//ns:ICMS', NS)
        if icms_node is not None and len(icms_node) > 0:
            sub = icms_node[0]
            p = sub.find('ns:pICMS', NS)
            if p is not None and p.text:
                icms = converte_float_br(p.text)
            o = sub.find('ns:orig', NS)
            if o is not None and o.text:
                origem = obter_descricao_origem(o.text.strip())

        pred_node = (
            icms_node[0].find('ns:pRedBC', NS)
            if icms_node is not None and len(icms_node) > 0
            else None
        )
        pred = converte_float_br(pred_node.text) if pred_node is not None and pred_node.text else 0.0

        itens.append({
            'Linha_NF': n_item,
            'Descricao_NF': txt(prod, 'xProd'),
            'EAN_NF': ean_exib,
            'EAN_Comp_NF': ean_comp,
            'EANTrib_Comp_NF': ean_trib_comp,
            'NCM_NF': limpar_numero(txt(prod, 'NCM')),
            'Qtd_NF': converte_float_br(txt(prod, 'qCom')),
            '_Fator_Embalagem': extrair_fator_embalagem(u_com),
            'Vl_Unit_NF': converte_float_br(txt(prod, 'vUnCom')),
            'Vl_Total_NF': converte_float_br(txt(prod, 'vProd')),
            'Origem_XML': origem,
            'ICMS_NF': icms,
            'pRedBC_NF': pred,
        })
    return numero_nf, itens


def _tokens(desc):
    t = _RE_TOKENS.findall(_sem_acento(desc).upper())
    return {x for x in t if len(x) >= 3 and not any(ch.isdigit() for ch in x) and x not in _IGNORAR}


def _numeros_do_nome(nome):
    stem = _RE_EXT.sub('', str(nome))
    return {g.lstrip('0') for g in _RE_NUMS.findall(stem)}


def validar_upload(itens_xml, itens_csv, numero_nf, nome_arquivo_csv):
    bloqueios, avisos, infos = [], [], []
    r = {
        'bloqueios': bloqueios,
        'avisos': avisos,
        'infos': infos,
        'marca_em_comum': [],
        'bloqueio_nf': False,
    }

    n_x, n_c = len(itens_xml), len(itens_csv)
    if n_x != n_c:
        bloqueios.append(f"Quantidade de itens diferente: XML tem {n_x} e o arquivo do WinThor tem {n_c}.")
    else:
        infos.append(f"Quantidade de itens confere: {n_x}.")

    ilegiveis = [i['Linha_WINTHOR'] for i in itens_csv if i['EAN_Status_WINTHOR'] == 'ilegivel']
    if ilegiveis:
        lista = ', '.join(map(str, ilegiveis[:10])) + ('…' if len(ilegiveis) > 10 else '')
        bloqueios.append(
            f"{len(ilegiveis)} EAN(s) do WinThor estão em notação científica truncada e não podem ser recuperados "
            f"(Nº seq.: {lista}). Reexporte com a coluna 'Cód. barra' formatada como texto."
        )

    if n_x == n_c and n_x > 0:
        compat = 0
        cont_x, cont_c = Counter(), Counter()
        for ix, ic in zip(itens_xml, itens_csv):
            tx, tc = _tokens(ix['Descricao_NF']), _tokens(ic['Descricao_WINTHOR'])
            ean_x = ix.get('EAN_Comp_NF', '')
            ean_c = ic.get('EAN_Comp_WINTHOR', '')
            mesmo_ean = bool(ean_x and ean_c and ean_x in (ean_c, ix.get('EANTrib_Comp_NF')))

            if (tx & tc) or mesmo_ean:
                compat += 1

            cont_x.update(tx)
            cont_c.update(tc)

        pct = compat / n_x
        comum = [
            t for t, _ in sorted(
                ((t, min(cont_x[t], cont_c[t])) for t in cont_x.keys() & cont_c.keys()),
                key=lambda p: -p[1]
            )[:4]
        ]
        r['marca_em_comum'] = comum
        if pct < LIMITE_MARCA:
            bloqueios.append(
                f"Os produtos não parecem ser da mesma nota: só {compat} de {n_x} itens ({pct:.0%}) "
                f"têm descrição ou EAN compatível entre XML e WinThor."
            )
        else:
            infos.append(
                f"Produtos compatíveis em {compat} de {n_x} itens ({pct:.0%})"
                + (f"; termos em comum: {', '.join(comum)}." if comum else '.')
            )

    candidatos = _numeros_do_nome(nome_arquivo_csv)
    nf_norm = str(numero_nf).strip().lstrip('0')
    if candidatos:
        if nf_norm in candidatos:
            infos.append(f"Número da NF confere com o nome do arquivo ({numero_nf}).")
        else:
            bloqueios.append(
                f"O nome do arquivo ({nome_arquivo_csv}) não contém o número da NF do XML ({numero_nf})."
            )
            r['bloqueio_nf'] = len(bloqueios) == 1
    else:
        avisos.append("O nome do arquivo não traz o número da NF; a conferência foi feita só por itens e produtos.")

    return r


@st.cache_data(show_spinner=False)
def gerar_excel_auditoria_cached(
    itens_xml_tuple, itens_csv_tuple, numero_nf: str,
    inc_ident: bool, inc_ean: bool, inc_desc: bool, inc_qtd: bool,
    inc_ncm: bool, inc_icms: bool, inc_red: bool, inc_st: bool
):
    itens_xml = [dict(x) for x in itens_xml_tuple]
    itens_csv = [dict(c) for c in itens_csv_tuple]
    data_simples = datetime.now().strftime('%Y%m%d')
    nome_arquivo_saida = f'relatorio_NF_{numero_nf}_{data_simples}.xlsx'

    dados_relatorio = []
    for item_x, item_c in zip_longest(itens_xml, itens_csv, fillvalue=None):
        linha = {}

        if item_x:
            linha.update(item_x)
        else:
            linha.update({
                'Linha_NF': '-', 'Descricao_NF': '-', 'EAN_NF': '-', 'NCM_NF': '-',
                'Qtd_NF': '-', '_Fator_Embalagem': 1.0, 'Vl_Unit_NF': '-', 'Vl_Total_NF': '-',
                'ICMS_NF': '-', 'pRedBC_NF': '-'
            })

        if item_c:
            linha.update(item_c)
        else:
            linha.update({
                'Linha_WINTHOR': '-', 'Cod_Produto_WINTHOR': '-', 'EAN_WINTHOR': '-',
                'Descricao_WINTHOR': '-', 'Qtd_WINTHOR': '-', 'NCM_WINTHOR': '-',
                'ICMS_WINTHOR': '-', '% ICMS red. Whintor': '-',
                '%Alíq.Ext. ST Guia_Whintor': '-', 'Tipo_Calculo_WINTHOR': '-'
            })

        if item_x and item_c:
            linha['Status_Geral'] = 'PAREADO'

            ean_x_comp = item_x.get('EAN_Comp_NF', '')
            eantrib_x_comp = item_x.get('EANTrib_Comp_NF', '')
            ean_win_comp = item_c.get('EAN_Comp_WINTHOR', '')

            if ean_win_comp != '' and (ean_win_comp == ean_x_comp or ean_win_comp == eantrib_x_comp):
                linha['Status_EAN'] = 'OK'
            elif ean_win_comp == '' and ean_x_comp == '' and eantrib_x_comp == '':
                linha['Status_EAN'] = 'OK'
            else:
                linha['Status_EAN'] = 'DIVERGENTE'

            icms_xml = round(item_x['ICMS_NF'], 2)
            icms_csv = round(item_c['ICMS_WINTHOR'], 2)
            st_guia = round(item_c['%Alíq.Ext. ST Guia_Whintor'], 2)
            if icms_xml == icms_csv or icms_xml == st_guia:
                linha['Status_ICMS_Aliq'] = 'OK'
            else:
                linha['Status_ICMS_Aliq'] = 'DIVERGENTE'

            pred_xml = round(item_x['pRedBC_NF'], 4)
            red_csv = round(item_c['% ICMS red. Whintor'], 4)
            carga_efetiva_xml = round(item_x['ICMS_NF'] * (1.0 - (item_x['pRedBC_NF'] / 100.0)), 4)

            if (
                carga_efetiva_xml == red_csv
                or round(carga_efetiva_xml, 2) == round(red_csv, 2)
                or pred_xml == red_csv
            ):
                linha['Status_Red_BC'] = 'OK'
            else:
                linha['Status_Red_BC'] = 'DIVERGENTE'

            tipo_calc = str(item_c.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
            if tipo_calc == 'N':
                linha['Tipo_Calculo_ST'] = 'NOTA FISCAL'
            elif tipo_calc == 'G':
                linha['Tipo_Calculo_ST'] = 'GUIA'
            else:
                linha['Tipo_Calculo_ST'] = item_c['Tipo_Calculo_WINTHOR']

            linha['Status_NCM'] = 'OK' if str(item_x['NCM_NF']) == str(item_c['NCM_WINTHOR']) else 'DIVERGENTE'

            qtd_xml_pura = item_x['Qtd_NF']
            qtd_win = item_c['Qtd_WINTHOR']
            fator = item_x['_Fator_Embalagem']
            qtd_xml_convertida = qtd_xml_pura * fator

            if round(qtd_xml_pura, 4) == round(qtd_win, 4) or round(qtd_xml_convertida, 4) == round(qtd_win, 4):
                linha['Status_Qtd'] = 'OK'
            else:
                linha['Status_Qtd'] = 'DIVERGENTE'

        else:
            linha['Status_Geral'] = 'SOBRA'
            for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                linha[st_col] = '-'

        dados_relatorio.append(linha)

    mapa_colunas_blocos = {
        'Identificação': ['Linha_NF', 'Linha_WINTHOR'],
        'Código & EAN': ['Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN'],
        'Descrições': ['Descricao_NF', 'Descricao_WINTHOR'],
        'Quantidades': ['Qtd_NF', 'Qtd_WINTHOR', 'Status_Qtd'],
        'NCM (Fiscal)': ['NCM_NF', 'NCM_WINTHOR', 'Status_NCM'],
        'ICMS': ['ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq'],
        'Redução de Base': ['pRedBC_NF', '% ICMS red. Whintor', 'Status_Red_BC'],
        'Tipo Cálculo ST': ['Tipo_Calculo_ST']
    }

    colunas_selecionadas = []
    if inc_ident: colunas_selecionadas.extend(mapa_colunas_blocos['Identificação'])
    if inc_ean: colunas_selecionadas.extend(mapa_colunas_blocos['Código & EAN'])
    if inc_desc: colunas_selecionadas.extend(mapa_colunas_blocos['Descrições'])
    if inc_qtd: colunas_selecionadas.extend(mapa_colunas_blocos['Quantidades'])
    if inc_ncm: colunas_selecionadas.extend(mapa_colunas_blocos['NCM (Fiscal)'])
    if inc_icms: colunas_selecionadas.extend(mapa_colunas_blocos['ICMS'])
    if inc_red: colunas_selecionadas.extend(mapa_colunas_blocos['Redução de Base'])
    if inc_st: colunas_selecionadas.extend(mapa_colunas_blocos['Tipo Cálculo ST'])

    if not colunas_selecionadas:
        colunas_selecionadas = ['Linha_NF']

    df_final = pd.DataFrame(dados_relatorio)
    df_final = df_final[[c for c in colunas_selecionadas if c in df_final.columns]]

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='Comparativo_ICMS_e_EAN', index=False, startrow=2)

        workbook = writer.book
        worksheet = workbook['Comparativo_ICMS_e_EAN']

        worksheet.sheet_view.showGridLines = False
        worksheet.sheet_view.zoomScale = 70

        worksheet.freeze_panes = 'A4'
        worksheet.auto_filter.ref = f"A3:{get_column_letter(len(df_final.columns))}{len(df_final) + 3}"

        COLOR_BLOCO_DARK = "1B2631"
        COLOR_XML_SOLID = "212F3D"
        COLOR_WIN_SOFT = "7F8C8D"

        fill_bloco = PatternFill(start_color=COLOR_BLOCO_DARK, end_color=COLOR_BLOCO_DARK, fill_type="solid")
        fill_xml = PatternFill(start_color=COLOR_XML_SOLID, end_color=COLOR_XML_SOLID, fill_type="solid")
        fill_win = PatternFill(start_color=COLOR_WIN_SOFT, end_color=COLOR_WIN_SOFT, fill_type="solid")

        fill_zebrado = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
        fill_cinza_sobra = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")

        fill_status_divergente = PatternFill(start_color="9C0006", end_color="9C0006", fill_type="solid")
        fill_status_convergente = PatternFill(start_color="196F3D", end_color="196F3D", fill_type="solid")
        fill_cel_divergente = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
        fill_cel_convergente = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")

        font_super_white = Font(bold=True, size=11, color="FFFFFF")
        font_header_white = Font(bold=True, size=10, color="FFFFFF")
        font_header_status = Font(bold=True, size=10, color="FFFFFF")
        font_divergente = Font(color="9C0006", bold=True)
        font_normal = Font(color="000000", bold=False)

        borda_padrao = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
        borda_fim_bloco = Border(left=Side(style=None), right=Side(style='medium', color='1B2631'), top=Side(style=None), bottom=Side(style=None))

        headers = list(df_final.columns)

        super_grupos_def = []
        if inc_ident: super_grupos_def.append(("IDENTIFICAÇÃO", ['Linha_NF', 'Linha_WINTHOR']))
        if inc_ean: super_grupos_def.append(("CÓDIGO E EAN", ['Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN']))
        if inc_desc: super_grupos_def.append(("DESCRIÇÕES", ['Descricao_NF', 'Descricao_WINTHOR']))
        if inc_qtd: super_grupos_def.append(("QUANTIDADES", ['Qtd_NF', 'Qtd_WINTHOR', 'Status_Qtd']))
        if inc_ncm: super_grupos_def.append(("CLASSIFICAÇÃO FISCAL (NCM)", ['NCM_NF', 'NCM_WINTHOR', 'Status_NCM']))
        if inc_icms: super_grupos_def.append(("ALÍQUOTAS DE ICMS", ['ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq']))
        if inc_red: super_grupos_def.append(("REDUÇÃO DE BASE", ['pRedBC_NF', '% ICMS red. Whintor', 'Status_Red_BC']))
        if inc_st: super_grupos_def.append(("TIPO CÁLCULO ST", ['Tipo_Calculo_ST']))

        curr_col_idx = 1
        for label, cols in super_grupos_def:
            cols_in_df = [c for c in cols if c in headers]
            if not cols_in_df:
                continue
            start_c = curr_col_idx
            end_c = curr_col_idx + len(cols_in_df) - 1
            curr_col_idx += len(cols_in_df)

            worksheet.merge_cells(start_row=1, start_column=start_c, end_row=2, end_column=end_c)
            cell = worksheet.cell(row=1, column=start_c)
            cell.value = label
            cell.font = font_super_white
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            for r in (1, 2):
                for c in range(start_c, end_c + 1):
                    cell_target = worksheet.cell(row=r, column=c)
                    cell_target.fill = fill_bloco
                    cell_target.border = borda_fim_bloco if c == end_c else borda_padrao

        worksheet.row_dimensions[1].height = 18
        worksheet.row_dimensions[2].height = 18
        worksheet.row_dimensions[3].height = 25

        status_columns = ['Status_EAN', 'Status_Qtd', 'Status_NCM', 'Status_ICMS_Aliq', 'Status_Red_BC']
        status_has_divergence = {sc: (df_final[sc] == 'DIVERGENTE').any() if sc in df_final.columns else False for sc in status_columns}

        def obter_fill_cabecalho(col_name):
            if 'WINTHOR' in col_name or 'Whintor' in col_name or 'Status' in col_name:
                return fill_win
            return fill_xml

        for idx, col_name in enumerate(headers, 1):
            cell = worksheet.cell(row=3, column=idx)
            cell.value = col_name
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            is_fim_bloco = False
            running_c = 0
            for _, cols in super_grupos_def:
                c_in = [c for c in cols if c in headers]
                if c_in:
                    running_c += len(c_in)
                    if idx == running_c:
                        is_fim_bloco = True
                        break

            cell.border = borda_fim_bloco if is_fim_bloco else borda_padrao

            if col_name in status_has_divergence:
                if status_has_divergence[col_name]:
                    cell.fill = fill_status_divergente
                    cell.font = font_header_status
                else:
                    cell.fill = fill_status_convergente
                    cell.font = font_header_status
            else:
                cell.fill = obter_fill_cabecalho(col_name)
                cell.font = font_header_white

        align_center_cols = [
            'Linha_NF', 'Linha_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Qtd_NF', 'Qtd_WINTHOR',
            'NCM_NF', 'NCM_WINTHOR', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor',
            'pRedBC_NF', '% ICMS red. Whintor'
        ]

        status_dependents = {
            'Status_EAN': ['EAN_NF', 'EAN_WINTHOR'],
            'Status_Qtd': ['Qtd_NF', 'Qtd_WINTHOR'],
            'Status_NCM': ['NCM_NF', 'NCM_WINTHOR'],
            'Status_ICMS_Aliq': ['ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor'],
            'Status_Red_BC': ['pRedBC_NF', '% ICMS red. Whintor']
        }

        for row in range(4, len(df_final) + 4):
            linha_dict = dados_relatorio[row - 4]
            for col_idx in range(1, len(headers) + 1):
                cell = worksheet.cell(row=row, column=col_idx)

                is_fim_bloco = False
                running_c = 0
                for _, cols in super_grupos_def:
                    c_in = [c for c in cols if c in headers]
                    if c_in:
                        running_c += len(c_in)
                        if col_idx == running_c:
                            is_fim_bloco = True
                            break
                cell.border = borda_fim_bloco if is_fim_bloco else borda_padrao

                col_name = headers[col_idx - 1]
                current_align = cell.alignment or Alignment()
                cell.alignment = Alignment(horizontal=current_align.horizontal or 'left', vertical='center', wrap_text=current_align.wrap_text)

                if col_name == 'Cod_Produto_WINTHOR' and str(cell.value).strip() not in ('', '-'):
                    try:
                        cell.value = int(float(cell.value))
                        cell.number_format = '0'
                    except ValueError:
                        pass
                elif col_name in ('NCM_NF', 'NCM_WINTHOR') and str(cell.value).strip() not in ('', '-'):
                    try:
                        cell.value = int(float(cell.value))
                        cell.number_format = '00000000'
                    except ValueError:
                        pass
                elif col_name in ('EAN_NF', 'EAN_WINTHOR') and str(cell.value).strip() not in ('', '-'):
                    cell.value = str(cell.value).strip()
                    cell.number_format = '@'
                elif isinstance(cell.value, (int, float)) and cell.value != '-':
                    if 'pRedBC' in col_name or 'red. Whintor' in col_name:
                        cell.number_format = '#,##0.0000'
                    elif 'Qtd' in col_name or 'ICMS' in col_name or 'Whintor' in col_name:
                        cell.number_format = '#,##0.00'

                if col_name == 'Cod_Produto_WINTHOR':
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_name in align_center_cols or 'Status' in col_name:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

            if 'SOBRA' in str(linha_dict.get('Status_Geral', '')):
                for col_idx in range(1, len(headers) + 1):
                    worksheet.cell(row=row, column=col_idx).fill = fill_cinza_sobra
                continue

            for col_status, colunas_dados in status_dependents.items():
                if col_status in headers:
                    idx_status = headers.index(col_status) + 1
                    cell_status = worksheet.cell(row=row, column=idx_status)
                    val_status = cell_status.value

                    if val_status == 'DIVERGENTE':
                        cell_status.fill = fill_cel_divergente
                        for col_afetada in colunas_dados:
                            if col_afetada in headers:
                                cell_dado = worksheet.cell(row=row, column=headers.index(col_afetada) + 1)
                                cell_dado.font = font_divergente
                    elif val_status == 'OK':
                        cell_status.fill = fill_cel_convergente
                        for col_afetada in colunas_dados:
                            if col_afetada in headers:
                                cell_dado = worksheet.cell(row=row, column=headers.index(col_afetada) + 1)
                                cell_dado.font = font_normal

        letras_nao_status = [get_column_letter(idx) for idx, col_name in enumerate(headers, 1) if not col_name.startswith('Status')]
        for l_col in letras_nao_status:
            rule = FormulaRule(formula=['MOD(ROW(),2)=0'], stopIfTrue=True, fill=fill_zebrado)
            worksheet.conditional_formatting.add(f'{l_col}4:{l_col}{len(df_final) + 3}', rule)

        for col in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value is not None:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                except Exception:
                    pass
            worksheet.column_dimensions[column_letter].width = max(max_length + 4, 12)

    excel_buffer.seek(0)
    return excel_buffer.getvalue(), nome_arquivo_saida


def _freeze_itens(itens: list[dict]) -> tuple[tuple[tuple[str, object], ...], ...]:
    return tuple(tuple(sorted(item.items())) for item in itens)


# ==========================================
# 2. CALLBACKS DE GERENCIAMENTO DO LOTE
# ==========================================
def _remover_nf_ativa():
    chave = st.session_state.get("sb_nf_ativa")
    if chave and chave in st.session_state["xml_lote"]:
        st.session_state["xml_lote"].pop(chave, None)
        st.session_state["xml_ignorados_sessao"].add(chave)
    restantes = list(st.session_state["xml_lote"].keys())
    if restantes:
        st.session_state["sb_nf_ativa"] = restantes[0]
    else:
        st.session_state.pop("sb_nf_ativa", None)


def _limpar_lote():
    st.session_state["xml_lote"] = {}
    st.session_state["xml_ignorados_sessao"] = set()
    st.session_state.pop("sb_nf_ativa", None)
    st.session_state["xml_uploader_key"] += 1


# ==========================================
# 3. INTERFACE STREAMLIT (UI/UX OTIMIZADA)
# ==========================================
st.markdown("""
    <div class="dashboard-header">
        <div class="header-left">
            <p class="dashboard-title">📊 Auditoria de Notas Fiscais</p>
            <p class="dashboard-subtitle">Conferência automatizada entre XML da NF-e e Pré-Entrada WinThor</p>
        </div>
        <span class="header-pill">Fiscal • XML × WinThor</span>
    </div>
""", unsafe_allow_html=True)

# Estado da sessão
if "xml_lote" not in st.session_state:
    st.session_state["xml_lote"] = {}
if "xml_ignorados_sessao" not in st.session_state:
    st.session_state["xml_ignorados_sessao"] = set()
if "xml_uploader_key" not in st.session_state:
    st.session_state["xml_uploader_key"] = 0
if "ultimo_csv_assinatura" not in st.session_state:
    st.session_state["ultimo_csv_assinatura"] = None

# Etapa 1 e 2: Uploaders lado a lado
col_up1, col_up2 = st.columns(2, gap="large")

with col_up1:
    st.markdown(
        '<p class="section-label"><span class="step-num">1</span> Lote de Notas Fiscais (XML)</p>',
        unsafe_allow_html=True
    )
    arquivos_xml = st.file_uploader(
        "Upload XMLs",
        type=["xml"],
        accept_multiple_files=True,
        key=f"xml_uploader_{st.session_state['xml_uploader_key']}",
        label_visibility="collapsed"
    )
    if arquivos_xml:
        assinaturas_atuais = {f"{a.name}:{a.size}" for a in arquivos_xml}
        st.session_state["xml_ignorados_sessao"] &= assinaturas_atuais

        for arq_x in arquivos_xml:
            sig_x = f"{arq_x.name}:{arq_x.size}"
            if sig_x in st.session_state["xml_ignorados_sessao"] or sig_x in st.session_state["xml_lote"]:
                continue
            bruto_x = arq_x.getvalue()
            num_nf_x = extrair_numero_nf_rapido(bruto_x)
            st.session_state["xml_lote"][sig_x] = {
                "numero_nf": num_nf_x,
                "nf_norm": num_nf_x.lstrip('0'),
                "nome_arquivo": arq_x.name,
                "bytes": bruto_x,
            }

with col_up2:
    st.markdown(
        '<p class="section-label"><span class="step-num">2</span> Relatório de Pré-Entrada (WinThor)</p>',
        unsafe_allow_html=True
    )
    arquivo_dados = st.file_uploader("Upload WinThor", type=["csv", "xlsx"], label_visibility="collapsed")

lote_xml = st.session_state["xml_lote"]
chaves_ordenadas = list(lote_xml.keys())

# Detecção automática da NF antes da instanciação do selectbox (quando um novo CSV/XLSX entra)
if arquivo_dados is not None:
    assinatura_csv = f"{arquivo_dados.name}:{arquivo_dados.size}"
    if assinatura_csv != st.session_state["ultimo_csv_assinatura"]:
        st.session_state["ultimo_csv_assinatura"] = assinatura_csv
        cand_csv = _numeros_do_nome(arquivo_dados.name)
        match_chave = next(
            (k for k in chaves_ordenadas if lote_xml[k]["nf_norm"] and lote_xml[k]["nf_norm"] in cand_csv),
            None
        )
        if match_chave:
            st.session_state["sb_nf_ativa"] = match_chave

if chaves_ordenadas:
    if st.session_state.get("sb_nf_ativa") not in lote_xml:
        st.session_state["sb_nf_ativa"] = chaves_ordenadas[0]
else:
    st.session_state.pop("sb_nf_ativa", None)

# Barra de gerenciamento do lote de XMLs
if lote_xml:
    with st.container(border=True):
        c_info, c_sel, c_rem, c_limp = st.columns([1.8, 3.0, 1.0, 1.1], vertical_alignment="center")
        opcoes_rotulo = {
            k: f"NF {lote_xml[k]['numero_nf']} — {lote_xml[k]['nome_arquivo']}"
            for k in chaves_ordenadas
        }

        with c_info:
            st.markdown(
                f'<span class="lote-pill">🗂️ {len(lote_xml)} XML(s) no lote</span>',
                unsafe_allow_html=True
            )

        with c_sel:
            st.selectbox(
                "NF ativa do lote",
                options=chaves_ordenadas,
                key="sb_nf_ativa",
                format_func=lambda k: opcoes_rotulo[k],
                label_visibility="collapsed"
            )

        with c_rem:
            st.button(
                "🗑️ Remover",
                key="btn_remover_nf",
                use_container_width=True,
                on_click=_remover_nf_ativa
            )

        with c_limp:
            st.button(
                "🧹 Limpar tudo",
                key="btn_limpar_lote",
                use_container_width=True,
                on_click=_limpar_lote
            )

# Estado padrão das flags de exportação
_DEFAULTS_FLAGS = {
    "flag_sel_todos": False,
    "flag_ident": False,
    "flag_qtd": False,
    "flag_ncm": False,
    "flag_ean": True,
    "flag_desc": True,
    "flag_icms": True,
    "flag_red": True,
    "flag_st": True,
}
for _k, _v in _DEFAULTS_FLAGS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


def _ao_clicar_selecionar_todos():
    marcar = st.session_state["flag_sel_todos"]
    for _k in ("flag_ident", "flag_qtd", "flag_ncm", "flag_ean", "flag_desc", "flag_icms", "flag_red", "flag_st"):
        st.session_state[_k] = marcar


# Seleção simétrica de blocos
with st.container(border=True):
    top_col1, top_col2 = st.columns([4, 1], vertical_alignment="center")
    with top_col1:
        st.markdown('<p class="section-label" style="margin-bottom:0;">⚙️ Blocos para Exportação</p>', unsafe_allow_html=True)
    with top_col2:
        st.checkbox("Selecionar Todos", key="flag_sel_todos", on_change=_ao_clicar_selecionar_todos)

    st.markdown("<hr style='margin: 6px 0 10px 0; border-color: #EEF2F6;'>", unsafe_allow_html=True)

    b_col1, b_col2, b_col3, b_col4 = st.columns(4, gap="medium")
    with b_col1:
        inc_ident = st.checkbox("Identificação", key="flag_ident")
        inc_qtd = st.checkbox("Quantidades", key="flag_qtd")
    with b_col2:
        inc_ean = st.checkbox("Código & EAN", key="flag_ean")
        inc_ncm = st.checkbox("NCM (Fiscal)", key="flag_ncm")
    with b_col3:
        inc_desc = st.checkbox("Descrições", key="flag_desc")
        inc_icms = st.checkbox("ICMS", key="flag_icms")
    with b_col4:
        inc_red = st.checkbox("Redução Base", key="flag_red")
        inc_st = st.checkbox("Cálculo ST", key="flag_st")

# Validação automática e botão de download centralizado
chave_ativa = st.session_state.get("sb_nf_ativa")
if lote_xml and arquivo_dados and chave_ativa in lote_xml:
    xml_selecionado = lote_xml[chave_ativa]
    try:
        numero_nf, itens_xml = ler_xml_nfe_bytes(xml_selecionado["bytes"])
    except ET.ParseError:
        st.markdown(
            f'<div class="v-badge err">🚫 O arquivo <b>{xml_selecionado["nome_arquivo"]}</b> não é um XML de NF-e válido.</div>',
            unsafe_allow_html=True
        )
        itens_xml = None

    if itens_xml is not None:
        df_csv, avisos_prep = ler_csv_winthor_cached(arquivo_dados.getvalue(), arquivo_dados.name)
        itens_csv = montar_itens_csv(df_csv)
        val = validar_upload(itens_xml, itens_csv, numero_nf, arquivo_dados.name)

        tem_bloqueio = len(val['bloqueios']) > 0
        ignorar_bloqueio_nf = False

        if tem_bloqueio or avisos_prep or val['avisos']:
            for bloq in val['bloqueios']:
                st.markdown(f'<div class="v-badge err">🚫 {bloq}</div>', unsafe_allow_html=True)

            for av in avisos_prep + val['avisos']:
                st.markdown(f'<div class="v-badge warn">⚠️ {av}</div>', unsafe_allow_html=True)

            if val['bloqueio_nf']:
                ignorar_bloqueio_nf = st.checkbox("Confirmar mesmo assim (o arquivo pertence a esta NF)")

        pode_processar = (not tem_bloqueio) or (val['bloqueio_nf'] and ignorar_bloqueio_nf)

        if pode_processar:
            dados_excel, nome_arquivo_saida = gerar_excel_auditoria_cached(
                _freeze_itens(itens_xml),
                _freeze_itens(itens_csv),
                numero_nf,
                inc_ident, inc_ean, inc_desc, inc_qtd, inc_ncm, inc_icms, inc_red, inc_st
            )

            st.markdown(f"""
                <div class="temp-notice-wrap">
                    <div class="temp-notice">
                        ✅ NF {numero_nf} pronta para auditoria
                    </div>
                </div>
            """, unsafe_allow_html=True)

            _, col_centro, _ = st.columns([1.4, 1.2, 1.4])
            with col_centro:
                st.download_button(
                    label="📥 Baixar Relatório",
                    data=dados_excel,
                    file_name=nome_arquivo_saida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
elif lote_xml and not arquivo_dados:
    st.markdown(
        '<div class="empty-hint">📄 Lote de XML pronto. Envie o relatório do <b>WinThor (CSV/XLSX)</b> no passo 2 para gerar a auditoria.</div>',
        unsafe_allow_html=True
    )
elif arquivo_dados and not lote_xml:
    st.markdown(
        '<div class="empty-hint">📂 Relatório do WinThor carregado. Envie os <b>XMLs das NF-e</b> no passo 1 para iniciar o cruzamento.</div>',
        unsafe_allow_html=True
    )
