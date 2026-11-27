import html
import io
import math
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
# CONFIGURAÇÃO DA PÁGINA E CSS (TEMA CLARO)
# ==========================================
st.set_page_config(page_title="Auditoria de Notas | Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .hero-header {
        background: linear-gradient(135deg, #1B2631 0%, #2C3E50 100%);
        color: #FFFFFF; padding: 1.15rem 1.5rem; border-radius: 12px;
        margin-bottom: 1.25rem; box-shadow: 0 2px 8px rgba(15, 23, 42, 0.08);
    }
    .hero-header h1 { margin: 0; font-size: 1.45rem; font-weight: 700; color: #FFFFFF; }
    .hero-header p { margin: 0.25rem 0 0 0; font-size: 0.9rem; color: #CBD5E1; }
    .section-label {
        font-size: 0.95rem; font-weight: 700; color: #1E293B; margin-bottom: 0.4rem;
        display: flex; align-items: center; gap: 0.45rem;
    }
    .step-num {
        display: inline-flex; align-items: center; justify-content: center;
        width: 22px; height: 22px; border-radius: 50%;
        background: #1B2631; color: #FFFFFF; font-size: 0.78rem; font-weight: 700;
    }
    .lote-pill {
        display: inline-block; background: #EFF6FF; color: #1D4ED8;
        border: 1px solid #BFDBFE; border-radius: 999px;
        padding: 0.3rem 0.85rem; font-size: 0.84rem; font-weight: 600;
    }
    .v-badge {
        padding: 0.65rem 0.95rem; border-radius: 8px; margin: 0.35rem 0;
        font-size: 0.88rem; font-weight: 500;
    }
    .v-badge.err { background: #FEE2E2; color: #991B1B; border: 1px solid #FCA5A5; }
    .v-badge.warn { background: #FEF3C7; color: #92400E; border: 1px solid #FCD34D; }
    .temp-notice-wrap { display: flex; justify-content: center; margin: 0.75rem 0 0.5rem 0; }
    .temp-notice {
        background: #DCFCE7; color: #14532D; border: 1px solid #86EFAC;
        border-radius: 999px; padding: 0.35rem 1rem; font-size: 0.88rem; font-weight: 600;
    }
    .empty-hint {
        text-align: center; padding: 1.25rem; margin-top: 1rem;
        background: #F8FAFC; border: 1px dashed #CBD5E1; border-radius: 10px;
        color: #475569; font-size: 0.92rem;
    }
    .div-ok-box {
        background: #DCFCE7; color: #14532D; border: 1px solid #86EFAC;
        border-radius: 10px; padding: 0.9rem 1.15rem; margin-top: 1rem;
        font-weight: 600; font-size: 0.92rem;
    }
    .div-panel {
        margin-top: 1.1rem; border: 1px solid #E2E8F0; border-radius: 10px;
        overflow: hidden; background: #FFFFFF;
    }
    .div-panel-header {
        background: #1B2631; color: #FFFFFF; padding: 0.7rem 1rem;
        display: flex; align-items: center; justify-content: space-between;
        font-weight: 700; font-size: 0.92rem;
    }
    .div-badges-wrap { display: flex; gap: 0.45rem; }
    .div-count-badge {
        background: #DC2626; color: #FFFFFF; border-radius: 999px;
        padding: 0.18rem 0.65rem; font-size: 0.78rem; font-weight: 700;
    }
    .div-warn-badge {
        background: #D97706; color: #FFFFFF; border-radius: 999px;
        padding: 0.18rem 0.65rem; font-size: 0.78rem; font-weight: 700;
    }
    .div-table-wrap { overflow-x: auto; }
    .div-table { width: 100%; border-collapse: collapse; font-size: 0.84rem; }
    .div-table th {
        background: #212F3D; color: #FFFFFF; padding: 0.55rem 0.75rem;
        text-align: center; font-weight: 600; white-space: nowrap;
    }
    .div-table th:first-child { text-align: left; }
    .div-table td {
        padding: 0.5rem 0.75rem; border-bottom: 1px solid #E2E8F0;
        text-align: center; vertical-align: middle;
    }
    .div-table tbody tr:nth-child(even) { background: #F8FAFC; }
    .cell-prod { text-align: left !important; font-weight: 600; color: #1B2631; white-space: nowrap; }
    .cell-desc { font-weight: 400; color: #475569; margin-left: 4px; }
    .cell-div { background: #FEE2E2; color: #991B1B; font-weight: 700; }
    .cell-warn { background: #FEF3C7; color: #B45309; font-weight: 700; }
    .cell-obs { background: #FFFBEB; color: #92400E; text-align: left !important; font-size: 0.8rem; }
    .cell-ok { color: #94A3B8; }
    .div-sep { color: #DC2626; margin: 0 5px; font-weight: 800; }
    .warn-sep { color: #D97706; margin: 0 5px; font-weight: 800; }
    .red-detail { display: block; font-size: 0.75rem; font-weight: 500; color: #7F1D1D; }
    .rodape-panel {
        margin-top: 1rem; border: 1px solid #CBD5E1; border-radius: 10px;
        overflow: hidden; background: #FFFFFF;
    }
    .rodape-header {
        background: #1B2631; color: #FFFFFF; padding: 0.65rem 1rem;
        font-weight: 700; font-size: 0.9rem;
    }
    .rodape-sec-title {
        background: #E2E8F0; color: #334155; padding: 0.4rem 0.9rem;
        font-size: 0.78rem; font-weight: 700; text-transform: uppercase;
    }
    .rodape-text {
        padding: 0.75rem 0.9rem; font-size: 0.84rem; color: #1E293B;
        white-space: pre-wrap; line-height: 1.5; background: #F8FAFC;
    }
    .rodape-highlight {
        background: #FDE047; color: #713F12; padding: 1px 3px;
        border-radius: 3px; font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# CONSTANTES E REGEX PRÉ-COMPILADAS
# ==========================================
NS = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
LIMITE_MARCA = 0.70

_RE_SCI = re.compile(r'^[0-9.,]+E[+-]?[0-9]+$')
_RE_DIGITS = re.compile(r'[^0-9]')
_RE_ALPHA = re.compile(r'[^a-z]')
_RE_TOKENS = re.compile(r'[A-Z0-9]+')
_RE_EXT = re.compile(r'\.[A-Za-z0-9]+$')
_RE_NUMS = re.compile(r'\d{3,}')
_RE_NNF_BYTES = re.compile(rb'<(?:[A-Za-z0-9]+:)?nNF>\s*(\d+)\s*</(?:[A-Za-z0-9]+:)?nNF>')
_RE_EMB = re.compile(r'\b(?:CX|DP|FD|PCT|EMB|PC|UN)X?(\d+)\b', re.IGNORECASE)
_IGNORAR = {'COM', 'PARA', 'DOS', 'DAS', 'UNID', 'UND', 'CX', 'KIT', 'PCT', 'FR', 'REF', 'ESC', 'ATIV'}

_PONTOS_IMPORTANTES_PADROES = [
    ("ICMS-ST / ST", re.compile(r'\b(?:icms[\s-]*st|st|b.?calc.?st)\b|substituicao\s+tributaria', re.IGNORECASE)),
    ("GNRE", re.compile(r'\bgnre\b|guia\s+nacional', re.IGNORECASE)),
    ("Redução de Base", re.compile(r'reducao\s+(?:de\s+|da\s+)?base|\bred.?\s*bc\b|base\s+de\s+calculo\s+reduzida', re.IGNORECASE)),
    ("Decreto / Convênio / Protocolo", re.compile(r'\b(?:decreto|convenio|protocolo|lei)\b', re.IGNORECASE)),
    ("FCP / FECP", re.compile(r'\b(?:fcp|fecp)\b', re.IGNORECASE)),
    ("DIFAL", re.compile(r'\bdifal\b|diferencial\s+de\s+aliquota', re.IGNORECASE)),
    ("Diferimento / Isenção / Benefício", re.compile(r'\b(?:diferimento|isencao|isento|beneficio|desonerado|desoneracao)\b', re.IGNORECASE)),
    ("Simples Nacional", re.compile(r'simples\s+nacional', re.IGNORECASE)),
    ("Crédito de ICMS", re.compile(r'credito\s+(?:de\s+)?icms|aproveitamento\s+(?:de\s+|do\s+)?credito', re.IGNORECASE)),
    ("Bonificação / Brinde", re.compile(r'\b(?:bonificacao|brinde)\b', re.IGNORECASE)),
    ("Refaturamento", re.compile(r'\b(?:refaturamento|refaturada|refatura|reemissao)\b', re.IGNORECASE)),
    (
        "Substituição / NF Anterior",
        re.compile(
            r'\b(?:substitui\s+nf|substituicao\s+de\s+nota|nota\s+anterior|nf\s+referenciada|'
            r'cancelamento|cancelada|estorno|nf\s+complementar)\b',
            re.IGNORECASE
        )
    ),
    ("Devolução / Retorno", re.compile(r'\b(?:devolucao|retorno|recusa)\b', re.IGNORECASE)),
    (
        "Avaria / Falta / Reclamação",
        re.compile(r'(?:^|[^a-zA-Z0-9])(?:avaria|falta|divergencia|reclamac[a-z]*|prazo\s+de\s+reclamacao)(?:$|[^a-zA-Z0-9])', re.IGNORECASE)
    ),
]

# Paleta padrão dos cabeçalhos Excel (XML x Sistema + linhas brancas)
COLOR_BLOCO_DARK = "1B2631"
COLOR_XML_SOLID = "212F3D"
COLOR_WIN_SOFT = "7F8C8D"
COLOR_STATUS_DIV = "9C0006"
COLOR_STATUS_WARN = "D97706"
COLOR_STATUS_OK = "196F3D"


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


def obter_sigla_origem(orig_cod) -> str:
    """Retorna sigla ultra-compacta de origem: NAC (Nacional) ou IMP (Importado)."""
    cod = str(orig_cod).strip()
    if cod in ('0', '3', '4', '5', '8'):
        return 'NAC'
    if cod in ('1', '2', '6', '7'):
        return 'IMP'
    return '-'


def extrair_fator_embalagem(uCom) -> float:
    """Extrai fator multiplicador apenas de siglas reais de embalagem (evita confundir ML/G/KG)."""
    if not uCom:
        return 1.0
    match = _RE_EMB.search(str(uCom).strip())
    if match:
        try:
            fator = float(match.group(1))
            if fator > 0:
                return fator
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


def _achar_coluna_tipo_calculo(colunas):
    for c in colunas:
        norm = _sem_acento(c).lower()
        if 'tipo' in norm and 'calculo' in norm:
            return c
    return None


def _achar_coluna_aliq_guia(colunas):
    for c in colunas:
        norm = _RE_ALPHA.sub('', _sem_acento(c).lower())
        if 'aliq' in norm and 'guia' in norm:
            return c
    return None


def _detectar_separador_csv(texto: str) -> str:
    """Detecta automaticamente se o CSV usa ';' ou ',' (prioriza ';' quando presente)."""
    for linha in texto.splitlines():
        l_strip = linha.strip()
        if l_strip:
            if ';' in l_strip:
                return ';'
            if ',' in l_strip:
                return ','
            break
    return ';'


def _validar_colunas_winthor(colunas) -> list[str]:
    """Verifica se todas as colunas essenciais do WinThor estão presentes."""
    col_set = set(colunas)
    obrigatorias = [
        'Cód. barra',
        'Descrição',
        'NCM',
        '%ICMS',
        '% ICMS red.',
    ]
    faltantes = [c for c in obrigatorias if c not in col_set]
    if _achar_coluna_aliq_guia(colunas) is None:
        faltantes.append('%Alíq.Ext.ST Guia')
    if _achar_coluna_tipo_calculo(colunas) is None:
        faltantes.append('Tipo cálculo ST')
    return faltantes


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
        sep_detectado = _detectar_separador_csv(texto)
        df = pd.read_csv(io.StringIO(texto), sep=sep_detectado, dtype=str)

    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how='all').reset_index(drop=True)

    colunas_faltantes = _validar_colunas_winthor(df.columns)

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

    return df, avisos, colunas_faltantes


def montar_itens_csv(df):
    col_tipo = _achar_coluna_tipo_calculo(df.columns)
    col_guia = _achar_coluna_aliq_guia(df.columns)

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
            '%Alíq.Ext. ST Guia_Whintor': converte_float_br(row.get(col_guia, 0.0)) if col_guia else 0.0,
            'Tipo_Calculo_WINTHOR': str(row.get(col_tipo, '')).strip() if col_tipo else '',
        })
    return itens


@st.cache_data(show_spinner=False)
def ler_xml_nfe_bytes(bruto: bytes):
    root = ET.fromstring(bruto)

    nNF = root.find('.//ns:ide/ns:nNF', NS)
    numero_nf = nNF.text.strip() if nNF is not None and nNF.text else 'DESCONHECIDO'

    id_dest_node = root.find('.//ns:ide/ns:idDest', NS)
    id_dest = id_dest_node.text.strip() if id_dest_node is not None and id_dest_node.text else ''

    def txt(no, tag):
        if no is None:
            return ''
        e = no.find(f'ns:{tag}', NS)
        return e.text.strip() if e is not None and e.text else ''

    inf_adic = root.find('.//ns:infAdic', NS)
    dados_adicionais = {
        'infAdFisco': txt(inf_adic, 'infAdFisco'),
        'infCpl': txt(inf_adic, 'infCpl'),
        'idDest': id_dest,
    }

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

        icms, origem, v_icms_st = 0.0, 'NAC', 0.0
        icms_node = det.find('.//ns:ICMS', NS)
        if icms_node is not None and len(icms_node) > 0:
            sub = icms_node[0]
            p = sub.find('ns:pICMS', NS)
            if p is not None and p.text:
                icms = converte_float_br(p.text)
            o = sub.find('ns:orig', NS)
            if o is not None and o.text:
                origem = obter_sigla_origem(o.text.strip())
            st_val_node = sub.find('ns:vICMSST', NS)
            if st_val_node is not None and st_val_node.text:
                v_icms_st = converte_float_br(st_val_node.text)

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
            'Orig.': origem,
            'ICMS_NF': icms,
            'pRedBC_NF': pred,
            '_vICMSST_NF': v_icms_st,
            '_idDest': id_dest,
        })
    return numero_nf, itens, dados_adicionais


def _texto_e_mapa_normalizado(texto: str) -> tuple[str, list[int]]:
    """Normaliza acentos para regex preservando o índice exato na string original."""
    chars_norm = []
    mapa_idx = []
    for idx, ch in enumerate(texto):
        decomposed = unicodedata.normalize('NFKD', ch)
        for dch in decomposed:
            if not unicodedata.combining(dch):
                chars_norm.append(dch)
                mapa_idx.append(idx)
    return ''.join(chars_norm), mapa_idx


def _expandir_trecho_inteiro(texto: str, start_idx: int, end_idx: int) -> tuple[int, int]:
    """Expande o intervalo [start_idx, end_idx) até os delimitadores da frase/trecho nos dois lados."""
    separadores = set('\n\r;•|')
    n = len(texto)

    esq = start_idx
    while esq > 0:
        if texto[esq - 1] in separadores:
            break
        if esq >= 3 and texto[esq - 3:esq] == ' - ':
            break
        esq -= 1

    dir_pos = end_idx
    while dir_pos < n:
        ch = texto[dir_pos]
        if ch in separadores:
            break
        if ch == '-' and dir_pos > 0 and dir_pos + 1 < n and texto[dir_pos - 1] == ' ' and texto[dir_pos + 1] == ' ':
            break
        dir_pos += 1

    while esq < dir_pos and texto[esq] in (' ', '-', '\t'):
        esq += 1

    while dir_pos > esq and texto[dir_pos - 1] in (' ', '\t'):
        dir_pos -= 1

    return esq, dir_pos


def destacar_trechos_rodape_html(texto: str, ponto_ativo: str | None) -> str:
    """Retorna HTML seguro com os trechos/frases inteiras correspondentes ao ponto destacado."""
    if not texto:
        return ''
    if not ponto_ativo:
        return html.escape(texto)

    padrao = next((p for rotulo, p in _PONTOS_IMPORTANTES_PADROES if rotulo == ponto_ativo), None)
    if padrao is None:
        return html.escape(texto)

    texto_norm, mapa_idx = _texto_e_mapa_normalizado(texto)
    if not mapa_idx:
        return html.escape(texto)

    intervalos = []
    for m in padrao.finditer(texto_norm):
        s_norm, e_norm = m.span()
        if s_norm >= len(mapa_idx):
            continue
        orig_start = mapa_idx[s_norm]
        orig_end = mapa_idx[min(e_norm - 1, len(mapa_idx) - 1)] + 1
        t_start, t_end = _expandir_trecho_inteiro(texto, orig_start, orig_end)
        if t_start < t_end:
            intervalos.append((t_start, t_end))

    if not intervalos:
        return html.escape(texto)

    intervalos.sort()
    mesclados = [intervalos[0]]
    for s, e in intervalos[1:]:
        prev_s, prev_e = mesclados[-1]
        if s <= prev_e + 1:
            mesclados[-1] = (prev_s, max(prev_e, e))
        else:
            mesclados.append((s, e))

    partes = []
    cursor = 0
    for s, e in mesclados:
        if cursor < s:
            partes.append(html.escape(texto[cursor:s]))
        partes.append(f'<mark class="rodape-highlight">{html.escape(texto[s:e])}</mark>')
        cursor = e
    if cursor < len(texto):
        partes.append(html.escape(texto[cursor:]))

    return ''.join(partes)


def analisar_rodape_nfe(dados_adicionais: dict | None) -> dict:
    """Analisa infAdFisco e infCpl e identifica pontos importantes no rodapé."""
    if not dados_adicionais:
        return {'tem_texto': False, 'infAdFisco': '', 'infCpl': '', 'alertas': []}

    fisco = (dados_adicionais.get('infAdFisco') or '').strip()
    cpl = (dados_adicionais.get('infCpl') or '').strip()
    texto_completo = f"{fisco}\n{cpl}".strip()
    if not texto_completo:
        return {'tem_texto': False, 'infAdFisco': '', 'infCpl': '', 'alertas': []}

    texto_norm, _ = _texto_e_mapa_normalizado(texto_completo)
    alertas = [rotulo for rotulo, padrao in _PONTOS_IMPORTANTES_PADROES if padrao.search(texto_norm)]

    return {
        'tem_texto': True,
        'infAdFisco': fisco,
        'infCpl': cpl,
        'alertas': alertas,
    }


# ==========================================
# 2. VALIDAÇÃO, CRUZAMENTO E DIAGNÓSTICOS
# ==========================================
def _tokens(desc):
    t = _RE_TOKENS.findall(_sem_acento(desc).upper())
    return {x for x in t if len(x) >= 3 and not any(ch.isdigit() for ch in x) and x not in _IGNORAR}


def _numeros_do_nome(nome):
    stem = _RE_EXT.sub('', str(nome))
    return {g.lstrip('0') for g in _RE_NUMS.findall(stem)}


def validar_upload(itens_xml, itens_csv, numero_nf, nome_arquivo_csv, colunas_faltantes=None):
    bloqueios, avisos, infos = [], [], []
    r = {
        'bloqueios': bloqueios,
        'avisos': avisos,
        'infos': infos,
        'marca_em_comum': [],
        'bloqueio_nf': False,
    }

    if colunas_faltantes:
        lista_cols = ', '.join(f"'{c}'" for c in colunas_faltantes)
        bloqueios.append(
            f"Coluna(s) obrigatória(s) ausente(s) no arquivo do WinThor: {lista_cols}. "
            f"Exporte o relatório completo antes de auditar."
        )
        return r

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
            eantrib_x = ix.get('EANTrib_Comp_NF', '')
            ean_c = ic.get('EAN_Comp_WINTHOR', '')
            mesmo_ean = bool(ean_c and (ean_c == ean_x or ean_c == eantrib_x))

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


def _fmt_num_curto(valor, max_dec: int = 2) -> str:
    try:
        v = float(valor)
        if v.is_integer():
            return str(int(v))
        return f"{v:.{max_dec}f}".rstrip('0').rstrip('.')
    except (TypeError, ValueError):
        return str(valor)


def _fmt_num_br(valor, max_dec: int = 2) -> str:
    return _fmt_num_curto(valor, max_dec=max_dec).replace('.', ',')


def _descricao_origem(orig_sigla: str) -> str:
    sigla = str(orig_sigla or '').strip().upper()
    if sigla == 'NAC':
        return 'origem nacional'
    if sigla == 'IMP':
        return 'origem importada'
    return ''


def cruzar_itens_auditoria(itens_xml, itens_csv):
    """Realiza o cruzamento linha a linha entre XML e WinThor (fonte única para tela e Excel)."""
    dados_relatorio = []
    for item_x, item_c in zip_longest(itens_xml, itens_csv, fillvalue=None):
        linha = {}

        if item_x:
            linha.update(item_x)
        else:
            linha.update({
                'Linha_NF': '-', 'Descricao_NF': '-', 'EAN_NF': '-', 'NCM_NF': '-',
                'Qtd_NF': '-', '_Fator_Embalagem': 1.0, 'Vl_Unit_NF': '-', 'Vl_Total_NF': '-',
                'Orig.': '-', 'ICMS_NF': '-', 'pRedBC_NF': '-'
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

        linha['EAN_Alerta_Vazio'] = False
        linha['ICMS_Alerta_Guia'] = False
        linha['ICMS_Obs'] = ''

        if item_x and item_c:
            linha['Status_Geral'] = 'PAREADO'

            # 1. Conferência de EAN (com alerta laranja quando ambos vazios)
            ean_x_comp = item_x.get('EAN_Comp_NF', '')
            eantrib_x_comp = item_x.get('EANTrib_Comp_NF', '')
            ean_win_comp = item_c.get('EAN_Comp_WINTHOR', '')

            if ean_win_comp != '' and (ean_win_comp == ean_x_comp or ean_win_comp == eantrib_x_comp):
                linha['Status_EAN'] = 'OK'
            elif ean_win_comp == '' and ean_x_comp == '' and eantrib_x_comp == '':
                linha['Status_EAN'] = 'ALERTA'
                linha['EAN_Alerta_Vazio'] = True
                linha['EAN_NF'] = 'SEM EAN'
                linha['EAN_WINTHOR'] = 'SEM EAN'
            else:
                linha['Status_EAN'] = 'DIVERGENTE'

            # 2. Tipo de Cálculo ST
            tipo_calc = str(item_c.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
            if tipo_calc == 'N':
                linha['Tipo_Calculo_ST'] = 'NOTA FISCAL'
            elif tipo_calc == 'G':
                linha['Tipo_Calculo_ST'] = 'GUIA'
            else:
                linha['Tipo_Calculo_ST'] = item_c['Tipo_Calculo_WINTHOR']

            # 3. Conferência de ICMS guiada pelo Tipo de Cálculo ST
            icms_xml = round(item_x['ICMS_NF'], 2)
            icms_csv = round(item_c['ICMS_WINTHOR'], 2)
            st_guia = round(item_c['%Alíq.Ext. ST Guia_Whintor'], 2)

            if tipo_calc == 'G':
                if icms_xml != st_guia:
                    linha['Status_ICMS_Aliq'] = 'DIVERGENTE'
                elif icms_csv == 0.0:
                    linha['Status_ICMS_Aliq'] = 'OK'
                else:
                    linha['Status_ICMS_Aliq'] = 'ALERTA'
                    linha['ICMS_Alerta_Guia'] = True
                    linha['ICMS_Obs'] = (
                        f"Item Guia (G) com %ICMS preenchido ({_fmt_num_br(icms_csv)}%) "
                        f"no WinThor — zerar ICMS."
                    )
            else:
                if icms_xml == icms_csv:
                    linha['Status_ICMS_Aliq'] = 'OK'
                else:
                    linha['Status_ICMS_Aliq'] = 'DIVERGENTE'

            # 4. Conferência de Redução de Base (ICMS Efetivo XML x % ICMS red. WinThor)
            pred_xml = round(item_x['pRedBC_NF'], 4)
            red_csv = round(item_c['% ICMS red. Whintor'], 4)
            carga_efetiva_xml = round(item_x['ICMS_NF'] * (1.0 - (item_x['pRedBC_NF'] / 100.0)), 4)
            linha['ICMS_Efetivo_NF'] = carga_efetiva_xml

            bate_efetivo = (
                carga_efetiva_xml == red_csv
                or round(carga_efetiva_xml, 2) == round(red_csv, 2)
            )
            protecao_guia = (tipo_calc == 'G' and pred_xml == 0.0 and red_csv == 0.0)

            if bate_efetivo or protecao_guia:
                linha['Status_Red_BC'] = 'OK'
            else:
                linha['Status_Red_BC'] = 'DIVERGENTE'

            # 5. Conferência de NCM (padronizado em 8 dígitos)
            ncm_x = str(item_x['NCM_NF']).strip()
            ncm_w = str(item_c['NCM_WINTHOR']).strip()
            ncm_x_norm = ncm_x.zfill(8) if ncm_x else ''
            ncm_w_norm = ncm_w.zfill(8) if ncm_w else ''
            linha['Status_NCM'] = 'OK' if ncm_x_norm == ncm_w_norm else 'DIVERGENTE'

            # 6. Conferência de Quantidade
            qtd_xml_pura = item_x['Qtd_NF']
            qtd_win = item_c['Qtd_WINTHOR']
            fator = item_x['_Fator_Embalagem']
            qtd_xml_convertida = qtd_xml_pura * fator

            if round(qtd_xml_pura, 4) == round(qtd_win, 4) or round(qtd_xml_convertida, 4) == round(qtd_win, 4):
                linha['Status_Qtd'] = 'OK'
            else:
                linha['Status_Qtd'] = 'DIVERGENTE'

            linha['Alerta_Origem_Aliq'] = ''
            linha['Alerta_Guia_Coerencia'] = ''
            id_dest_item = str(item_x.get('_idDest', '')).strip()
            v_st_item = converte_float_br(item_x.get('_vICMSST_NF', 0.0))
            origem = str(item_x.get('Orig.', 'NAC')).strip().upper()

            # Alerta de coerência Origem x Alíquota (operação interestadual idDest == '2')
            if id_dest_item == '2' and icms_xml > 0.0:
                if origem == 'IMP' and icms_xml != 4.0:
                    linha['Alerta_Origem_Aliq'] = (
                        f"Item IMP em operação interestadual com ICMS {_fmt_num_br(icms_xml)}% "
                        f"(esperado 4% — validar com tributário)."
                    )
                elif origem == 'NAC' and icms_xml == 4.0:
                    linha['Alerta_Origem_Aliq'] = (
                        f"Item NAC em operação interestadual com ICMS 4% "
                        f"(alíquota de importado — validar com tributário)."
                    )

            # Coerência de Guia: ST já destacada na NF mas marcado como GUIA (G) no WinThor
            if tipo_calc == 'G' and v_st_item > 0.0:
                linha['Alerta_Guia_Coerencia'] = (
                    f"Item com ST já destacada na NF (vICMSST R$ {v_st_item:.2f}) "
                    f"marcado como GUIA (G) no WinThor — validar com tributário."
                )

        else:
            linha['Status_Geral'] = 'SOBRA'
            linha['ICMS_Efetivo_NF'] = '-'
            linha['Alerta_Origem_Aliq'] = ''
            linha['Alerta_Guia_Coerencia'] = ''
            for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                linha[st_col] = '-'

        dados_relatorio.append(linha)

    # Detectar inconsistência G vs N no mesmo grupo (NCM + família de descrição)
    def _chave_familia(desc_txt: str) -> str:
        norm = _sem_acento(str(desc_txt or '')).upper()
        tokens = [
            t for t in _RE_TOKENS.findall(norm)
            if len(t) >= 3 and not any(ch.isdigit() for ch in t) and t not in _IGNORAR
        ]
        return ' '.join(tokens[:2])

    grupos_tipo = {}
    for idx_l, l_item in enumerate(dados_relatorio):
        if l_item.get('Status_Geral') != 'PAREADO':
            continue
        t_calc = str(l_item.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
        if t_calc not in ('G', 'N'):
            continue
        fam = _chave_familia(l_item.get('Descricao_NF') or l_item.get('Descricao_WINTHOR') or '')
        ncm_raw = str(l_item.get('NCM_NF') or l_item.get('NCM_WINTHOR') or '').strip()
        ncm_k = ncm_raw.zfill(8) if ncm_raw and ncm_raw.strip('0') else ''
        if ncm_k and fam:
            chave_grp = f"NCM {ncm_k} • {fam}"
        elif ncm_k:
            chave_grp = f"NCM {ncm_k}"
        elif len(fam) >= 5:
            chave_grp = fam
        else:
            chave_grp = ''
        if chave_grp:
            grupos_tipo.setdefault(chave_grp, []).append((idx_l, t_calc))

    for chave_grp, membros in grupos_tipo.items():
        qtd_g = sum(1 for _, tc in membros if tc == 'G')
        qtd_n = sum(1 for _, tc in membros if tc == 'N')
        if qtd_g > 0 and qtd_n > 0:
            tipo_min = 'G' if qtd_g < qtd_n else ('N' if qtd_n < qtd_g else None)
            for idx_l, tc in membros:
                if tipo_min is None or tc == tipo_min:
                    rotulo_tc = 'GUIA (G)' if tc == 'G' else 'NOTA FISCAL (N)'
                    msg_grupo = (
                        f"Grupo '{chave_grp}' mistura GUIA e NOTA ({qtd_g} G / {qtd_n} N) "
                        f"— este item está como {rotulo_tc}, verificar."
                    )
                    prev = dados_relatorio[idx_l].get('Alerta_Guia_Coerencia', '')
                    dados_relatorio[idx_l]['Alerta_Guia_Coerencia'] = f"{prev} | {msg_grupo}" if prev else msg_grupo

    return dados_relatorio


def _diagnostico_sobra(linha: dict) -> str:
    if linha.get('Linha_NF') != '-':
        return (
            "Item presente apenas no XML (sem correspondência no WinThor). "
            "Ação: verificar se faltou incluir o item na pré-entrada ou se houve erro de sequência."
        )
    return (
        "Item presente apenas no WinThor (sem correspondência no XML). "
        "Ação: verificar se o item foi lançado indevidamente na pré-entrada."
    )


def _diagnostico_ean(linha: dict) -> str:
    st_ean = linha.get('Status_EAN')
    if st_ean == 'DIVERGENTE':
        ean_x = linha.get('EAN_NF') or 'vazio'
        ean_w = linha.get('EAN_WINTHOR') or 'vazio'
        return (
            f"Divergência de EAN: NF veio com '{ean_x}', mas WinThor está com '{ean_w}'. "
            f"Ação: conferir o código de barras físico do produto e atualizar 'Cód. barra' no WinThor."
        )
    if linha.get('EAN_Alerta_Vazio'):
        return (
            "Produto sem código de barras no XML e no WinThor. "
            "Ação: validar se o item não possui GTIN ou preencher 'Cód. barra' no cadastro do WinThor."
        )
    return ""


def _diagnostico_ncm(linha: dict) -> str:
    if linha.get('Status_NCM') == 'DIVERGENTE':
        ncm_x = str(linha.get('NCM_NF') or '').strip()
        ncm_w = str(linha.get('NCM_WINTHOR') or '').strip()
        ncm_x_fmt = ncm_x.zfill(8) if ncm_x else '-'
        ncm_w_fmt = ncm_w.zfill(8) if ncm_w else '-'
        return (
            f"Divergência de NCM: NF veio com {ncm_x_fmt}, mas WinThor está com {ncm_w_fmt}. "
            f"Ação: validar a classificação fiscal correta com o setor tributário e ajustar o NCM no WinThor."
        )
    return ""


def _diagnostico_icms_basered(linha: dict) -> list[str]:
    motivos = []
    st_icms = linha.get('Status_ICMS_Aliq')
    st_red = linha.get('Status_Red_BC')
    tipo_calc = str(linha.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
    orig_desc = _descricao_origem(linha.get('Orig.', ''))
    orig_par = f" ({orig_desc})" if orig_desc else ""

    i_x_val = converte_float_br(linha.get('ICMS_NF', 0.0))
    i_w_val = converte_float_br(linha.get('ICMS_WINTHOR', 0.0))
    i_g_val = converte_float_br(linha.get('%Alíq.Ext. ST Guia_Whintor', 0.0))
    pred_x_val = converte_float_br(linha.get('pRedBC_NF', 0.0))
    ef_x_val = converte_float_br(linha.get('ICMS_Efetivo_NF', 0.0))
    r_w_val = converte_float_br(linha.get('% ICMS red. Whintor', 0.0))

    i_x = _fmt_num_br(i_x_val, max_dec=2)
    i_w = _fmt_num_br(i_w_val, max_dec=2)
    i_g = _fmt_num_br(i_g_val, max_dec=2)
    pred_x = _fmt_num_br(pred_x_val, max_dec=2)
    ef_x = _fmt_num_br(ef_x_val, max_dec=2)
    r_w = _fmt_num_br(r_w_val, max_dec=2)

    if st_icms == 'DIVERGENTE':
        if tipo_calc == 'G':
            efeito_guia = (
                "para não subestimar a guia" if i_x_val > i_g_val else "para não superestimar a guia"
            )
            motivos.append(
                f"Divergência de Alíquota Guia: NF veio com {i_x}%{orig_par}, mas WinThor está com {i_g}%. "
                f"Ação: ajustar '%Alíq.Ext.ST Guia' para {i_x}% na pré-entrada {efeito_guia}."
            )
            if round(i_w_val, 2) != 0.0:
                motivos.append(
                    f"Item marcado como GUIA (G) com %ICMS preenchido ({i_w}%). "
                    f"Ação: zerar '%ICMS' no WinThor para evitar crédito indevido junto com cálculo de guia."
                )
        else:
            motivos.append(
                f"Divergência de Alíquota de ICMS: NF veio com {i_x}%{orig_par}, mas WinThor está com {i_w}%. "
                f"Ação: ajustar '%ICMS' para {i_x}% na pré-entrada para refletir corretamente o crédito da nota."
            )
    elif linha.get('ICMS_Alerta_Guia'):
        motivos.append(
            f"Item marcado como GUIA (G) com %ICMS preenchido ({i_w}%). "
            f"Ação: zerar '%ICMS' no WinThor para evitar crédito indevido junto com cálculo de guia."
        )

    if st_red == 'DIVERGENTE':
        if pred_x_val > 0.0:
            motivos.append(
                f"Redução de Base divergente: NF tem ICMS {i_x}% com redução de {pred_x}% "
                f"(carga efetiva = {ef_x}%), mas WinThor está com {r_w}%. "
                f"Ação: ajustar '% ICMS red.' para {ef_x}%."
            )
        else:
            motivos.append(
                f"Redução de Base divergente: NF tem ICMS {i_x}% sem redução "
                f"(carga efetiva = {ef_x}%), mas WinThor está com {r_w}%. "
                f"Ação: ajustar '% ICMS red.' para {ef_x}%."
            )

    if linha.get('Alerta_Origem_Aliq'):
        motivos.append(
            f"Alerta de Origem × Alíquota: {linha['Alerta_Origem_Aliq']} "
            f"Ação: confirmar a origem da mercadoria e a alíquota interestadual com o setor tributário."
        )

    if linha.get('Alerta_Guia_Coerencia'):
        motivos.append(
            f"Alerta de Tipo Cálculo ST: {linha['Alerta_Guia_Coerencia']} "
            f"Ação: confirmar com o fiscal se o item deve entrar como GUIA (G) ou NOTA FISCAL (N)."
        )

    return motivos


def _gerar_diagnostico_escrito(linha: dict) -> str:
    """Gera explicação consolidada no formato Causa + Ação para o Excel."""
    if linha.get('Status_Geral') == 'SOBRA':
        return _diagnostico_sobra(linha)

    motivos = []
    d_ean = _diagnostico_ean(linha)
    if d_ean:
        motivos.append(d_ean)

    d_ncm = _diagnostico_ncm(linha)
    if d_ncm:
        motivos.append(d_ncm)

    motivos.extend(_diagnostico_icms_basered(linha))
    return "\n".join(motivos)


def _calcular_altura_linha(textos_e_larguras: list[tuple[str, float]], altura_min: int = 24) -> int:
    """Calcula a altura ideal da linha no Excel com base nos textos com wrap_text."""
    max_linhas = 1
    for txt, largura in textos_e_larguras:
        if not txt:
            continue
        chars_linha = max(int(largura) - 3, 18)
        linhas_txt = 0
        for paragrafo in str(txt).splitlines():
            linhas_txt += max(1, math.ceil(len(paragrafo) / chars_linha))
        max_linhas = max(max_linhas, linhas_txt)
    return max(altura_min, min(max_linhas * 15 + 10, 260))


def _rotulo_ean_desc(linha: dict, max_desc: int = 38) -> tuple[str, str, str]:
    ean_nf = str(linha.get('EAN_NF') or '').strip()
    if ean_nf and ean_nf not in ('-', 'SEM EAN'):
        rotulo_ean = ean_nf
    else:
        rotulo_ean = f"Sem EAN (Item {linha.get('Linha_NF', '-')})"

    desc_completa = str(linha.get('Descricao_NF') or linha.get('Descricao_WINTHOR') or '').strip()
    desc_resumida = (desc_completa[:max_desc].rstrip() + '…') if len(desc_completa) > max_desc else desc_completa
    texto_ean_desc = f"{rotulo_ean} — {desc_resumida}" if desc_resumida else rotulo_ean
    return rotulo_ean, desc_resumida, texto_ean_desc


def preparar_dados_resumo_divergencias(dados_relatorio):
    """Extrai registros de divergências e alertas laranja (fonte única para tela e Excel)."""
    registros = []
    colunas_encontradas = {
        'EAN': False,
        'NCM': False,
        'ICMS': False,
        'Base Reduzida': False,
        'Tipo Cálculo': False,
        'Sobra': False,
        'Observação': False,
    }
    qtd_divergencias = 0
    qtd_alertas = 0

    for linha in dados_relatorio:
        celulas_texto = {}
        celulas_html = {}
        celulas_tipo = {}
        observacoes = []
        tem_div = False
        tem_warn = False

        if linha.get('Status_Geral') == 'SOBRA':
            origem_sobra = 'XML' if linha.get('Linha_NF') != '-' else 'WinThor'
            txt_sobra = f'Apenas {origem_sobra}'
            celulas_texto['Sobra'] = txt_sobra
            celulas_html['Sobra'] = txt_sobra
            celulas_tipo['Sobra'] = 'div'
            colunas_encontradas['Sobra'] = True
            tem_div = True
        else:
            if linha.get('Status_EAN') == 'DIVERGENTE':
                ean_x = str(linha.get('EAN_NF') or 'vazio')
                ean_w = str(linha.get('EAN_WINTHOR') or 'vazio')
                celulas_texto['EAN'] = f'{ean_x} × {ean_w}'
                celulas_html['EAN'] = f'{html.escape(ean_x)}<span class="div-sep">×</span>{html.escape(ean_w)}'
                celulas_tipo['EAN'] = 'div'
                colunas_encontradas['EAN'] = True
                tem_div = True
            elif linha.get('EAN_Alerta_Vazio'):
                celulas_texto['EAN'] = 'SEM EAN × SEM EAN'
                celulas_html['EAN'] = 'SEM EAN<span class="warn-sep">×</span>SEM EAN'
                celulas_tipo['EAN'] = 'warn'
                colunas_encontradas['EAN'] = True
                observacoes.append('Produto sem código de barras no XML e no WinThor — verificar.')
                tem_warn = True

            if linha.get('Status_NCM') == 'DIVERGENTE':
                n_x = str(linha.get('NCM_NF') or '-')
                n_w = str(linha.get('NCM_WINTHOR') or '-')
                celulas_texto['NCM'] = f'{n_x} × {n_w}'
                celulas_html['NCM'] = f'{html.escape(n_x)}<span class="div-sep">×</span>{html.escape(n_w)}'
                celulas_tipo['NCM'] = 'div'
                colunas_encontradas['NCM'] = True
                tem_div = True

            tipo_calc = str(linha.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
            if linha.get('Status_ICMS_Aliq') == 'DIVERGENTE':
                i_x = _fmt_num_curto(linha.get('ICMS_NF'))
                if tipo_calc == 'G':
                    i_g = _fmt_num_curto(linha.get('%Alíq.Ext. ST Guia_Whintor'))
                    celulas_texto['ICMS'] = f'{i_x}% × Guia {i_g}%'
                    celulas_html['ICMS'] = f'{i_x}%<span class="div-sep">×</span>Guia {i_g}%'
                else:
                    i_w = _fmt_num_curto(linha.get('ICMS_WINTHOR'))
                    celulas_texto['ICMS'] = f'{i_x}% × {i_w}%'
                    celulas_html['ICMS'] = f'{i_x}%<span class="div-sep">×</span>{i_w}%'
                celulas_tipo['ICMS'] = 'div'
                colunas_encontradas['ICMS'] = True
                tem_div = True
            elif linha.get('ICMS_Alerta_Guia'):
                i_x = _fmt_num_curto(linha.get('ICMS_NF'))
                i_w = _fmt_num_curto(linha.get('ICMS_WINTHOR'))
                celulas_texto['ICMS'] = f'Guia {i_x}% • ICMS {i_w}% (zerar ICMS)'
                celulas_html['ICMS'] = f'Guia {i_x}% <span class="warn-sep">•</span> ICMS {i_w}% (zerar ICMS)'
                celulas_tipo['ICMS'] = 'warn'
                colunas_encontradas['ICMS'] = True
                if linha.get('ICMS_Obs'):
                    observacoes.append(linha['ICMS_Obs'])
                tem_warn = True

            if linha.get('Status_Red_BC') == 'DIVERGENTE':
                ef_x = _fmt_num_curto(linha.get('ICMS_Efetivo_NF'), max_dec=4)
                r_w = _fmt_num_curto(linha.get('% ICMS red. Whintor'), max_dec=4)
                icms_x = _fmt_num_curto(linha.get('ICMS_NF'), max_dec=2)
                pred_x = _fmt_num_curto(linha.get('pRedBC_NF'), max_dec=4)
                celulas_texto['Base Reduzida'] = f'{ef_x}% × {r_w}% (ICMS {icms_x}% | Red. {pred_x}%)'
                celulas_html['Base Reduzida'] = (
                    f'{ef_x}%<span class="div-sep">×</span>{r_w}%'
                    f'<span class="red-detail">(ICMS {icms_x}% | Red. {pred_x}%)</span>'
                )
                celulas_tipo['Base Reduzida'] = 'div'
                colunas_encontradas['Base Reduzida'] = True
                tem_div = True

            if linha.get('Alerta_Origem_Aliq'):
                orig_txt = str(linha.get('Orig.') or 'IMP')
                aliq_txt = _fmt_num_curto(linha.get('ICMS_NF'))
                if 'ICMS' not in celulas_texto:
                    celulas_texto['ICMS'] = f'{orig_txt} {aliq_txt}% (validar)'
                    celulas_html['ICMS'] = f'{html.escape(orig_txt)} {aliq_txt}% <span class="warn-sep">•</span> validar'
                    celulas_tipo['ICMS'] = 'warn'
                    colunas_encontradas['ICMS'] = True
                observacoes.append(linha['Alerta_Origem_Aliq'])
                tem_warn = True

            if linha.get('Alerta_Guia_Coerencia'):
                tc_lbl = 'GUIA (G)' if tipo_calc == 'G' else 'NOTA FISCAL (N)'
                celulas_texto['Tipo Cálculo'] = f'{tc_lbl} — verificar'
                celulas_html['Tipo Cálculo'] = f'{html.escape(tc_lbl)} <span class="warn-sep">•</span> verificar'
                celulas_tipo['Tipo Cálculo'] = 'warn'
                colunas_encontradas['Tipo Cálculo'] = True
                observacoes.append(linha['Alerta_Guia_Coerencia'])
                tem_warn = True

        if observacoes:
            txt_obs = ' | '.join(observacoes)
            celulas_texto['Observação'] = txt_obs
            celulas_html['Observação'] = html.escape(txt_obs)
            celulas_tipo['Observação'] = 'obs'
            colunas_encontradas['Observação'] = True

        if celulas_texto:
            if tem_div:
                qtd_divergencias += 1
            elif tem_warn:
                qtd_alertas += 1

            cod_prod = str(linha.get('Cod_Produto_WINTHOR') or '').strip()
            cod_prod = cod_prod if cod_prod and cod_prod != '-' else '-'

            rotulo_ean, desc_resumida, texto_ean_desc = _rotulo_ean_desc(linha)
            desc_completa = str(linha.get('Descricao_NF') or linha.get('Descricao_WINTHOR') or '').strip()
            desc_html = f'<span class="cell-desc">— {html.escape(desc_resumida)}</span>' if desc_resumida else ''

            registros.append({
                'cod_produto': cod_prod,
                'ean': html.escape(rotulo_ean),
                'desc_html': desc_html,
                'texto_ean_desc': texto_ean_desc,
                'tooltip': html.escape(desc_completa),
                'celulas_texto': celulas_texto,
                'celulas_html': celulas_html,
                'celulas_tipo': celulas_tipo,
                'diagnostico': _gerar_diagnostico_escrito(linha),
                'tem_div': tem_div,
                'tem_warn': tem_warn,
            })

    ordem_colunas = ['EAN', 'NCM', 'ICMS', 'Base Reduzida', 'Tipo Cálculo', 'Sobra', 'Observação']
    colunas_ativas = [c for c in ordem_colunas if colunas_encontradas[c]]
    return registros, colunas_ativas, qtd_divergencias, qtd_alertas


# ==========================================
# 3. RENDERIZAÇÃO EM TELA (STREAMLIT)
# ==========================================
def _alternar_ponto_rodape(ponto_rotulo: str):
    atual = st.session_state.get("ponto_rodape_ativo")
    st.session_state["ponto_rodape_ativo"] = None if atual == ponto_rotulo else ponto_rotulo


def renderizar_rodape_nfe(dados_adicionais: dict | None):
    """Renderiza o bloco de Dados Adicionais / Rodapé da NF-e na tela com tags interativas."""
    info = analisar_rodape_nfe(dados_adicionais)
    if not info['tem_texto']:
        return

    ponto_ativo = st.session_state.get("ponto_rodape_ativo")
    if ponto_ativo and ponto_ativo not in info['alertas']:
        ponto_ativo = None
        st.session_state["ponto_rodape_ativo"] = None

    st.markdown(
        '<div class="rodape-panel">'
        '  <div class="rodape-header">'
        '    <span>📌 Dados Adicionais / Rodapé da NF-e</span>'
        '  </div>'
        '</div>',
        unsafe_allow_html=True
    )

    if info['alertas']:
        n_tags = len(info['alertas'])
        with st.container(border=True):
            st.markdown(
                '<div style="font-size:0.82rem;font-weight:700;color:#92400E;margin-bottom:6px;">'
                '⚠️ Pontos importantes <span style="font-weight:500;color:#78350F;">(clique para destacar no texto)</span>:'
                '</div>',
                unsafe_allow_html=True
            )
            cols_tags = st.columns(min(max(n_tags, 1), 4))
            for idx_p, ponto_rotulo in enumerate(info['alertas']):
                ativo = (ponto_ativo == ponto_rotulo)
                prefixo = "✔️ " if ativo else "⚠️ "
                with cols_tags[idx_p % len(cols_tags)]:
                    st.button(
                        f"{prefixo}{ponto_rotulo}",
                        key=f"btn_ponto_rodape_{idx_p}_{ponto_rotulo}",
                        use_container_width=True,
                        type="primary" if ativo else "secondary",
                        on_click=_alternar_ponto_rodape,
                        args=(ponto_rotulo,)
                    )

    blocos_texto = []
    if info['infAdFisco']:
        fisco_html = destacar_trechos_rodape_html(info["infAdFisco"], ponto_ativo)
        blocos_texto.append(
            f'<div class="rodape-sec-title">Informações ao Fisco (infAdFisco)</div>'
            f'<div class="rodape-text">{fisco_html}</div>'
        )
    if info['infCpl']:
        cpl_html = destacar_trechos_rodape_html(info["infCpl"], ponto_ativo)
        blocos_texto.append(
            f'<div class="rodape-sec-title">Informações Complementares (infCpl)</div>'
            f'<div class="rodape-text">{cpl_html}</div>'
        )

    st.markdown(
        f'<div class="rodape-panel" style="margin-top:6px;">'
        f'  {"".join(blocos_texto)}'
        f'</div>',
        unsafe_allow_html=True
    )


def renderizar_resumo_divergencias(dados_relatorio, dados_adicionais: dict | None = None):
    """Exibe divergências, alertas laranja e rodapé da NF-e na tela."""
    registros, colunas_ativas, qtd_div, qtd_warn = preparar_dados_resumo_divergencias(dados_relatorio)

    if not registros:
        st.markdown(
            '<div class="div-ok-box">✅ Nenhuma divergência ou alerta encontrado nesta nota (EAN, NCM, ICMS e Base Reduzida 100% conferidos).</div>',
            unsafe_allow_html=True
        )
        renderizar_rodape_nfe(dados_adicionais)
        return

    badges = []
    if qtd_div > 0:
        badges.append(f'<span class="div-count-badge">{qtd_div} divergência(s)</span>')
    if qtd_warn > 0:
        badges.append(f'<span class="div-warn-badge">⚠️ {qtd_warn} alerta(s)</span>')

    ths = ['<th>EAN (NF) + Descrição</th>'] + [f'<th>{col}</th>' for col in colunas_ativas]
    trs = []
    for reg in registros:
        tds = [f'<td class="cell-prod" title="{reg["tooltip"]}">{reg["ean"]}{reg["desc_html"]}</td>']
        for col in colunas_ativas:
            val = reg['celulas_html'].get(col)
            if val is not None:
                tipo_cel = reg['celulas_tipo'].get(col, 'div')
                cls = 'cell-div' if tipo_cel == 'div' else ('cell-warn' if tipo_cel == 'warn' else 'cell-obs')
                tds.append(f'<td class="{cls}">{val}</td>')
            else:
                tds.append('<td class="cell-ok">—</td>')
        trs.append(f'<tr>{"".join(tds)}</tr>')

    st.markdown(
        f'<div class="div-panel">'
        f'  <div class="div-panel-header">'
        f'    <span>🔎 Resumo de Divergências e Alertas (XML × WinThor)</span>'
        f'    <div class="div-badges-wrap">{"".join(badges)}</div>'
        f'  </div>'
        f'  <div class="div-table-wrap">'
        f'    <table class="div-table">'
        f'      <thead><tr>{"".join(ths)}</tr></thead>'
        f'      <tbody>{"".join(trs)}</tbody>'
        f'    </table>'
        f'  </div>'
        f'</div>',
        unsafe_allow_html=True
    )

    renderizar_rodape_nfe(dados_adicionais)


# ==========================================
# 4. HELPERS DE ESTILIZAÇÃO EXCEL (PADRONIZAÇÃO GLOBAL)
# ==========================================
def _obter_origem_coluna(col_name: str) -> str:
    """Classifica a coluna para cor do cabeçalho: 'status', 'winthor', 'sintese' ou 'xml'."""
    if col_name.startswith('Status'):
        return 'status'
    if col_name in ('Diagnóstico e Ação', 'Origem', 'EAN', 'NCM', 'ICMS', 'Base Reduzida', 'Tipo Cálculo', 'Sobra'):
        return 'sintese'
    if (
        'WINTHOR' in col_name
        or 'Whintor' in col_name
        or col_name in ('Tipo_Calculo_ST', 'Cód. Produto')
    ):
        return 'winthor'
    return 'xml'


def _estilizar_cabecalhos_padrao(
    ws,
    super_grupos: list[tuple[str, list[str]]],
    headers: list[str],
    super_start_row: int,
    super_end_row: int,
    sub_row: int,
    status_severidade: dict[str, str] | None = None,
    colunas_esquerda: set[str] | None = None,
):
    """
    Aplica o cabeçalho padrão em qualquer aba do Excel:
    - Super-cabeçalhos de bloco em #1B2631 com divisórias brancas
    - Sub-cabeçalho: XML (#212F3D), Sistema/WinThor (#7F8C8D), Síntese (#1B2631), Status (semântico)
    - Linhas brancas finas entre células e divisória branca média entre blocos
    Retorna o conjunto de índices (1-based) que fecham cada bloco.
    """
    status_severidade = status_severidade or {}
    colunas_esquerda = colunas_esquerda or set()

    fill_bloco = PatternFill(start_color=COLOR_BLOCO_DARK, end_color=COLOR_BLOCO_DARK, fill_type="solid")
    fill_xml = PatternFill(start_color=COLOR_XML_SOLID, end_color=COLOR_XML_SOLID, fill_type="solid")
    fill_win = PatternFill(start_color=COLOR_WIN_SOFT, end_color=COLOR_WIN_SOFT, fill_type="solid")
    fill_st_div = PatternFill(start_color=COLOR_STATUS_DIV, end_color=COLOR_STATUS_DIV, fill_type="solid")
    fill_st_warn = PatternFill(start_color=COLOR_STATUS_WARN, end_color=COLOR_STATUS_WARN, fill_type="solid")
    fill_st_ok = PatternFill(start_color=COLOR_STATUS_OK, end_color=COLOR_STATUS_OK, fill_type="solid")

    font_super = Font(bold=True, size=10, color="FFFFFF")
    font_sub = Font(bold=True, size=10, color="FFFFFF")

    side_branco_fino = Side(style='thin', color='FFFFFF')
    side_branco_medio = Side(style='medium', color='FFFFFF')

    fins_de_bloco = set()
    curr_col = 1
    total_cols = len(headers)

    for label, cols in super_grupos:
        cols_in = [c for c in cols if c in headers]
        if not cols_in:
            continue
        start_c = curr_col
        end_c = curr_col + len(cols_in) - 1
        curr_col += len(cols_in)
        fins_de_bloco.add(end_c)

        if super_start_row != super_end_row or start_c != end_c:
            ws.merge_cells(
                start_row=super_start_row, start_column=start_c,
                end_row=super_end_row, end_column=end_c
            )
        c_sup = ws.cell(row=super_start_row, column=start_c, value=label)
        c_sup.font = font_super
        c_sup.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for r in range(super_start_row, super_end_row + 1):
            for ci in range(start_c, end_c + 1):
                cell_t = ws.cell(row=r, column=ci)
                cell_t.fill = fill_bloco
                cell_t.border = Border(
                    left=side_branco_medio if ci == start_c and start_c > 1 else side_branco_fino,
                    right=side_branco_medio if ci == end_c and end_c < total_cols else side_branco_fino,
                    top=side_branco_fino,
                    bottom=side_branco_fino,
                )

    for r in range(super_start_row, super_end_row + 1):
        ws.row_dimensions[r].height = 20 if super_start_row == super_end_row else 18

    for idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=sub_row, column=idx, value=col_name)
        cell.font = font_sub
        cell.alignment = Alignment(
            horizontal='left' if col_name in colunas_esquerda else 'center',
            vertical='center',
            wrap_text=True
        )

        origem = _obter_origem_coluna(col_name)
        if origem == 'status':
            sev = status_severidade.get(col_name, 'ok')
            if sev == 'div':
                cell.fill = fill_st_div
            elif sev == 'warn':
                cell.fill = fill_st_warn
            else:
                cell.fill = fill_st_ok
        elif origem == 'winthor':
            cell.fill = fill_win
        elif origem == 'sintese':
            cell.fill = fill_bloco
        else:
            cell.fill = fill_xml

        eh_fim = idx in fins_de_bloco and idx < total_cols
        eh_inicio_bloco = (idx - 1) in fins_de_bloco
        cell.border = Border(
            left=side_branco_medio if eh_inicio_bloco else side_branco_fino,
            right=side_branco_medio if eh_fim else side_branco_fino,
            top=side_branco_fino,
            bottom=side_branco_fino,
        )

    ws.row_dimensions[sub_row].height = 25
    return fins_de_bloco


def _adicionar_bloco_rodape_excel(ws, start_row: int, total_cols: int, dados_adicionais: dict | None):
    """Escreve o bloco de Dados Adicionais / Rodapé da NF-e na aba Resumo_Divergencias."""
    info = analisar_rodape_nfe(dados_adicionais)
    if not info['tem_texto']:
        return

    end_col = max(total_cols, 5)
    fill_rodape_hdr = PatternFill(start_color=COLOR_BLOCO_DARK, end_color=COLOR_BLOCO_DARK, fill_type="solid")
    fill_rodape_alerta = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_rodape_sec = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_rodape_txt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    font_hdr = Font(bold=True, size=10, color="FFFFFF")
    font_alerta = Font(bold=True, size=10, color="92400E")
    font_sec = Font(bold=True, size=9, color="334155")
    font_txt = Font(bold=False, size=10, color="1E293B")

    side_cinza = Side(style='thin', color='CBD5E1')
    side_alerta = Side(style='thin', color='F59E0B')
    borda_cinza = Border(left=side_cinza, right=side_cinza, top=side_cinza, bottom=side_cinza)
    borda_alerta = Border(left=side_alerta, right=side_alerta, top=side_alerta, bottom=side_alerta)

    r = start_row

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
    c_hdr = ws.cell(row=r, column=1, value="📌 Dados Adicionais / Rodapé da NF-e")
    c_hdr.font = font_hdr
    c_hdr.alignment = Alignment(horizontal='left', vertical='center')
    for ci in range(1, end_col + 1):
        ws.cell(row=r, column=ci).fill = fill_rodape_hdr
    ws.row_dimensions[r].height = 24
    r += 1

    if info['alertas']:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
        txt_alerta = "⚠️ Pontos importantes: " + "  •  ".join(info['alertas'])
        c_al = ws.cell(row=r, column=1, value=txt_alerta)
        c_al.font = font_alerta
        c_al.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        for ci in range(1, end_col + 1):
            cell_t = ws.cell(row=r, column=ci)
            cell_t.fill = fill_rodape_alerta
            cell_t.border = borda_alerta
        ws.row_dimensions[r].height = 24
        r += 1

    largura_total = sum(
        (ws.column_dimensions[get_column_letter(ci)].width or 18) for ci in range(1, end_col + 1)
    )
    chars_por_linha = max(int(largura_total * 0.92), 80)

    def _escrever_secao(titulo_sec: str, texto_sec: str):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
        c_s = ws.cell(row=r, column=1, value=titulo_sec)
        c_s.font = font_sec
        c_s.alignment = Alignment(horizontal='left', vertical='center')
        for ci in range(1, end_col + 1):
            cell_t = ws.cell(row=r, column=ci)
            cell_t.fill = fill_rodape_sec
            cell_t.border = borda_cinza
        ws.row_dimensions[r].height = 20
        r += 1

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
        c_t = ws.cell(row=r, column=1, value=texto_sec)
        c_t.font = font_txt
        c_t.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        for ci in range(1, end_col + 1):
            cell_t = ws.cell(row=r, column=ci)
            cell_t.fill = fill_rodape_txt
            cell_t.border = borda_cinza
        n_linhas = sum(max(1, math.ceil(len(p) / chars_por_linha)) for p in texto_sec.splitlines())
        ws.row_dimensions[r].height = max(24, min(n_linhas * 16 + 10, 409))
        r += 1

    if info['infAdFisco']:
        _escrever_secao("INFORMAÇÕES AO FISCO (infAdFisco)", info['infAdFisco'])
    if info['infCpl']:
        _escrever_secao("INFORMAÇÕES COMPLEMENTARES (infCpl)", info['infCpl'])


def _montar_aba_resumo_excel(workbook, dados_relatorio, numero_nf: str, dados_adicionais: dict | None = None):
    """Cria a aba 'Resumo_Divergencias' no Excel com cabeçalho padronizado, Diagnóstico e Ação e rodapé."""
    ws = workbook.create_sheet(title='Resumo_Divergencias', index=0)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 80

    registros, colunas_ativas, qtd_div, qtd_warn = preparar_dados_resumo_divergencias(dados_relatorio)
    colunas_excel = [c for c in colunas_ativas if c != 'Observação']

    qtd_tipo_nf = sum(1 for l in dados_relatorio if str(l.get('Tipo_Calculo_WINTHOR', '')).strip().upper() == 'N')
    qtd_tipo_guia = sum(1 for l in dados_relatorio if str(l.get('Tipo_Calculo_WINTHOR', '')).strip().upper() == 'G')

    fill_titulo = PatternFill(start_color=COLOR_BLOCO_DARK, end_color=COLOR_BLOCO_DARK, fill_type="solid")
    fill_zebra_par = PatternFill(start_color="E8EFF7", end_color="E8EFF7", fill_type="solid")
    fill_zebra_impar = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_div_impar = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_div_par = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
    fill_warn_impar = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_warn_par = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
    fill_ok_box = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_kpi_bg = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_kpi_div = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_kpi_guia = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    font_titulo = Font(bold=True, size=11, color="FFFFFF")
    font_prod = Font(bold=True, size=10, color="1B2631")
    font_div = Font(bold=True, size=10, color="991B1B")
    font_warn = Font(bold=True, size=10, color="B45309")
    font_ok_dash = Font(bold=False, size=10, color="94A3B8")
    font_ok_box = Font(bold=True, size=11, color="14532D")
    font_kpi_val = Font(bold=True, size=11, color="1B2631")
    font_kpi_div_val = Font(bold=True, size=11, color="991B1B")
    font_kpi_guia_val = Font(bold=True, size=11, color="B45309")

    side_cinza = Side(style='thin', color='CBD5E1')
    side_bloco = Side(style='medium', color=COLOR_BLOCO_DARK)
    borda_celula = Border(left=side_cinza, right=side_cinza, top=side_cinza, bottom=side_cinza)

    def _escrever_faixa_topo(n_cols: int):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
        c_tit = ws.cell(row=1, column=1)
        c_tit.value = f"🔎 Resumo de Divergências — NF {numero_nf} (XML × WinThor)"
        c_tit.font = font_titulo
        c_tit.alignment = Alignment(horizontal='left', vertical='center')
        for col_i in range(1, n_cols + 1):
            ws.cell(row=1, column=col_i).fill = fill_titulo
        ws.row_dimensions[1].height = 28

        kpis = [
            (f"NF: {numero_nf}", fill_kpi_bg, font_kpi_val),
            (f"Itens Divergentes: {qtd_div}", fill_kpi_div if qtd_div > 0 else fill_kpi_bg, font_kpi_div_val if qtd_div > 0 else font_kpi_val),
            (f"Alertas: {qtd_warn}", fill_kpi_guia if qtd_warn > 0 else fill_kpi_bg, font_kpi_guia_val if qtd_warn > 0 else font_kpi_val),
            (f"Qtd NF: {qtd_tipo_nf}  |  Qtd GUIA: {qtd_tipo_guia}", fill_kpi_guia if qtd_tipo_guia > 0 else fill_kpi_bg, font_kpi_guia_val if qtd_tipo_guia > 0 else font_kpi_val),
        ]
        ws.row_dimensions[2].height = 24
        for idx_k, (txt_k, fl_k, ft_k) in enumerate(kpis, start=1):
            if idx_k <= n_cols:
                c_k = ws.cell(row=2, column=idx_k, value=txt_k)
                c_k.font = ft_k
                c_k.fill = fl_k
                c_k.border = borda_celula
                c_k.alignment = Alignment(horizontal='center', vertical='center')
        for col_i in range(len(kpis) + 1, n_cols + 1):
            c_extra = ws.cell(row=2, column=col_i)
            c_extra.fill = fill_kpi_bg
            c_extra.border = borda_celula
        ws.row_dimensions[3].height = 8

    if not registros:
        _escrever_faixa_topo(5)
        ws.merge_cells('A4:E4')
        c_ok = ws['A4']
        c_ok.value = "✅ Nenhuma divergência ou alerta encontrado nesta nota (EAN, NCM, ICMS e Base Reduzida 100% conferidos)."
        c_ok.font = font_ok_box
        c_ok.fill = fill_ok_box
        c_ok.alignment = Alignment(horizontal='center', vertical='center')
        for col_i in range(1, 6):
            ws.cell(row=4, column=col_i).fill = fill_ok_box
        ws.row_dimensions[4].height = 30
        for col_i in range(1, 6):
            ws.column_dimensions[get_column_letter(col_i)].width = 24

        _adicionar_bloco_rodape_excel(ws, start_row=6, total_cols=5, dados_adicionais=dados_adicionais)
        return

    cabecalhos = ['Cód. Produto', 'EAN (NF) + Descrição'] + colunas_excel + ['Diagnóstico e Ação']
    total_cols = len(cabecalhos)

    _escrever_faixa_topo(total_cols)

    super_grupos = [
        ("IDENTIFICAÇÃO", ['Cód. Produto', 'EAN (NF) + Descrição']),
        ("OCORRÊNCIAS (XML × WINTHOR)", colunas_excel),
        ("AÇÃO", ['Diagnóstico e Ação']),
    ]
    super_row = 4
    hdr_row = 5
    start_data_row = 6

    fins_bloco = _estilizar_cabecalhos_padrao(
        ws=ws,
        super_grupos=super_grupos,
        headers=cabecalhos,
        super_start_row=super_row,
        super_end_row=super_row,
        sub_row=hdr_row,
        colunas_esquerda={'EAN (NF) + Descrição', 'Diagnóstico e Ação'},
    )

    def _borda_dado(col_i: int):
        if col_i in fins_bloco and col_i < total_cols:
            return Border(left=side_cinza, right=side_bloco, top=side_cinza, bottom=side_cinza)
        return borda_celula

    for idx_reg, reg in enumerate(registros, start=start_data_row):
        eh_par = (idx_reg % 2 == 0)
        fill_linha = fill_zebra_par if eh_par else fill_zebra_impar
        fill_div = fill_div_par if eh_par else fill_div_impar
        fill_warn = fill_warn_par if eh_par else fill_warn_impar

        cod_prod_val = reg.get('cod_produto', '-')
        c_cod = ws.cell(row=idx_reg, column=1)
        if cod_prod_val not in ('', '-'):
            try:
                c_cod.value = int(float(cod_prod_val))
                c_cod.number_format = '0'
            except ValueError:
                c_cod.value = cod_prod_val
                c_cod.number_format = '@'
        else:
            c_cod.value = '-'
            c_cod.number_format = '@'
        c_cod.font = font_prod
        c_cod.fill = fill_linha
        c_cod.border = _borda_dado(1)
        c_cod.alignment = Alignment(horizontal='center', vertical='center')

        c_prod = ws.cell(row=idx_reg, column=2, value=reg['texto_ean_desc'])
        c_prod.font = font_prod
        c_prod.fill = fill_linha
        c_prod.border = _borda_dado(2)
        c_prod.alignment = Alignment(horizontal='left', vertical='center')
        c_prod.number_format = '@'

        for offset, col_nome in enumerate(colunas_excel, start=3):
            val = reg['celulas_texto'].get(col_nome)
            cell = ws.cell(row=idx_reg, column=offset)
            cell.border = _borda_dado(offset)
            cell.number_format = '@'

            if val is not None:
                tipo_cel = reg['celulas_tipo'].get(col_nome, 'div')
                cell.value = val
                if tipo_cel == 'warn':
                    cell.font = font_warn
                    cell.fill = fill_warn
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.font = font_div
                    cell.fill = fill_div
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.value = '—'
                cell.font = font_ok_dash
                cell.fill = fill_linha
                cell.alignment = Alignment(horizontal='center', vertical='center')

        c_diag = ws.cell(row=idx_reg, column=total_cols, value=reg.get('diagnostico', ''))
        c_diag.border = _borda_dado(total_cols)
        c_diag.number_format = '@'
        if reg.get('tem_div'):
            c_diag.font = font_div
            c_diag.fill = fill_div
        elif reg.get('tem_warn'):
            c_diag.font = font_warn
            c_diag.fill = fill_warn
        else:
            c_diag.font = font_prod
            c_diag.fill = fill_linha
        c_diag.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    last_data_row = start_data_row + len(registros) - 1
    ws.freeze_panes = f'A{start_data_row}'
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(total_cols)}{last_data_row}"

    larguras_col = {}
    for col_i in range(1, total_cols + 1):
        col_letter = get_column_letter(col_i)
        nome_col = cabecalhos[col_i - 1]
        max_len = len(nome_col)
        for row_i in range(start_data_row, last_data_row + 1):
            val = ws.cell(row=row_i, column=col_i).value
            if val is not None:
                for p in str(val).splitlines():
                    max_len = max(max_len, len(p))
        if nome_col == 'Cód. Produto':
            larg = min(max(max_len + 4, 14), 22)
        elif nome_col == 'Diagnóstico e Ação':
            larg = min(max(max_len + 4, 58), 88)
        else:
            larg = min(max(max_len + 5, 16), 55)
        ws.column_dimensions[col_letter].width = larg
        larguras_col[col_i] = larg

    for row_i in range(start_data_row, last_data_row + 1):
        val_diag = ws.cell(row=row_i, column=total_cols).value
        ws.row_dimensions[row_i].height = _calcular_altura_linha(
            [(str(val_diag or ''), larguras_col[total_cols])],
            altura_min=24
        )

    _adicionar_bloco_rodape_excel(
        ws,
        start_row=last_data_row + 2,
        total_cols=total_cols,
        dados_adicionais=dados_adicionais
    )


def _montar_aba_grupo_excel(
    workbook,
    titulo_aba: str,
    titulo_banner: str,
    super_grupos: list[tuple[str, list[str]]],
    cabecalhos: list[str],
    linhas_dados: list[dict],
    status_severidade: dict[str, str] | None = None,
    colunas_esquerda: set[str] | None = None,
    colunas_ncm: set[str] | None = None,
    colunas_num_int: set[str] | None = None,
    colunas_num_2dec: set[str] | None = None,
    colunas_num_4dec: set[str] | None = None,
):
    """Helper reutilizável para montar abas individuais com cabeçalho padronizado e destaque cirúrgico."""
    if not linhas_dados:
        return

    colunas_esquerda = colunas_esquerda or set()
    colunas_ncm = colunas_ncm or set()
    colunas_num_int = colunas_num_int or set()
    colunas_num_2dec = colunas_num_2dec or set()
    colunas_num_4dec = colunas_num_4dec or set()

    ws = workbook.create_sheet(title=titulo_aba)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 80

    total_cols = len(cabecalhos)
    qtd_div = sum(1 for r in linhas_dados if r['severidade'] == 'div')
    qtd_warn = sum(1 for r in linhas_dados if r['severidade'] == 'warn')

    fill_titulo = PatternFill(start_color=COLOR_BLOCO_DARK, end_color=COLOR_BLOCO_DARK, fill_type="solid")
    fill_zebra_par = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
    fill_zebra_impar = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_cel_div = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    fill_cel_warn = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
    fill_cel_warn_dado = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_cel_ok = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")
    fill_diag_div_impar = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_diag_div_par = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
    fill_diag_warn_impar = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_diag_warn_par = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
    fill_kpi_bg = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_kpi_div = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_kpi_warn = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    font_titulo = Font(bold=True, size=11, color="FFFFFF")
    font_normal = Font(bold=False, size=10, color="000000")
    font_div = Font(bold=True, size=10, color="9C0006")
    font_warn = Font(bold=True, size=10, color="B45309")
    font_kpi_val = Font(bold=True, size=10, color="1B2631")
    font_kpi_div_val = Font(bold=True, size=10, color="991B1B")
    font_kpi_warn_val = Font(bold=True, size=10, color="B45309")

    side_cinza = Side(style='thin', color='CBD5E1')
    side_bloco = Side(style='medium', color=COLOR_BLOCO_DARK)
    borda_celula = Border(left=side_cinza, right=side_cinza, top=side_cinza, bottom=side_cinza)

    # Linha 1: Banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c_tit = ws.cell(row=1, column=1, value=titulo_banner)
    c_tit.font = font_titulo
    c_tit.alignment = Alignment(horizontal='left', vertical='center')
    for ci in range(1, total_cols + 1):
        ws.cell(row=1, column=ci).fill = fill_titulo
    ws.row_dimensions[1].height = 28

    # Linha 2: KPIs rápidos
    kpis = [
        (f"Ocorrências: {len(linhas_dados)}", fill_kpi_bg, font_kpi_val),
        (f"Divergências: {qtd_div}", fill_kpi_div if qtd_div > 0 else fill_kpi_bg, font_kpi_div_val if qtd_div > 0 else font_kpi_val),
        (f"Alertas: {qtd_warn}", fill_kpi_warn if qtd_warn > 0 else fill_kpi_bg, font_kpi_warn_val if qtd_warn > 0 else font_kpi_val),
    ]
    ws.row_dimensions[2].height = 22
    for idx_k, (txt_k, fl_k, ft_k) in enumerate(kpis, start=1):
        if idx_k <= total_cols:
            c_k = ws.cell(row=2, column=idx_k, value=txt_k)
            c_k.font = ft_k
            c_k.fill = fl_k
            c_k.border = borda_celula
            c_k.alignment = Alignment(horizontal='center', vertical='center')
    for ci in range(len(kpis) + 1, total_cols + 1):
        c_ex = ws.cell(row=2, column=ci)
        c_ex.fill = fill_kpi_bg
        c_ex.border = borda_celula
    ws.row_dimensions[3].height = 8

    # Linhas 4 e 5: Cabeçalho padrão (Super-bloco + Colunas XML/Sistema/Status + linhas brancas)
    super_row = 4
    hdr_row = 5
    start_data_row = 6

    fins_bloco = _estilizar_cabecalhos_padrao(
        ws=ws,
        super_grupos=super_grupos,
        headers=cabecalhos,
        super_start_row=super_row,
        super_end_row=super_row,
        sub_row=hdr_row,
        status_severidade=status_severidade,
        colunas_esquerda=colunas_esquerda,
    )

    for idx_r, reg in enumerate(linhas_dados, start=start_data_row):
        eh_par = (idx_r % 2 == 0)
        fill_linha = fill_zebra_par if eh_par else fill_zebra_impar
        cols_div_row = reg.get('cols_div', set())
        cols_warn_row = reg.get('cols_warn', set())

        for ci, val in enumerate(reg['valores'], start=1):
            nome_col = cabecalhos[ci - 1]
            cell = ws.cell(row=idx_r, column=ci)
            cell.border = (
                Border(left=side_cinza, right=side_bloco, top=side_cinza, bottom=side_cinza)
                if (ci in fins_bloco and ci < total_cols)
                else borda_celula
            )

            # Tipagem e formatação numérica idêntica ao Comparativo
            if nome_col in colunas_num_int and str(val).strip() not in ('', '-'):
                try:
                    cell.value = int(float(val))
                    cell.number_format = '0'
                except ValueError:
                    cell.value = str(val)
                    cell.number_format = '@'
            elif nome_col in colunas_ncm and str(val).strip() not in ('', '-'):
                try:
                    cell.value = int(float(val))
                    cell.number_format = '00000000'
                except ValueError:
                    cell.value = str(val)
                    cell.number_format = '@'
            elif nome_col in colunas_num_2dec and isinstance(val, (int, float)):
                cell.value = float(val)
                cell.number_format = '#,##0.00'
            elif nome_col in colunas_num_4dec and isinstance(val, (int, float)):
                cell.value = float(val)
                cell.number_format = '#,##0.0000'
            else:
                cell.value = val
                cell.number_format = '@'

            # Cores e fontes (cirúrgico por célula)
            if nome_col == 'Diagnóstico e Ação':
                if reg['severidade'] == 'div':
                    cell.fill = fill_diag_div_par if eh_par else fill_diag_div_impar
                    cell.font = font_div
                else:
                    cell.fill = fill_diag_warn_par if eh_par else fill_diag_warn_impar
                    cell.font = font_warn
            elif nome_col.startswith('Status'):
                if val == 'DIVERGENTE':
                    cell.fill = fill_cel_div
                    cell.font = font_div
                elif val == 'ALERTA':
                    cell.fill = fill_cel_warn
                    cell.font = font_warn
                elif val == 'OK':
                    cell.fill = fill_cel_ok
                    cell.font = font_normal
                else:
                    cell.fill = fill_linha
                    cell.font = font_normal
            elif nome_col in cols_div_row:
                cell.fill = fill_linha
                cell.font = font_div
            elif nome_col in cols_warn_row:
                cell.fill = fill_cel_warn_dado
                cell.font = font_warn
            else:
                cell.fill = fill_linha
                cell.font = font_normal

            if nome_col == 'Cod_Produto_WINTHOR':
                horiz = 'right'
            elif nome_col in colunas_esquerda:
                horiz = 'left'
            else:
                horiz = 'center'

            cell.alignment = Alignment(
                horizontal=horiz,
                vertical='center',
                wrap_text=(nome_col == 'Diagnóstico e Ação')
            )

    last_data_row = start_data_row + len(linhas_dados) - 1
    ws.freeze_panes = f'A{start_data_row}'
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(total_cols)}{last_data_row}"

    larguras_col = {}
    for ci in range(1, total_cols + 1):
        col_letter = get_column_letter(ci)
        nome_col = cabecalhos[ci - 1]
        if nome_col == 'Orig.':
            larg = 7
        else:
            max_len = len(nome_col)
            for ri in range(start_data_row, last_data_row + 1):
                val = ws.cell(row=ri, column=ci).value
                if val is not None:
                    for p in str(val).splitlines():
                        max_len = max(max_len, len(p))
            if nome_col == 'Diagnóstico e Ação':
                larg = min(max(max_len + 4, 60), 90)
            elif nome_col in ('Descricao_NF', 'Descricao_WINTHOR'):
                larg = min(max(max_len + 4, 24), 48)
            else:
                larg = min(max(max_len + 4, 12), 34)
        ws.column_dimensions[col_letter].width = larg
        larguras_col[ci] = larg

    for ri in range(start_data_row, last_data_row + 1):
        val_diag = ws.cell(row=ri, column=total_cols).value
        ws.row_dimensions[ri].height = _calcular_altura_linha(
            [(str(val_diag or ''), larguras_col[total_cols])],
            altura_min=24
        )


def _montar_abas_individuais_excel(workbook, dados_relatorio, numero_nf: str):
    """Gera abas individuais (Div_EAN, Div_NCM, Div_ICMS_BaseRed, Div_Sobra) apenas se houver ocorrências."""
    linhas_ean = []
    linhas_ncm = []
    linhas_icms = []
    linhas_sobra = []

    for linha in dados_relatorio:
        cod_prod = str(linha.get('Cod_Produto_WINTHOR') or '').strip()
        cod_prod = cod_prod if cod_prod and cod_prod != '-' else '-'
        item_nf = linha.get('Linha_NF', '-')
        desc_nf = str(linha.get('Descricao_NF') or '-').strip()
        desc_win = str(linha.get('Descricao_WINTHOR') or '-').strip()
        ean_nf = str(linha.get('EAN_NF') or '-').strip()
        ean_win = str(linha.get('EAN_WINTHOR') or '-').strip()

        if linha.get('Status_Geral') == 'SOBRA':
            origem_sobra = 'XML' if item_nf != '-' else 'WinThor'
            linhas_sobra.append({
                'severidade': 'div',
                'cols_div': {'Origem'},
                'cols_warn': set(),
                'valores': [
                    origem_sobra,
                    item_nf,
                    linha.get('Linha_WINTHOR', '-'),
                    cod_prod,
                    ean_nf,
                    ean_win,
                    desc_nf,
                    desc_win,
                    _diagnostico_sobra(linha),
                ],
            })
            continue

        # 1. Grupo EAN
        if linha.get('Status_EAN') == 'DIVERGENTE' or linha.get('EAN_Alerta_Vazio'):
            eh_div = linha.get('Status_EAN') == 'DIVERGENTE'
            linhas_ean.append({
                'severidade': 'div' if eh_div else 'warn',
                'cols_div': {'EAN_NF', 'EAN_WINTHOR'} if eh_div else set(),
                'cols_warn': set() if eh_div else {'EAN_NF', 'EAN_WINTHOR'},
                'valores': [
                    item_nf,
                    cod_prod,
                    desc_nf,
                    ean_nf,
                    ean_win,
                    linha.get('Status_EAN', '-'),
                    _diagnostico_ean(linha),
                ],
            })

        # 2. Grupo NCM
        if linha.get('Status_NCM') == 'DIVERGENTE':
            linhas_ncm.append({
                'severidade': 'div',
                'cols_div': {'NCM_NF', 'NCM_WINTHOR'},
                'cols_warn': set(),
                'valores': [
                    item_nf,
                    cod_prod,
                    ean_nf,
                    desc_nf,
                    linha.get('NCM_NF') or '-',
                    linha.get('NCM_WINTHOR') or '-',
                    linha.get('Status_NCM', '-'),
                    _diagnostico_ncm(linha),
                ],
            })

        # 3. Grupo ICMS + Base Reduzida (espelhando colunas e destaque cirúrgico do Comparativo)
        diags_icms = _diagnostico_icms_basered(linha)
        if diags_icms:
            tipo_calc = str(linha.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
            st_icms = linha.get('Status_ICMS_Aliq', '-')
            st_red = linha.get('Status_Red_BC', '-')
            eh_div_icms = (st_icms == 'DIVERGENTE' or st_red == 'DIVERGENTE')

            cols_div = set()
            cols_warn = set()

            if st_icms == 'DIVERGENTE':
                if tipo_calc == 'G':
                    cols_div.update({'ICMS_NF', '%Alíq.Ext. ST Guia_Whintor'})
                    if round(converte_float_br(linha.get('ICMS_WINTHOR', 0.0)), 2) != 0.0:
                        cols_div.add('ICMS_WINTHOR')
                else:
                    cols_div.update({'ICMS_NF', 'ICMS_WINTHOR'})
            elif linha.get('ICMS_Alerta_Guia'):
                cols_warn.add('ICMS_WINTHOR')

            if st_red == 'DIVERGENTE':
                cols_div.update({'ICMS_Efetivo_NF', '% ICMS red. Whintor'})

            if linha.get('Alerta_Origem_Aliq'):
                cols_warn.update({'Orig.', 'ICMS_NF'})

            if linha.get('Alerta_Guia_Coerencia'):
                cols_warn.add('Tipo_Calculo_ST')

            linhas_icms.append({
                'severidade': 'div' if eh_div_icms else 'warn',
                'cols_div': cols_div,
                'cols_warn': cols_warn,
                'valores': [
                    item_nf,
                    cod_prod,
                    ean_nf,
                    desc_nf,
                    linha.get('Orig.', '-'),
                    linha.get('ICMS_NF', 0.0),
                    linha.get('ICMS_WINTHOR', 0.0),
                    linha.get('%Alíq.Ext. ST Guia_Whintor', 0.0),
                    st_icms,
                    linha.get('pRedBC_NF', 0.0),
                    linha.get('ICMS_Efetivo_NF', 0.0),
                    linha.get('% ICMS red. Whintor', 0.0),
                    st_red,
                    linha.get('Tipo_Calculo_ST', '-'),
                    "\n".join(diags_icms),
                ],
            })

    # Severidade de cabeçalho para Status das abas individuais
    st_ean_sev = 'div' if any(r['valores'][5] == 'DIVERGENTE' for r in linhas_ean) else 'warn'
    st_icms_sev = (
        'div' if any(r['valores'][8] == 'DIVERGENTE' for r in linhas_icms)
        else ('warn' if any(r['valores'][8] == 'ALERTA' for r in linhas_icms) else 'ok')
    )
    st_red_sev = 'div' if any(r['valores'][12] == 'DIVERGENTE' for r in linhas_icms) else 'ok'

    _montar_aba_grupo_excel(
        workbook=workbook,
        titulo_aba='Div_EAN',
        titulo_banner=f"🏷️ Divergências e Alertas de EAN — NF {numero_nf}",
        super_grupos=[
            ("IDENTIFICAÇÃO", ['Linha_NF', 'Cod_Produto_WINTHOR', 'Descricao_NF']),
            ("CÓDIGO E EAN", ['EAN_NF', 'EAN_WINTHOR', 'Status_EAN']),
            ("AÇÃO", ['Diagnóstico e Ação']),
        ],
        cabecalhos=['Linha_NF', 'Cod_Produto_WINTHOR', 'Descricao_NF', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN', 'Diagnóstico e Ação'],
        linhas_dados=linhas_ean,
        status_severidade={'Status_EAN': st_ean_sev},
        colunas_esquerda={'Descricao_NF', 'Diagnóstico e Ação'},
        colunas_num_int={'Linha_NF', 'Cod_Produto_WINTHOR'},
    )

    _montar_aba_grupo_excel(
        workbook=workbook,
        titulo_aba='Div_NCM',
        titulo_banner=f"📑 Divergências de NCM — NF {numero_nf}",
        super_grupos=[
            ("IDENTIFICAÇÃO", ['Linha_NF', 'Cod_Produto_WINTHOR', 'EAN_NF', 'Descricao_NF']),
            ("CLASSIFICAÇÃO FISCAL (NCM)", ['NCM_NF', 'NCM_WINTHOR', 'Status_NCM']),
            ("AÇÃO", ['Diagnóstico e Ação']),
        ],
        cabecalhos=['Linha_NF', 'Cod_Produto_WINTHOR', 'EAN_NF', 'Descricao_NF', 'NCM_NF', 'NCM_WINTHOR', 'Status_NCM', 'Diagnóstico e Ação'],
        linhas_dados=linhas_ncm,
        status_severidade={'Status_NCM': 'div'},
        colunas_esquerda={'Descricao_NF', 'Diagnóstico e Ação'},
        colunas_ncm={'NCM_NF', 'NCM_WINTHOR'},
        colunas_num_int={'Linha_NF', 'Cod_Produto_WINTHOR'},
    )

    _montar_aba_grupo_excel(
        workbook=workbook,
        titulo_aba='Div_ICMS_BaseRed',
        titulo_banner=f"💰 Divergências e Alertas de ICMS / Base Reduzida — NF {numero_nf}",
        super_grupos=[
            ("IDENTIFICAÇÃO", ['Linha_NF', 'Cod_Produto_WINTHOR', 'EAN_NF', 'Descricao_NF']),
            ("ALÍQUOTAS DE ICMS", ['Orig.', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq']),
            ("REDUÇÃO DE BASE", ['pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor', 'Status_Red_BC']),
            ("TIPO CÁLCULO ST", ['Tipo_Calculo_ST']),
            ("AÇÃO", ['Diagnóstico e Ação']),
        ],
        cabecalhos=[
            'Linha_NF', 'Cod_Produto_WINTHOR', 'EAN_NF', 'Descricao_NF',
            'Orig.', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq',
            'pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor', 'Status_Red_BC',
            'Tipo_Calculo_ST', 'Diagnóstico e Ação'
        ],
        linhas_dados=linhas_icms,
        status_severidade={'Status_ICMS_Aliq': st_icms_sev, 'Status_Red_BC': st_red_sev},
        colunas_esquerda={'Descricao_NF', 'Diagnóstico e Ação'},
        colunas_num_int={'Linha_NF', 'Cod_Produto_WINTHOR'},
        colunas_num_2dec={'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor'},
        colunas_num_4dec={'pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor'},
    )

    _montar_aba_grupo_excel(
        workbook=workbook,
        titulo_aba='Div_Sobra',
        titulo_banner=f"⚠️ Itens sem Correspondência (Sobra) — NF {numero_nf}",
        super_grupos=[
            ("IDENTIFICAÇÃO", ['Origem', 'Linha_NF', 'Linha_WINTHOR', 'Cod_Produto_WINTHOR']),
            ("PRODUTO", ['EAN_NF', 'EAN_WINTHOR', 'Descricao_NF', 'Descricao_WINTHOR']),
            ("AÇÃO", ['Diagnóstico e Ação']),
        ],
        cabecalhos=[
            'Origem', 'Linha_NF', 'Linha_WINTHOR', 'Cod_Produto_WINTHOR',
            'EAN_NF', 'EAN_WINTHOR', 'Descricao_NF', 'Descricao_WINTHOR',
            'Diagnóstico e Ação'
        ],
        linhas_dados=linhas_sobra,
        colunas_esquerda={'Descricao_NF', 'Descricao_WINTHOR', 'Diagnóstico e Ação'},
        colunas_num_int={'Linha_NF', 'Linha_WINTHOR', 'Cod_Produto_WINTHOR'},
    )


@st.cache_data(show_spinner=False)
def gerar_excel_auditoria_cached(
    itens_xml_tuple, itens_csv_tuple, numero_nf: str, dados_adicionais_tuple,
    inc_ident: bool, inc_ean: bool, inc_desc: bool, inc_qtd: bool,
    inc_ncm: bool, inc_icms: bool, inc_red: bool, inc_st: bool
):
    itens_xml = [dict(x) for x in itens_xml_tuple]
    itens_csv = [dict(c) for c in itens_csv_tuple]
    dados_adicionais = dict(dados_adicionais_tuple) if dados_adicionais_tuple else {}
    data_simples = datetime.now().strftime('%Y%m%d')
    nome_arquivo_saida = f'relatorio_NF_{numero_nf}_{data_simples}.xlsx'

    dados_relatorio = cruzar_itens_auditoria(itens_xml, itens_csv)

    mapa_colunas_blocos = {
        'Identificação': ['Linha_NF', 'Linha_WINTHOR'],
        'Código & EAN': ['Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN'],
        'Descrições': ['Descricao_NF', 'Descricao_WINTHOR'],
        'Quantidades': ['Qtd_NF', 'Qtd_WINTHOR', 'Status_Qtd'],
        'NCM (Fiscal)': ['NCM_NF', 'NCM_WINTHOR', 'Status_NCM'],
        'ICMS': ['Orig.', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq'],
        'Redução de Base': ['pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor', 'Status_Red_BC'],
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
        df_final.to_excel(writer, sheet_name='Comparativo', index=False, startrow=2)

        workbook = writer.book
        worksheet = workbook['Comparativo']

        # Ordem das abas: 1) Resumo_Divergencias | 2) Comparativo | 3+) Abas individuais por grupo
        _montar_aba_resumo_excel(workbook, dados_relatorio, numero_nf, dados_adicionais)
        _montar_abas_individuais_excel(workbook, dados_relatorio, numero_nf)

        worksheet.sheet_view.showGridLines = False
        worksheet.sheet_view.zoomScale = 80

        worksheet.freeze_panes = 'A4'
        worksheet.auto_filter.ref = f"A3:{get_column_letter(len(df_final.columns))}{len(df_final) + 3}"

        fill_zebrado = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
        fill_cinza_sobra = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")

        fill_cel_divergente = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
        fill_cel_alerta = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
        fill_cel_alerta_dado = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        fill_cel_convergente = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")

        font_divergente = Font(color="9C0006", bold=True)
        font_alerta = Font(color="B45309", bold=True)
        font_normal = Font(color="000000", bold=False)

        borda_padrao = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
        borda_fim_bloco = Border(left=Side(style=None), right=Side(style='medium', color=COLOR_BLOCO_DARK), top=Side(style=None), bottom=Side(style=None))

        headers = list(df_final.columns)

        super_grupos_def = []
        if inc_ident: super_grupos_def.append(("IDENTIFICAÇÃO", ['Linha_NF', 'Linha_WINTHOR']))
        if inc_ean: super_grupos_def.append(("CÓDIGO E EAN", ['Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN']))
        if inc_desc: super_grupos_def.append(("DESCRIÇÕES", ['Descricao_NF', 'Descricao_WINTHOR']))
        if inc_qtd: super_grupos_def.append(("QUANTIDADES", ['Qtd_NF', 'Qtd_WINTHOR', 'Status_Qtd']))
        if inc_ncm: super_grupos_def.append(("CLASSIFICAÇÃO FISCAL (NCM)", ['NCM_NF', 'NCM_WINTHOR', 'Status_NCM']))
        if inc_icms: super_grupos_def.append(("ALÍQUOTAS DE ICMS", ['Orig.', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq']))
        if inc_red: super_grupos_def.append(("REDUÇÃO DE BASE", ['pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor', 'Status_Red_BC']))
        if inc_st: super_grupos_def.append(("TIPO CÁLCULO ST", ['Tipo_Calculo_ST']))

        status_columns = ['Status_EAN', 'Status_Qtd', 'Status_NCM', 'Status_ICMS_Aliq', 'Status_Red_BC']
        status_has_warning = {
            'Status_EAN': any(l.get('EAN_Alerta_Vazio') for l in dados_relatorio),
            'Status_ICMS_Aliq': any(l.get('ICMS_Alerta_Guia') for l in dados_relatorio),
        }
        status_severidade = {}
        for sc in status_columns:
            if sc in df_final.columns and (df_final[sc] == 'DIVERGENTE').any():
                status_severidade[sc] = 'div'
            elif status_has_warning.get(sc):
                status_severidade[sc] = 'warn'
            else:
                status_severidade[sc] = 'ok'

        fins_bloco = _estilizar_cabecalhos_padrao(
            ws=worksheet,
            super_grupos=super_grupos_def,
            headers=headers,
            super_start_row=1,
            super_end_row=2,
            sub_row=3,
            status_severidade=status_severidade,
        )

        align_center_cols = [
            'Linha_NF', 'Linha_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Qtd_NF', 'Qtd_WINTHOR',
            'NCM_NF', 'NCM_WINTHOR', 'Orig.', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor',
            'pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor', 'Tipo_Calculo_ST'
        ]

        status_dependents = {
            'Status_EAN': ['EAN_NF', 'EAN_WINTHOR'],
            'Status_Qtd': ['Qtd_NF', 'Qtd_WINTHOR'],
            'Status_NCM': ['NCM_NF', 'NCM_WINTHOR'],
            'Status_ICMS_Aliq': ['ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor'],
            'Status_Red_BC': ['pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor']
        }

        for row in range(4, len(df_final) + 4):
            linha_dict = dados_relatorio[row - 4]
            for col_idx in range(1, len(headers) + 1):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.border = borda_fim_bloco if col_idx in fins_bloco else borda_padrao

                col_name = headers[col_idx - 1]
                current_align = cell.alignment or Alignment()
                cell.alignment = Alignment(
                    horizontal=current_align.horizontal or 'left',
                    vertical='center',
                    wrap_text=current_align.wrap_text
                )

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
                    if 'pRedBC' in col_name or 'ICMS_Efetivo' in col_name or 'red. Whintor' in col_name:
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

            if 'Tipo_Calculo_ST' in headers:
                cell_tipo_st = worksheet.cell(row=row, column=headers.index('Tipo_Calculo_ST') + 1)
                cell_tipo_st.font = font_alerta if linha_dict.get('Alerta_Guia_Coerencia') else font_normal

            tipo_calc_linha = str(linha_dict.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
            for col_status, colunas_dados in status_dependents.items():
                if col_status in headers:
                    idx_status = headers.index(col_status) + 1
                    cell_status = worksheet.cell(row=row, column=idx_status)
                    val_status = cell_status.value

                    eh_alerta_laranja = (
                        val_status == 'ALERTA'
                        or (col_status == 'Status_EAN' and linha_dict.get('EAN_Alerta_Vazio'))
                        or (col_status == 'Status_ICMS_Aliq' and linha_dict.get('ICMS_Alerta_Guia'))
                    )

                    for col_afetada in colunas_dados:
                        if col_afetada in headers:
                            worksheet.cell(row=row, column=headers.index(col_afetada) + 1).font = font_normal

                    if val_status == 'DIVERGENTE':
                        cell_status.fill = fill_cel_divergente
                        cell_status.font = font_divergente

                        if col_status == 'Status_ICMS_Aliq':
                            if tipo_calc_linha == 'G':
                                cols_erro = ['ICMS_NF', '%Alíq.Ext. ST Guia_Whintor']
                                if round(converte_float_br(linha_dict.get('ICMS_WINTHOR', 0.0)), 2) != 0.0:
                                    cols_erro.append('ICMS_WINTHOR')
                            else:
                                cols_erro = ['ICMS_NF', 'ICMS_WINTHOR']
                        elif col_status == 'Status_Red_BC':
                            cols_erro = ['ICMS_Efetivo_NF', '% ICMS red. Whintor']
                        else:
                            cols_erro = colunas_dados

                        for col_afetada in cols_erro:
                            if col_afetada in headers:
                                cell_dado = worksheet.cell(row=row, column=headers.index(col_afetada) + 1)
                                cell_dado.font = font_divergente
                    elif eh_alerta_laranja:
                        cell_status.fill = fill_cel_alerta
                        cell_status.font = font_alerta

                        if col_status == 'Status_ICMS_Aliq':
                            cols_alerta = ['ICMS_WINTHOR']
                        else:
                            cols_alerta = colunas_dados

                        for col_afetada in cols_alerta:
                            if col_afetada in headers:
                                cell_dado = worksheet.cell(row=row, column=headers.index(col_afetada) + 1)
                                cell_dado.font = font_alerta
                                cell_dado.fill = fill_cel_alerta_dado
                    elif val_status == 'OK':
                        cell_status.fill = fill_cel_convergente

        letras_nao_status = [
            get_column_letter(idx)
            for idx, col_name in enumerate(headers, 1)
            if not col_name.startswith('Status')
        ]
        for l_col in letras_nao_status:
            rule = FormulaRule(formula=['MOD(ROW(),2)=0'], stopIfTrue=True, fill=fill_zebrado)
            worksheet.conditional_formatting.add(f'{l_col}4:{l_col}{len(df_final) + 3}', rule)

        for idx, col in enumerate(worksheet.columns, 1):
            col_name = headers[idx - 1]
            column_letter = get_column_letter(col[0].column)
            if col_name == 'Orig.':
                worksheet.column_dimensions[column_letter].width = 7
                continue
            max_length = 0
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


def _freeze_dict(d: dict) -> tuple[tuple[str, object], ...]:
    return tuple(sorted(d.items()))


# ==========================================
# 5. CALLBACKS DE GERENCIAMENTO DO LOTE
# ==========================================
def _remover_nf_ativa():
    chave = st.session_state.get("sb_nf_ativa")
    if chave and chave in st.session_state["xml_lote"]:
        st.session_state["xml_lote"].pop(chave, None)
    restantes = list(st.session_state["xml_lote"].keys())
    if restantes:
        st.session_state["sb_nf_ativa"] = restantes[0]
    else:
        st.session_state.pop("sb_nf_ativa", None)


def _limpar_lote():
    st.session_state["xml_lote"] = {}
    st.session_state.pop("sb_nf_ativa", None)
    st.session_state["xml_uploader_key"] += 1


# ==========================================
# 6. INTERFACE STREAMLIT (UI/UX OTIMIZADA)
# ==========================================
st.markdown("""
<div class="hero-header">
    <h1>📊 Auditoria Fiscal de Pré-Entrada (XML × WinThor)</h1>
    <p>Cruzamento automatizado de EAN, NCM, ICMS, Redução de Base e Tipo de Cálculo ST com diagnóstico orientado à ação.</p>
</div>
""", unsafe_allow_html=True)

if "xml_lote" not in st.session_state:
    st.session_state["xml_lote"] = {}
if "xml_uploader_key" not in st.session_state:
    st.session_state["xml_uploader_key"] = 0
if "ultimo_csv_assinatura" not in st.session_state:
    st.session_state["ultimo_csv_assinatura"] = None
if "ultimo_lote_assinatura" not in st.session_state:
    st.session_state["ultimo_lote_assinatura"] = None
if "ponto_rodape_ativo" not in st.session_state:
    st.session_state["ponto_rodape_ativo"] = None

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
        for arq_x in arquivos_xml:
            sig_x = f"{arq_x.name}:{arq_x.size}"
            if sig_x not in st.session_state["xml_lote"]:
                bruto_x = arq_x.getvalue()
                num_nf_x = extrair_numero_nf_rapido(bruto_x)
                st.session_state["xml_lote"][sig_x] = {
                    "numero_nf": num_nf_x,
                    "nf_norm": num_nf_x.lstrip('0'),
                    "nome_arquivo": arq_x.name,
                    "bytes": bruto_x,
                }
        st.session_state["xml_uploader_key"] += 1
        st.rerun()

with col_up2:
    st.markdown(
        '<p class="section-label"><span class="step-num">2</span> Relatório de Pré-Entrada (WinThor)</p>',
        unsafe_allow_html=True
    )
    arquivo_dados = st.file_uploader("Upload WinThor", type=["csv", "xlsx"], label_visibility="collapsed")

lote_xml = st.session_state["xml_lote"]
chaves_ordenadas = list(lote_xml.keys())
assinatura_lote = tuple(chaves_ordenadas)

if arquivo_dados is not None:
    assinatura_csv = f"{arquivo_dados.name}:{arquivo_dados.size}"
    csv_mudou = assinatura_csv != st.session_state["ultimo_csv_assinatura"]
    lote_mudou = assinatura_lote != st.session_state["ultimo_lote_assinatura"]

    if csv_mudou or lote_mudou:
        st.session_state["ultimo_csv_assinatura"] = assinatura_csv
        st.session_state["ultimo_lote_assinatura"] = assinatura_lote
        cand_csv = _numeros_do_nome(arquivo_dados.name)
        match_chave = next(
            (k for k in chaves_ordenadas if lote_xml[k]["nf_norm"] and lote_xml[k]["nf_norm"] in cand_csv),
            None
        )
        if match_chave:
            st.session_state["sb_nf_ativa"] = match_chave
else:
    st.session_state["ultimo_lote_assinatura"] = assinatura_lote

if chaves_ordenadas:
    if st.session_state.get("sb_nf_ativa") not in lote_xml:
        st.session_state["sb_nf_ativa"] = chaves_ordenadas[0]
else:
    st.session_state.pop("sb_nf_ativa", None)

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


with st.container(border=True):
    top_col1, top_col2 = st.columns([4, 1], vertical_alignment="center")
    with top_col1:
        st.markdown('<p class="section-label" style="margin-bottom:0;">⚙️ Blocos para Exportação</p>', unsafe_allow_html=True)
    with top_col2:
        st.checkbox("Selecionar Todos", key="flag_sel_todos", on_change=_ao_clicar_selecionar_todos)

    st.markdown("<hr style='margin: 6px 0 10px 0; border-color: #EEF2F6;'>", unsafe_allow_html=True)

    _, b_col1, b_col2, b_col3, b_col4, _ = st.columns([0.4, 1.1, 1.1, 1.1, 1.1, 0.4], gap="medium")
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

chave_ativa = st.session_state.get("sb_nf_ativa")
if lote_xml and arquivo_dados and chave_ativa in lote_xml:
    xml_selecionado = lote_xml[chave_ativa]
    try:
        numero_nf, itens_xml, dados_adicionais = ler_xml_nfe_bytes(xml_selecionado["bytes"])
    except ET.ParseError:
        st.markdown(
            f'<div class="v-badge err">🚫 O arquivo <b>{xml_selecionado["nome_arquivo"]}</b> não é um XML de NF-e válido.</div>',
            unsafe_allow_html=True
        )
        itens_xml = None
        dados_adicionais = None

    if itens_xml is not None:
        df_csv, avisos_prep, colunas_faltantes = ler_csv_winthor_cached(arquivo_dados.getvalue(), arquivo_dados.name)
        itens_csv = montar_itens_csv(df_csv) if not colunas_faltantes else []
        val = validar_upload(itens_xml, itens_csv, numero_nf, arquivo_dados.name, colunas_faltantes)

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
                _freeze_dict(dados_adicionais),
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

            dados_cruzados = cruzar_itens_auditoria(itens_xml, itens_csv)
            renderizar_resumo_divergencias(dados_cruzados, dados_adicionais)

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
