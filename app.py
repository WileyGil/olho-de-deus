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
    :root {
        color-scheme: light !important;
        --primary-color: #1B2631 !important;
        --background-color: #F4F6F9 !important;
        --secondary-background-color: #FFFFFF !important;
        --text-color: #1B2631 !important;
    }
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"], .stApp {
        background-color: #F4F6F9 !important;
        color: #1B2631 !important;
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

    /* Uploaders claros e compactos (remove fundo preto do cartão do CSV) */
    [data-testid="stFileUploader"],
    [data-testid="stFileUploader"] > div,
    [data-testid="stFileUploader"] section,
    [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
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
        color: #1E293B !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        background-color: #EFF6FF !important;
        color: #1D4ED8 !important;
        border: 1px solid #BFDBFE !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }

    /* Lista de arquivos e cartão de arquivo enviado sempre claros */
    [data-testid="stFileUploaderDropzone"] > div:last-child,
    [data-testid="stFileUploaderDropzone"] ul,
    [data-testid="stFileUploader"] ul {
        width: 100% !important;
        max-width: 100% !important;
        max-height: 145px !important;
        overflow-y: auto !important;
        flex: 1 1 100% !important;
        background-color: transparent !important;
    }
    [data-testid="stFileUploader"] li,
    [data-testid="stFileUploaderFile"],
    [data-testid="stUploadedFile"],
    [data-testid="stFileUploaderFileData"],
    [data-testid="stFileUploader"] [role="listitem"] {
        background-color: #EEF2F6 !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        padding: 6px 10px !important;
        margin-bottom: 4px !important;
    }
    [data-testid="stFileUploader"] li *,
    [data-testid="stFileUploaderFile"] *,
    [data-testid="stUploadedFile"] *,
    [data-testid="stFileUploaderFileData"] * {
        background-color: transparent !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
    }
    [data-testid="stFileUploaderFileName"],
    [data-testid="stFileUploaderFile"] span,
    [data-testid="stUploadedFile"] span {
        font-weight: 700 !important;
        font-size: 0.82rem !important;
        line-height: 1.3 !important;
    }
    [data-testid="stFileUploaderFile"] small,
    [data-testid="stUploadedFile"] small,
    [data-testid="stFileUploader"] li small {
        color: #475569 !important;
        -webkit-text-fill-color: #475569 !important;
        font-size: 0.74rem !important;
    }
    [data-testid="stFileUploaderFile"] svg,
    [data-testid="stUploadedFile"] svg,
    [data-testid="stFileUploaderDeleteBtn"] button,
    [data-testid="stFileUploaderFile"] button,
    [data-testid="stUploadedFile"] button {
        color: #1B2631 !important;
        fill: #1B2631 !important;
        background-color: transparent !important;
        border: none !important;
        flex-shrink: 0 !important;
    }

    /* Containers e listagem do lote em fundo claro */
    [data-testid="stVerticalBlockBorderWrapper"],
    [data-testid="stVerticalBlockBorderWrapper"] > div,
    [data-testid="stVerticalBlockBorderWrapper"] > div > div {
        background-color: #FFFFFF !important;
        border-color: #E2E8F0 !important;
        border-radius: 8px !important;
        color: #0F172A !important;
    }

    /* Selectbox da NF ativa: fundo claro e fonte escura tanto fechado quanto aberto */
    [data-testid="stSelectbox"],
    [data-testid="stSelectbox"] > div,
    [data-testid="stSelectbox"] > div > div,
    [data-testid="stSelectbox"] [data-baseweb="select"],
    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stSelectbox"] [data-baseweb="select"] > div > div,
    [data-baseweb="popover"],
    [data-baseweb="popover"] > div,
    [data-baseweb="popover"] ul,
    [data-baseweb="popover"] li,
    [role="listbox"],
    [role="option"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border-color: #CBD5E1 !important;
    }
    [data-testid="stSelectbox"] span,
    [data-testid="stSelectbox"] div,
    [data-testid="stSelectbox"] [data-baseweb="select"] *,
    [role="option"] span,
    [role="option"] div,
    [role="option"] * {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        opacity: 1 !important;
        font-weight: 700 !important;
    }
    [role="option"]:hover,
    [role="option"][aria-selected="true"] {
        background-color: #E8EFF7 !important;
        color: #0F172A !important;
    }
    [data-testid="stSelectbox"] svg {
        color: #0F172A !important;
        fill: #0F172A !important;
    }

    /* Botões padrão (Remover / Limpar tudo / tags) sempre claros */
    div[data-testid="stButton"] > button,
    div[data-testid="stButton"] > button[kind="secondary"],
    button[kind="secondary"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }
    div[data-testid="stButton"] > button:hover,
    div[data-testid="stButton"] > button[kind="secondary"]:hover,
    button[kind="secondary"]:hover {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
        border-color: #94A3B8 !important;
    }
    div[data-testid="stButton"] > button span,
    div[data-testid="stButton"] > button div,
    div[data-testid="stButton"] > button p {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
    }

    /* Flags (checkboxes do relatório): caixinha desmarcada sempre branca com borda cinza */
    div[data-testid="stCheckbox"] {
        margin-bottom: 4px;
    }
    div[data-testid="stCheckbox"] label,
    div[data-testid="stCheckbox"] label p,
    div[data-testid="stCheckbox"] label div {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        opacity: 1 !important;
    }
    label[data-baseweb="checkbox"] > span:first-child,
    div[data-testid="stCheckbox"] label[data-baseweb="checkbox"] > span:first-of-type {
        background-color: #FFFFFF !important;
        border: 1.5px solid #94A3B8 !important;
        box-shadow: none !important;
    }
    label[data-baseweb="checkbox"]:has(input:checked) > span:first-child,
    div[data-testid="stCheckbox"] label[data-baseweb="checkbox"]:has(input:checked) > span:first-of-type {
        background-color: #DC2626 !important;
        border-color: #DC2626 !important;
    }
    label[data-baseweb="checkbox"]:has(input:checked) > span:first-child svg {
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
    }

    /* Barra do lote */
    .lote-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background-color: #EEF2F6;
        color: #0F172A;
        border: 1px solid #CBD5E1;
        padding: 5px 12px;
        border-radius: 20px;
        font-size: 0.82rem;
        font-weight: 700;
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
        color: #475569;
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
        -webkit-text-fill-color: #FFFFFF !important;
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

    /* Resumo de Divergências em 1 linha, rolagem horizontal e EAN congelado */
    .div-panel {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 10px;
        padding: 14px 16px 16px 16px;
        margin-top: 20px;
        box-shadow: 0 3px 10px rgba(15, 23, 42, 0.06);
    }
    .div-panel-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 12px;
        font-size: 0.90rem;
        font-weight: 700;
        color: #1B2631;
    }
    .div-badges-wrap {
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .div-count-badge {
        background-color: #DC2626;
        color: #FFFFFF;
        font-size: 0.76rem;
        font-weight: 700;
        padding: 3px 11px;
        border-radius: 14px;
    }
    .div-warn-badge {
        background-color: #D97706;
        color: #FFFFFF;
        font-size: 0.76rem;
        font-weight: 700;
        padding: 3px 11px;
        border-radius: 14px;
    }
    .div-table-wrap {
        width: 100%;
        overflow-x: auto;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
        position: relative;
        scrollbar-width: thin;
        scrollbar-color: #94A3B8 #F1F5F9;
    }
    .div-table-wrap::-webkit-scrollbar {
        height: 10px;
    }
    .div-table-wrap::-webkit-scrollbar-track {
        background: #F1F5F9;
        border-radius: 0 0 8px 8px;
    }
    .div-table-wrap::-webkit-scrollbar-thumb {
        background: #94A3B8;
        border-radius: 6px;
    }
    .div-table-wrap::-webkit-scrollbar-thumb:hover {
        background: #64748B;
    }
    .div-table {
        width: max-content;
        min-width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 0.84rem;
        background-color: #FFFFFF;
    }
    .div-table thead th {
        background-color: #1B2631;
        color: #FFFFFF;
        font-weight: 700;
        text-align: center;
        padding: 10px 14px;
        border-right: 1px solid #334155;
        letter-spacing: 0.03em;
        text-transform: uppercase;
        font-size: 0.76rem;
        white-space: nowrap;
    }
    .div-table thead th:first-child {
        text-align: left;
        position: sticky;
        left: 0;
        z-index: 3;
        background-color: #1B2631;
        box-shadow: 2px 0 6px rgba(15, 23, 42, 0.16);
    }
    .div-table thead th:last-child {
        border-right: none;
    }
    .div-table tbody tr {
        background-color: #FFFFFF;
        transition: background-color 0.15s ease;
    }
    .div-table tbody tr:nth-child(even) {
        background-color: #E8EFF7;
    }
    .div-table tbody tr:hover {
        background-color: #DBEAFE;
    }
    .div-table tbody td {
        padding: 9px 14px;
        border-top: 1px solid #CBD5E1;
        border-right: 1px solid #CBD5E1;
        text-align: center;
        white-space: nowrap !important;
    }
    .div-table tbody td:last-child {
        border-right: none;
    }
    .div-table tbody td.cell-prod {
        text-align: left;
        font-weight: 700;
        color: #1B2631;
        position: sticky;
        left: 0;
        z-index: 2;
        background-color: #FFFFFF;
        box-shadow: 2px 0 6px rgba(15, 23, 42, 0.08);
    }
    .div-table tbody tr:nth-child(even) td.cell-prod {
        background-color: #E8EFF7;
    }
    .div-table tbody tr:hover td.cell-prod {
        background-color: #DBEAFE;
    }
    .div-table tbody td.cell-prod .cell-desc {
        font-weight: 500;
        color: #475569;
        font-size: 0.78rem;
        margin-left: 6px;
    }
    .div-table tbody td.cell-div {
        background-color: #FEE2E2;
        color: #991B1B;
        font-weight: 700;
    }
    .div-table tbody tr:nth-child(even) td.cell-div {
        background-color: #FECACA;
    }
    .div-table tbody tr:hover td.cell-div {
        background-color: #FCA5A5;
        color: #7F1D1D;
    }
    .div-table tbody td.cell-warn {
        background-color: #FEF3C7;
        color: #B45309;
        font-weight: 700;
    }
    .div-table tbody tr:nth-child(even) td.cell-warn {
        background-color: #FDE68A;
    }
    .div-table tbody tr:hover td.cell-warn {
        background-color: #FCD34D;
        color: #78350F;
    }
    .div-table tbody td.cell-obs {
        text-align: left;
        white-space: nowrap !important;
        font-size: 0.79rem;
        color: #92400E;
        background-color: #FFFBEB;
        font-weight: 600;
    }
    .div-table tbody tr:nth-child(even) td.cell-obs {
        background-color: #FEF3C7;
    }
    .div-table tbody td.cell-ok {
        color: #94A3B8;
        font-weight: 500;
    }
    .div-sep {
        color: #DC2626;
        font-weight: 800;
        margin: 0 3px;
    }
    .warn-sep {
        color: #D97706;
        font-weight: 800;
        margin: 0 3px;
    }
    .red-detail {
        display: inline;
        font-size: 0.76rem;
        font-weight: 500;
        color: #7F1D1D;
        margin-left: 6px;
    }
    .div-ok-box {
        text-align: center;
        background-color: #DCFCE7;
        border: 1px solid #86EFAC;
        color: #14532D;
        border-radius: 8px;
        padding: 11px 14px;
        margin-top: 20px;
        font-size: 0.87rem;
        font-weight: 700;
    }

    /* Bloco de Dados Adicionais / Rodapé da NF-e */
    .rodape-panel {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-left: 4px solid #1B2631;
        border-radius: 10px;
        padding: 14px 16px;
        margin-top: 16px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
    }
    .rodape-header {
        font-size: 0.90rem;
        font-weight: 700;
        color: #1B2631;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 8px;
    }
    .rodape-highlight {
        background-color: #FDE047;
        color: #78350F;
        font-weight: 700;
        padding: 1px 4px;
        border-radius: 3px;
        border-bottom: 2px solid #D97706;
    }
    .rodape-sec-title {
        font-size: 0.76rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #475569;
        margin: 6px 0 4px 0;
    }
    .rodape-text {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        padding: 10px 12px;
        font-size: 0.82rem;
        line-height: 1.5;
        color: #1E293B;
        white-space: pre-wrap;
        word-break: break-word;
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
    ("ICMS-ST / ST", re.compile(r'\b(?:icms[\s-]*st|st|b\.?calc\.?st)\b|substituicao\s+tributaria', re.IGNORECASE)),
    ("GNRE", re.compile(r'\bgnre\b|guia\s+nacional', re.IGNORECASE)),
    ("Redução de Base", re.compile(r'reducao\s+(?:de\s+|da\s+)?base|\bred\.?\s*bc\b|base\s+de\s+calculo\s+reduzida', re.IGNORECASE)),
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
        '%Alíq.Ext.ST Guia',
    ]
    faltantes = [c for c in obrigatorias if c not in col_set]
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

    inf_adic = root.find('.//ns:infAdic', NS)
    dados_adicionais = {
        'infAdFisco': txt(inf_adic, 'infAdFisco'),
        'infCpl': txt(inf_adic, 'infCpl'),
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

        icms, origem = 0.0, 'NAC'
        icms_node = det.find('.//ns:ICMS', NS)
        if icms_node is not None and len(icms_node) > 0:
            sub = icms_node[0]
            p = sub.find('ns:pICMS', NS)
            if p is not None and p.text:
                icms = converte_float_br(p.text)
            o = sub.find('ns:orig', NS)
            if o is not None and o.text:
                origem = obter_sigla_origem(o.text.strip())

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
# 2. VALIDAÇÃO, CRUZAMENTO E RELATÓRIOS
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
                linha['Status_EAN'] = 'OK'
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
                    linha['Status_ICMS_Aliq'] = 'OK'
                    linha['ICMS_Alerta_Guia'] = True
                    linha['ICMS_Obs'] = (
                        f"Item Guia (G) com %ICMS preenchido ({_fmt_num_curto(icms_csv)}%) "
                        f"no WinThor — deveria estar zerado."
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

        else:
            linha['Status_Geral'] = 'SOBRA'
            linha['ICMS_Efetivo_NF'] = '-'
            for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                linha[st_col] = '-'

        dados_relatorio.append(linha)

    return dados_relatorio


def preparar_dados_resumo_divergencias(dados_relatorio):
    """Extrai registros de divergências e alertas laranja (fonte única para tela e Excel)."""
    registros = []
    colunas_encontradas = {
        'EAN': False,
        'NCM': False,
        'ICMS': False,
        'Base Reduzida': False,
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
                celulas_texto['ICMS'] = f'Guia {i_x}% (ICMS {i_w}%)'
                celulas_html['ICMS'] = f'Guia {i_x}% <span class="warn-sep">•</span> ICMS {i_w}%'
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

            ean_nf = str(linha.get('EAN_NF') or '').strip()
            if ean_nf and ean_nf not in ('-', 'SEM EAN'):
                rotulo_ean = ean_nf
            else:
                rotulo_ean = f"Sem EAN (Item {linha.get('Linha_NF', '-')})"

            desc_completa = str(linha.get('Descricao_NF') or linha.get('Descricao_WINTHOR') or '').strip()
            desc_resumida = (desc_completa[:38].rstrip() + '…') if len(desc_completa) > 38 else desc_completa
            texto_ean_desc = f"{rotulo_ean} — {desc_resumida}" if desc_resumida else rotulo_ean
            desc_html = f'<span class="cell-desc">— {html.escape(desc_resumida)}</span>' if desc_resumida else ''

            registros.append({
                'ean': html.escape(rotulo_ean),
                'desc_html': desc_html,
                'texto_ean_desc': texto_ean_desc,
                'tooltip': html.escape(desc_completa),
                'celulas_texto': celulas_texto,
                'celulas_html': celulas_html,
                'celulas_tipo': celulas_tipo,
            })

    ordem_colunas = ['EAN', 'NCM', 'ICMS', 'Base Reduzida', 'Sobra', 'Observação']
    colunas_ativas = [c for c in ordem_colunas if colunas_encontradas[c]]
    return registros, colunas_ativas, qtd_divergencias, qtd_alertas


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


def _adicionar_bloco_rodape_excel(ws, start_row: int, total_cols: int, dados_adicionais: dict | None):
    """Escreve o bloco de Dados Adicionais / Rodapé da NF-e na aba Resumo_Divergencias."""
    info = analisar_rodape_nfe(dados_adicionais)
    if not info['tem_texto']:
        return

    end_col = max(total_cols, 5)
    fill_rodape_hdr = PatternFill(start_color="1B2631", end_color="1B2631", fill_type="solid")
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
    """Cria a aba 'Resumo_Divergencias' no Excel idêntica ao resumo exibido no front."""
    ws = workbook.create_sheet(title='Resumo_Divergencias', index=0)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 80

    registros, colunas_ativas, qtd_div, qtd_warn = preparar_dados_resumo_divergencias(dados_relatorio)

    fill_titulo = PatternFill(start_color="1B2631", end_color="1B2631", fill_type="solid")
    fill_header = PatternFill(start_color="212F3D", end_color="212F3D", fill_type="solid")
    fill_zebra_par = PatternFill(start_color="E8EFF7", end_color="E8EFF7", fill_type="solid")
    fill_zebra_impar = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_div_impar = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_div_par = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
    fill_warn_impar = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_warn_par = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
    fill_obs_impar = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    fill_obs_par = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_ok_box = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

    font_titulo = Font(bold=True, size=11, color="FFFFFF")
    font_header = Font(bold=True, size=10, color="FFFFFF")
    font_prod = Font(bold=True, size=10, color="1B2631")
    font_div = Font(bold=True, size=10, color="991B1B")
    font_warn = Font(bold=True, size=10, color="B45309")
    font_obs = Font(bold=True, size=9, color="92400E")
    font_ok_dash = Font(bold=False, size=10, color="94A3B8")
    font_ok_box = Font(bold=True, size=11, color="14532D")

    side_cinza = Side(style='thin', color='CBD5E1')
    borda_celula = Border(left=side_cinza, right=side_cinza, top=side_cinza, bottom=side_cinza)

    if not registros:
        ws.merge_cells('A1:E1')
        c_tit = ws['A1']
        c_tit.value = f"🔎 Resumo de Divergências — NF {numero_nf} (XML × WinThor)"
        c_tit.font = font_titulo
        c_tit.fill = fill_titulo
        c_tit.alignment = Alignment(horizontal='center', vertical='center')
        for col_i in range(1, 6):
            ws.cell(row=1, column=col_i).fill = fill_titulo
        ws.row_dimensions[1].height = 28

        ws.merge_cells('A3:E3')
        c_ok = ws['A3']
        c_ok.value = "✅ Nenhuma divergência ou alerta encontrado nesta nota (EAN, NCM, ICMS e Base Reduzida 100% conferidos)."
        c_ok.font = font_ok_box
        c_ok.fill = fill_ok_box
        c_ok.alignment = Alignment(horizontal='center', vertical='center')
        for col_i in range(1, 6):
            ws.cell(row=3, column=col_i).fill = fill_ok_box
        ws.row_dimensions[3].height = 30
        for col_i in range(1, 6):
            ws.column_dimensions[get_column_letter(col_i)].width = 24

        _adicionar_bloco_rodape_excel(ws, start_row=5, total_cols=5, dados_adicionais=dados_adicionais)
        return

    cabecalhos = ['EAN (NF) + Descrição'] + colunas_ativas
    total_cols = len(cabecalhos)

    partes_resumo = []
    if qtd_div > 0:
        partes_resumo.append(f"{qtd_div} divergência(s)")
    if qtd_warn > 0:
        partes_resumo.append(f"{qtd_warn} alerta(s)")
    resumo_contagem = "  •  ".join(partes_resumo)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c_tit = ws.cell(row=1, column=1)
    c_tit.value = f"🔎 Resumo de Divergências e Alertas — NF {numero_nf} (XML × WinThor)  |  {resumo_contagem}"
    c_tit.font = font_titulo
    c_tit.alignment = Alignment(horizontal='left', vertical='center')
    for col_i in range(1, total_cols + 1):
        ws.cell(row=1, column=col_i).fill = fill_titulo
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 8

    for col_i, nome_col in enumerate(cabecalhos, start=1):
        cell = ws.cell(row=3, column=col_i, value=nome_col.upper())
        cell.font = font_header
        cell.fill = fill_header
        cell.border = borda_celula
        cell.alignment = Alignment(horizontal='left' if col_i in (1,) or nome_col == 'Observação' else 'center', vertical='center')
    ws.row_dimensions[3].height = 24

    for idx_reg, reg in enumerate(registros, start=4):
        eh_par = (idx_reg % 2 == 0)
        fill_linha = fill_zebra_par if eh_par else fill_zebra_impar
        fill_div = fill_div_par if eh_par else fill_div_impar
        fill_warn = fill_warn_par if eh_par else fill_warn_impar
        fill_obs = fill_obs_par if eh_par else fill_obs_impar

        c_prod = ws.cell(row=idx_reg, column=1, value=reg['texto_ean_desc'])
        c_prod.font = font_prod
        c_prod.fill = fill_linha
        c_prod.border = borda_celula
        c_prod.alignment = Alignment(horizontal='left', vertical='center')
        c_prod.number_format = '@'

        for offset, col_nome in enumerate(colunas_ativas, start=2):
            val = reg['celulas_texto'].get(col_nome)
            cell = ws.cell(row=idx_reg, column=offset)
            cell.border = borda_celula
            cell.number_format = '@'

            if val is not None:
                tipo_cel = reg['celulas_tipo'].get(col_nome, 'div')
                cell.value = val
                if tipo_cel == 'warn':
                    cell.font = font_warn
                    cell.fill = fill_warn
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif tipo_cel == 'obs':
                    cell.font = font_obs
                    cell.fill = fill_obs
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                else:
                    cell.font = font_div
                    cell.fill = fill_div
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.value = '—'
                cell.font = font_ok_dash
                cell.fill = fill_linha
                cell.alignment = Alignment(horizontal='center', vertical='center')

        ws.row_dimensions[idx_reg].height = 22

    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f"A3:{get_column_letter(total_cols)}{len(registros) + 3}"

    for col_i in range(1, total_cols + 1):
        col_letter = get_column_letter(col_i)
        max_len = len(cabecalhos[col_i - 1])
        for row_i in range(4, len(registros) + 4):
            val = ws.cell(row=row_i, column=col_i).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 5, 16), 65)

    _adicionar_bloco_rodape_excel(
        ws,
        start_row=len(registros) + 6,
        total_cols=total_cols,
        dados_adicionais=dados_adicionais
    )


@st.cache_data(show_spinner=False)
def gerar_excel_auditoria_cached(
    itens_xml_tuple,
    itens_csv_tuple,
    numero_nf: str,
    dados_adicionais_tuple,
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
        df_final.to_excel(writer, sheet_name='Comparativo_ICMS_e_EAN', index=False, startrow=2)

        workbook = writer.book
        worksheet = workbook['Comparativo_ICMS_e_EAN']

        _montar_aba_resumo_excel(workbook, dados_relatorio, numero_nf, dados_adicionais)

        worksheet.sheet_view.showGridLines = False
        worksheet.sheet_view.zoomScale = 80

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
        fill_status_alerta = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
        fill_status_convergente = PatternFill(start_color="196F3D", end_color="196F3D", fill_type="solid")
        fill_cel_divergente = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
        fill_cel_alerta = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
        fill_cel_alerta_dado = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        fill_cel_convergente = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")

        font_super_white = Font(bold=True, size=11, color="FFFFFF")
        font_header_white = Font(bold=True, size=10, color="FFFFFF")
        font_header_status = Font(bold=True, size=10, color="FFFFFF")
        font_divergente = Font(color="9C0006", bold=True)
        font_alerta = Font(color="B45309", bold=True)
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
        if inc_icms: super_grupos_def.append(("ALÍQUOTAS DE ICMS", ['Orig.', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq']))
        if inc_red: super_grupos_def.append(("REDUÇÃO DE BASE", ['pRedBC_NF', 'ICMS_Efetivo_NF', '% ICMS red. Whintor', 'Status_Red_BC']))
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
        status_has_divergence = {
            sc: (df_final[sc] == 'DIVERGENTE').any() if sc in df_final.columns else False
            for sc in status_columns
        }
        status_has_warning = {
            'Status_EAN': any(l.get('EAN_Alerta_Vazio') for l in dados_relatorio),
            'Status_ICMS_Aliq': any(l.get('ICMS_Alerta_Guia') for l in dados_relatorio),
        }

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
                elif status_has_warning.get(col_name):
                    cell.fill = fill_status_alerta
                else:
                    cell.fill = fill_status_convergente
                cell.font = font_header_status
            else:
                cell.fill = obter_fill_cabecalho(col_name)
                cell.font = font_header_white

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

            for col_status, colunas_dados in status_dependents.items():
                if col_status in headers:
                    idx_status = headers.index(col_status) + 1
                    cell_status = worksheet.cell(row=row, column=idx_status)
                    val_status = cell_status.value

                    eh_alerta_laranja = (
                        (col_status == 'Status_EAN' and linha_dict.get('EAN_Alerta_Vazio'))
                        or (col_status == 'Status_ICMS_Aliq' and linha_dict.get('ICMS_Alerta_Guia'))
                    )

                    if val_status == 'DIVERGENTE':
                        cell_status.fill = fill_cel_divergente
                        cell_status.font = font_divergente
                        for col_afetada in colunas_dados:
                            if col_afetada in headers:
                                cell_dado = worksheet.cell(row=row, column=headers.index(col_afetada) + 1)
                                cell_dado.font = font_divergente
                    elif eh_alerta_laranja:
                        cell_status.fill = fill_cel_alerta
                        cell_status.font = font_alerta
                        for col_afetada in colunas_dados:
                            if col_afetada in headers:
                                cell_dado = worksheet.cell(row=row, column=headers.index(col_afetada) + 1)
                                cell_dado.font = font_alerta
                                cell_dado.fill = fill_cel_alerta_dado
                    elif val_status == 'OK':
                        cell_status.fill = fill_cel_convergente
                        for col_afetada in colunas_dados:
                            if col_afetada in headers:
                                cell_dado = worksheet.cell(row=row, column=headers.index(col_afetada) + 1)
                                cell_dado.font = font_normal

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
# 3. CALLBACKS DE GERENCIAMENTO DO LOTE
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
# 4. INTERFACE STREAMLIT (UI/UX OTIMIZADA)
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
