import io
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
# CONFIGURAÇÃO DA PÁGINA E ESTILO
# ==========================================
st.set_page_config(page_title="Auditoria de Notas | Dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
        .stApp {
            background-color: #F4F6F9;
            color: #2C3E50;
        }
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }
        .dashboard-header {
            background: linear-gradient(135deg, #1B2631 0%, #2C3E50 100%);
            padding: 16px 20px;
            border-radius: 8px;
            color: #FFFFFF;
            margin-bottom: 15px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.06);
        }
        .dashboard-title {
            font-size: 1.5rem;
            font-weight: 700;
            margin: 0;
            color: #FFFFFF;
        }
        .dashboard-subtitle {
            font-size: 0.85rem;
            color: #BDC3C7;
            margin-top: 3px;
            margin-bottom: 0;
        }
        .card-panel {
            background-color: #FFFFFF;
            border: 1px solid #E5E7E9;
            padding: 14px 18px;
            border-radius: 8px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.03);
            margin-bottom: 12px;
        }
        .card-title {
            font-size: 0.95rem;
            font-weight: 600;
            color: #34495E;
            margin-bottom: 8px;
        }
        /* Barra de validação compacta e executiva */
        .val-strip {
            background-color: #FFFFFF;
            border: 1px solid #DCE1E7;
            border-left: 4px solid #196F3D;
            border-radius: 6px;
            padding: 7px 12px;
            display: flex;
            align-items: center;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03);
            min-height: 40px;
        }
        .val-strip.bloqueado {
            border-left-color: #C0392B;
            background-color: #FDEDEC;
        }
        .v-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 0.82rem;
            font-weight: 600;
            padding: 3px 10px;
            border-radius: 4px;
            background-color: #F4F6F9;
            color: #2C3E50;
            border: 1px solid #E5E8E8;
        }
        .v-badge.ok {
            background-color: #EAFAF1;
            color: #145A32;
            border-color: #ABEBC6;
        }
        .v-badge.warn {
            background-color: #FEF9E7;
            color: #7D6608;
            border-color: #F9E79F;
        }
        .v-badge.err {
            background-color: #FADBD8;
            color: #78281F;
            border-color: #F5B7B1;
        }
        /* Botão de Download em Verde Executivo */
        div[data-testid="stDownloadButton"] > button {
            background-color: #196F3D !important;
            color: #FFFFFF !important;
            border: 1px solid #145A32 !important;
            font-weight: 600 !important;
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
_IGNORAR = {'COM', 'PARA', 'DOS', 'DAS', 'UNID', 'UND', 'CX', 'KIT', 'PCT', 'FR', 'REF', 'ESC', 'ATIV'}


# ==========================================
# 1. FUNÇÕES DE SANEAMENTO, EAN E LEITURA
# ==========================================
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


def ler_csv_winthor(arquivo):
    avisos = []
    nome = getattr(arquivo, 'name', '').lower()
    bruto = arquivo.getvalue()

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


def ler_xml_nfe(arquivo):
    bruto = arquivo.getvalue()
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


# ==========================================
# 2. INTERFACE STREAMLIT (DASHBOARD COMPACTO)
# ==========================================
st.markdown("""
    <div class="dashboard-header">
        <p class="dashboard-title">📊 Painel de Auditoria de Notas Fiscais</p>
        <p class="dashboard-subtitle">Cruzamento automatizado entre XML e relatórios WinThor.</p>
    </div>
""", unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2, gap="medium")

with col_up1:
    st.markdown('<div class="card-panel">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📂 1. Nota Fiscal Eletrônica (XML)</p>', unsafe_allow_html=True)
    arquivo_xml = st.file_uploader("Upload XML", type=["xml"], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with col_up2:
    st.markdown('<div class="card-panel">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📊 2. Relatório de Pré-Entrada (WinThor)</p>', unsafe_allow_html=True)
    arquivo_dados = st.file_uploader("Upload WinThor", type=["csv", "xlsx"], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="card-panel">', unsafe_allow_html=True)
top_col1, top_col2 = st.columns([4, 1])
with top_col1:
    st.markdown('<p class="card-title" style="margin-bottom:0;">⚙️ Seleção de Blocos para Exportação</p>', unsafe_allow_html=True)
with top_col2:
    selecionar_todos = st.checkbox("Selecionar Todos", value=True)

st.write("")
b_col1, b_col2, b_col3, b_col4 = st.columns(4)
with b_col1:
    inc_ident = st.checkbox("Identificação", value=selecionar_todos)
    inc_qtd = st.checkbox("Quantidades", value=selecionar_todos)
with b_col2:
    inc_ean = st.checkbox("Código & EAN", value=selecionar_todos)
    inc_ncm = st.checkbox("NCM (Fiscal)", value=selecionar_todos)
with b_col3:
    inc_desc = st.checkbox("Descrições", value=selecionar_todos)
    inc_icms = st.checkbox("ICMS & Origem", value=selecionar_todos)
with b_col4:
    inc_red = st.checkbox("Redução Base", value=selecionar_todos)
    inc_st = st.checkbox("Cálculo ST", value=selecionar_todos)
st.markdown('</div>', unsafe_allow_html=True)

if arquivo_xml and arquivo_dados:
    numero_nf, itens_xml = ler_xml_nfe(arquivo_xml)
    df_csv, avisos_prep = ler_csv_winthor(arquivo_dados)
    itens_csv = montar_itens_csv(df_csv)
    val = validar_upload(itens_xml, itens_csv, numero_nf, arquivo_dados.name)

    tem_bloqueio = len(val['bloqueios']) > 0
    ignorar_bloqueio_nf = False

    # Exibe alertas APENAS se houver bloqueio ou aviso
    if tem_bloqueio or avisos_prep or val['avisos']:
        for bloq in val['bloqueios']:
            st.markdown(f'<div class="v-badge err" style="width:100%; margin-bottom:6px;">🚫 {bloq}</div>', unsafe_allow_html=True)

        for av in avisos_prep + val['avisos']:
            st.markdown(f'<div class="v-badge warn" style="width:100%; margin-bottom:6px;">⚠️ {av}</div>', unsafe_allow_html=True)

        if val['bloqueio_nf']:
            ignorar_bloqueio_nf = st.checkbox("Confirmar mesmo assim (o arquivo pertence a esta NF)")

    pode_processar = (not tem_bloqueio) or (val['bloqueio_nf'] and ignorar_bloqueio_nf)

    # Linha única compacta: [Status NF menor] [Botão Processar] [Botão Baixar Relatório] [Espaço livre]
    col_strip, col_btn_proc, col_btn_dl, _ = st.columns([1.4, 0.9, 1.3, 1.4], vertical_alignment="center")

    with col_strip:
        if pode_processar:
            st.markdown(f"""
                <div class="val-strip">
                    <span class="v-badge ok">✅ NF {numero_nf} pronta para auditoria</span>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div class="val-strip bloqueado">
                    <span class="v-badge err">⛔ Bloqueado</span>
                </div>
            """, unsafe_allow_html=True)

    with col_btn_proc:
        processar = st.button(
            "🚀 Processar",
            use_container_width=True,
            type="primary",
            disabled=not pode_processar
        )

    if pode_processar and processar:
        with st.spinner("Gerando relatório..."):
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
                        'Origem_XML': '-', 'ICMS_NF': '-', 'pRedBC_NF': '-'
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
                'ICMS & Origem': ['Origem_XML', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq'],
                'Redução de Base': ['pRedBC_NF', '% ICMS red. Whintor', 'Status_Red_BC'],
                'Tipo Cálculo ST': ['Tipo_Calculo_ST']
            }

            colunas_selecionadas = []
            if inc_ident: colunas_selecionadas.extend(mapa_colunas_blocos['Identificação'])
            if inc_ean: colunas_selecionadas.extend(mapa_colunas_blocos['Código & EAN'])
            if inc_desc: colunas_selecionadas.extend(mapa_colunas_blocos['Descrições'])
            if inc_qtd: colunas_selecionadas.extend(mapa_colunas_blocos['Quantidades'])
            if inc_ncm: colunas_selecionadas.extend(mapa_colunas_blocos['NCM (Fiscal)'])
            if inc_icms: colunas_selecionadas.extend(mapa_colunas_blocos['ICMS & Origem'])
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
                if inc_icms: super_grupos_def.append(("ALÍQUOTAS DE ICMS E ORIGEM", ['Origem_XML', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq']))
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
                    else:
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
                    'NCM_NF', 'NCM_WINTHOR', 'Origem_XML', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor',
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
                        elif col_name in ('EAN_NF', 'EAN_WINTHOR', 'Origem_XML') and str(cell.value).strip() not in ('', '-'):
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

            with col_btn_dl:
                st.download_button(
                    label="📥 Baixar Relatório",
                    data=excel_buffer.getvalue(),
                    file_name=nome_arquivo_saida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
