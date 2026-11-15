import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from itertools import zip_longest
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
import re
from datetime import datetime
import io
import html

# ==========================================
# 1. FUNÇÕES AUXILIARES
# ==========================================
def limpar_numero(valor):
    if pd.isna(valor):
        return ""
    val_str = str(valor).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return re.sub(r'[^0-9]', '', val_str)

def limpar_ean(valor):
    if pd.isna(valor):
        return ""
    val_str = str(valor).strip().upper()
    if val_str in ('SEM GTIN', 'S/GTIN', 'NONE', 'NAN', ''):
        return ""
    if val_str.endswith('.0'):
        val_str = val_str[:-2]

    if 'E+' in val_str or 'E-' in val_str:
        try:
            val_limpo = val_str.replace(',', '.')
            val_str = str(int(float(val_limpo)))
        except Exception:
            pass

    apenas_numeros = re.sub(r'[^0-9]', '', val_str)
    return apenas_numeros.lstrip('0')

def converte_float_br(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return 0.0
    try:
        val_str = str(valor).strip()
        if '.' in val_str and ',' in val_str:
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '.')
        return float(val_str)
    except Exception:
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
        except Exception:
            pass
    match_nums = re.findall(r'\d+', str(uCom))
    if match_nums:
        try:
            return float(match_nums[-1])
        except Exception:
            pass
    return 1.0

def formatar_num_br(valor, casas=2, sufixo=""):
    if valor in ('-', '', None) or pd.isna(valor):
        return "-"
    try:
        num = float(valor)
        formato = f"{num:,.{casas}f}"
        formato_br = formato.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{formato_br}{sufixo}"
    except Exception:
        return str(valor)

def formatar_moeda_br(valor):
    if valor in ('-', '', None) or pd.isna(valor):
        return "-"
    return f"R$ {formatar_num_br(valor, 2)}"

def badge_status_html(status_val):
    st_str = str(status_val).strip().upper()
    if st_str == 'OK':
        return '<span class="badge-ok">✓ OK</span>'
    elif st_str == 'DIVERGENTE':
        return '<span class="badge-div">✗ DIVERGENTE</span>'
    return '<span class="badge-neutral">-</span>'

# ==========================================
# 2. INTERFACE STREAMLIT (DASHBOARD COMPACTO E INTERATIVO)
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
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
            margin-bottom: 14px;
        }
        .kpi-card {
            background: #FFFFFF;
            border-radius: 8px;
            padding: 14px 16px;
            border: 1px solid #E5E7E9;
            border-left: 5px solid #2C3E50;
            box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        }
        .kpi-card.ok { border-left-color: #196F3D; }
        .kpi-card.div { border-left-color: #9C0006; }
        .kpi-card.sobra { border-left-color: #D68910; }
        .kpi-label {
            font-size: 0.78rem;
            font-weight: 600;
            text-transform: uppercase;
            color: #7F8C8D;
            margin-bottom: 4px;
        }
        .kpi-value {
            font-size: 1.45rem;
            font-weight: 700;
            color: #1B2631;
            line-height: 1.2;
        }
        .kpi-sub {
            font-size: 0.78rem;
            color: #566573;
            margin-top: 4px;
        }
        .pill-bar {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 14px;
        }
        .err-pill {
            padding: 5px 12px;
            border-radius: 999px;
            font-size: 0.8rem;
            font-weight: 600;
            border: 1px solid transparent;
        }
        .err-pill.zero {
            background-color: #E8F8F5;
            color: #117864;
            border-color: #A3E4D7;
        }
        .err-pill.alert {
            background-color: #FDEDEC;
            color: #9C0006;
            border-color: #F5B7B1;
        }
        .prod-section {
            background-color: #FFFFFF;
            border: 1px solid #E5E7E9;
            border-radius: 8px;
            padding: 12px 14px;
            height: 100%;
        }
        .prod-section.div-border {
            border: 1.5px solid #E6B0AA;
            background-color: #FFFDFD;
        }
        .prod-section.ok-border {
            border: 1px solid #D4EFDF;
        }
        .prod-sec-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.85rem;
            font-weight: 700;
            color: #1B2631;
            border-bottom: 1px solid #EAECEE;
            padding-bottom: 6px;
            margin-bottom: 8px;
        }
        .field-row {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            padding: 4px 0;
            border-bottom: 1px dashed #F2F4F4;
            font-size: 0.83rem;
        }
        .field-row:last-child {
            border-bottom: none;
        }
        .field-label {
            color: #566573;
            font-weight: 500;
            padding-right: 8px;
        }
        .field-val {
            font-weight: 600;
            color: #1B2631;
            text-align: right;
        }
        .field-val.val-div {
            color: #9C0006;
            font-weight: 700;
            background-color: #FADBD8;
            padding: 1px 6px;
            border-radius: 4px;
        }
        .field-val.val-ok {
            color: #196F3D;
            font-weight: 700;
        }
        .badge-ok {
            background-color: #D4EFDF;
            color: #196F3D;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.74rem;
            font-weight: 700;
        }
        .badge-div {
            background-color: #FADBD8;
            color: #9C0006;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.74rem;
            font-weight: 700;
        }
        .badge-neutral {
            background-color: #EAEDED;
            color: #566573;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.74rem;
            font-weight: 600;
        }
        .desc-box {
            background-color: #F8F9F9;
            border: 1px solid #E5E7E9;
            border-radius: 6px;
            padding: 8px 12px;
            margin-bottom: 10px;
            font-size: 0.84rem;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="dashboard-header">
        <p class="dashboard-title">📊 Painel de Auditoria de Notas Fiscais</p>
        <p class="dashboard-subtitle">Cruzamento automatizado entre XML e relatórios WinThor com inspeção interativa e exportação executiva.</p>
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

# Seção de Seleção de Blocos com a opção "Marcar/Desmarcar Todos"
st.markdown('<div class="card-panel">', unsafe_allow_html=True)
top_col1, top_col2 = st.columns([4, 1])
with top_col1:
    st.markdown('<p class="card-title" style="margin-bottom:0;">⚙️ Seleção de Blocos para Visualização e Exportação</p>', unsafe_allow_html=True)
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


# ==========================================
# FUNÇÃO DE PROCESSAMENTO COMPLETO (XML + WINTHOR)
# ==========================================
def executar_auditoria(xml_bytes, dados_bytes, nome_dados):
    ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
    tree = ET.parse(io.BytesIO(xml_bytes))
    root = tree.getroot()

    numero_nf = "DESCONHECIDO"
    nNF_node = root.find('.//ns:ide/ns:nNF', ns)
    if nNF_node is not None and nNF_node.text:
        numero_nf = nNF_node.text

    emit_nome = ""
    xNome_node = root.find('.//ns:emit/ns:xNome', ns)
    if xNome_node is not None and xNome_node.text:
        emit_nome = xNome_node.text

    itens_xml = []
    for det in root.findall('.//ns:det', ns):
        prod = det.find('ns:prod', ns)

        try:
            nItem = int(det.attrib.get('nItem', '0'))
        except Exception:
            nItem = det.attrib.get('nItem', '')

        desc = prod.find('ns:xProd', ns).text if prod.find('ns:xProd', ns) is not None else ''
        ean_raw = prod.find('ns:cEAN', ns).text if prod.find('ns:cEAN', ns) is not None else ''
        ean_trib_raw = prod.find('ns:cEANTrib', ns).text if prod.find('ns:cEANTrib', ns) is not None else ''

        ean_exibicao = limpar_numero(ean_raw) if limpar_ean(ean_raw) != '' else limpar_numero(ean_trib_raw)
        ean_comp = limpar_ean(ean_raw)
        ean_trib_comp = limpar_ean(ean_trib_raw)

        ncm = limpar_numero(prod.find('ns:NCM', ns).text) if prod.find('ns:NCM', ns) is not None else ''

        uCom = prod.find('ns:uCom', ns).text if prod.find('ns:uCom', ns) is not None else 'UN'
        qCom_xml = float(prod.find('ns:qCom', ns).text) if prod.find('ns:qCom', ns) is not None else 0.0

        fator = extrair_fator_embalagem(uCom)

        vUn = float(prod.find('ns:vUnCom', ns).text) if prod.find('ns:vUnCom', ns) is not None else 0.0
        vTot = float(prod.find('ns:vProd', ns).text) if prod.find('ns:vProd', ns) is not None else 0.0

        icms_node = det.find('.//ns:ICMS', ns)
        icms = 0.0
        origem_desc = '0 - Nacional'

        if icms_node is not None and len(icms_node) > 0:
            sub_icms = icms_node[0]
            pICMS_node = sub_icms.find('ns:pICMS', ns)
            if pICMS_node is not None and pICMS_node.text:
                icms = float(pICMS_node.text)

            orig_node = sub_icms.find('ns:orig', ns)
            if orig_node is not None and orig_node.text:
                origem_desc = obter_descricao_origem(orig_node.text)

        pRedBC_node = det.find('.//ns:pRedBC', ns)
        pRedBC = float(pRedBC_node.text) if pRedBC_node is not None and pRedBC_node.text else 0.0

        itens_xml.append({
            'Linha_NF': nItem,
            'Descricao_NF': desc,
            'EAN_NF': ean_exibicao,
            'EAN_Raw_NF': ean_raw,
            'EANTrib_Raw_NF': ean_trib_raw,
            'EAN_Comp_NF': ean_comp,
            'EANTrib_Comp_NF': ean_trib_comp,
            'NCM_NF': ncm,
            'Unid_XML': uCom,
            'Qtd_NF': qCom_xml,
            '_Fator_Embalagem': fator,
            '_Qtd_Convertida_XML': round(qCom_xml * fator, 4),
            'Vl_Unit_NF': vUn,
            'Vl_Total_NF': vTot,
            'Origem_XML': origem_desc,
            'ICMS_NF': icms,
            'pRedBC_NF': pRedBC
        })

    # 4. EXTRAÇÃO DA PLANILHA (WINTHOR)
    buffer_dados = io.BytesIO(dados_bytes)
    if nome_dados.lower().endswith('.xlsx'):
        df_csv = pd.read_excel(buffer_dados, dtype=str)
    else:
        df_csv = pd.read_csv(buffer_dados, sep=';', encoding='latin1', dtype=str)

    df_csv.columns = df_csv.columns.str.strip()

    col_tipo_calc = next((c for c in df_csv.columns if 'tipo' in c.lower() and ('cálculo' in c.lower() or 'calculo' in c.lower())), None)
    if not col_tipo_calc:
        col_tipo_calc = df_csv.columns[-2]

    itens_csv = []
    for index, row in df_csv.iterrows():
        ean_win_raw = row.get('Cód. barra', '')
        itens_csv.append({
            'Linha_WINTHOR': index + 1,
            'Cod_Produto_WINTHOR': limpar_numero(row.get('Cód. produto', '')),
            'EAN_WINTHOR': limpar_numero(ean_win_raw),
            'EAN_Comp_WINTHOR': limpar_ean(ean_win_raw),
            'Descricao_WINTHOR': str(row.get('Descrição', '')).strip() if not pd.isna(row.get('Descrição', '')) else '',
            'Qtd_WINTHOR': converte_float_br(row.get('Qtde Pré-Entrada', 0.0)),
            'NCM_WINTHOR': limpar_numero(row.get('NCM', '')),
            'ICMS_WINTHOR': converte_float_br(row.get('%ICMS', 0.0)),
            '% ICMS red. Whintor': converte_float_br(row.get('% ICMS red.', 0.0)),
            '%Alíq.Ext. ST Guia_Whintor': converte_float_br(row.get('%Alíq.Ext.ST Guia', 0.0)),
            'Tipo_Calculo_WINTHOR': str(row.get(col_tipo_calc, '')).strip() if not pd.isna(row.get(col_tipo_calc, '')) else ''
        })

    # 5. CRUZAMENTO BASEADO ESTRITAMENTE NA ORDEM DO XML
    dados_relatorio = []

    for item_x, item_c in zip_longest(itens_xml, itens_csv, fillvalue=None):
        linha = {}

        if item_x:
            linha.update(item_x)
        else:
            linha.update({
                'Linha_NF': '-', 'Descricao_NF': '-', 'EAN_NF': '-', 'EAN_Raw_NF': '-', 'EANTrib_Raw_NF': '-',
                'NCM_NF': '-', 'Unid_XML': '-', 'Qtd_NF': '-', '_Fator_Embalagem': 1.0, '_Qtd_Convertida_XML': '-',
                'Vl_Unit_NF': '-', 'Vl_Total_NF': '-', 'Origem_XML': '-', 'ICMS_NF': '-', 'pRedBC_NF': '-'
            })

        if item_c:
            linha.update(item_c)
        else:
            linha.update({
                'Linha_WINTHOR': '-', 'Cod_Produto_WINTHOR': '-', 'EAN_WINTHOR': '-', 'Descricao_WINTHOR': '-',
                'Qtd_WINTHOR': '-', 'NCM_WINTHOR': '-', 'ICMS_WINTHOR': '-', '% ICMS red. Whintor': '-',
                '%Alíq.Ext. ST Guia_Whintor': '-', 'Tipo_Calculo_WINTHOR': '-'
            })

        lista_erros_item = []

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
                lista_erros_item.append('EAN')

            icms_xml = round(item_x['ICMS_NF'], 2)
            icms_csv = round(item_c['ICMS_WINTHOR'], 2)
            st_guia = round(item_c['%Alíq.Ext. ST Guia_Whintor'], 2)
            if icms_xml == icms_csv or icms_xml == st_guia:
                linha['Status_ICMS_Aliq'] = 'OK'
            else:
                linha['Status_ICMS_Aliq'] = 'DIVERGENTE'
                lista_erros_item.append('ICMS Alíquota')

            pred_xml = round(item_x['pRedBC_NF'], 4)
            red_csv = round(item_c['% ICMS red. Whintor'], 4)
            carga_efetiva_xml = round(item_x['ICMS_NF'] * (1.0 - (item_x['pRedBC_NF'] / 100.0)), 4)
            linha['_Carga_Efetiva_XML'] = carga_efetiva_xml

            if (
                carga_efetiva_xml == red_csv
                or round(carga_efetiva_xml, 2) == round(red_csv, 2)
                or pred_xml == red_csv
            ):
                linha['Status_Red_BC'] = 'OK'
            else:
                linha['Status_Red_BC'] = 'DIVERGENTE'
                lista_erros_item.append('Redução de Base')

            tipo_calc = str(item_c.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
            if tipo_calc == 'N':
                linha['Tipo_Calculo_ST'] = 'NOTA FISCAL'
            elif tipo_calc == 'G':
                linha['Tipo_Calculo_ST'] = 'GUIA'
            else:
                linha['Tipo_Calculo_ST'] = item_c['Tipo_Calculo_WINTHOR']

            linha['Status_NCM'] = 'OK' if str(item_x['NCM_NF']) == str(item_c['NCM_WINTHOR']) else 'DIVERGENTE'
            if linha['Status_NCM'] == 'DIVERGENTE':
                lista_erros_item.append('NCM')

            qtd_xml_pura = item_x['Qtd_NF']
            qtd_win = item_c['Qtd_WINTHOR']
            fator = item_x['_Fator_Embalagem']
            qtd_xml_convertida = qtd_xml_pura * fator

            if round(qtd_xml_pura, 4) == round(qtd_win, 4) or round(qtd_xml_convertida, 4) == round(qtd_win, 4):
                linha['Status_Qtd'] = 'OK'
            else:
                linha['Status_Qtd'] = 'DIVERGENTE'
                lista_erros_item.append('Quantidade')

        elif item_x and not item_c:
            linha['Status_Geral'] = 'SOBRA NF (SEM PAR)'
            linha['_Carga_Efetiva_XML'] = '-'
            for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                linha[st_col] = '-'
            lista_erros_item.append('Sobra NF')

        else:
            linha['Status_Geral'] = 'SOBRA WINTHOR (SEM PAR)'
            linha['_Carga_Efetiva_XML'] = '-'
            for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                linha[st_col] = '-'
            lista_erros_item.append('Sobra WinThor')

        linha['_Erros_Item'] = lista_erros_item
        linha['_Qtd_Erros'] = len(lista_erros_item)
        dados_relatorio.append(linha)

    data_simples = datetime.now().strftime('%Y%m%d')
    nome_arquivo_saida = f'relatorio_NF_{numero_nf}_{data_simples}.xlsx'

    return {
        'numero_nf': numero_nf,
        'emit_nome': emit_nome,
        'nome_arquivo_saida': nome_arquivo_saida,
        'qtd_xml': len(itens_xml),
        'qtd_csv': len(itens_csv),
        'dados_relatorio': dados_relatorio
    }


def gerar_excel_formatado(df_final, dados_relatorio, super_grupos_def):
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
    return excel_buffer.getvalue()


def estilizar_dataframe_streamlit(df_view, dados_filtrados):
    status_dependents = {
        'Status_EAN': ['EAN_NF', 'EAN_WINTHOR'],
        'Status_Qtd': ['Qtd_NF', 'Qtd_WINTHOR'],
        'Status_NCM': ['NCM_NF', 'NCM_WINTHOR'],
        'Status_ICMS_Aliq': ['ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor'],
        'Status_Red_BC': ['pRedBC_NF', '% ICMS red. Whintor']
    }

    def style_rows(row):
        idx_pos = row.name
        item_raw = dados_filtrados[idx_pos] if idx_pos < len(dados_filtrados) else {}
        styles = ['' for _ in row.index]

        if 'SOBRA' in str(item_raw.get('Status_Geral', '')):
            return ['background-color: #EAEDED; color: #566573; font-style: italic;' for _ in row.index]

        col_to_pos = {c: i for i, c in enumerate(row.index)}

        for st_col, dep_cols in status_dependents.items():
            st_val = item_raw.get(st_col, '')
            if st_col in col_to_pos:
                if st_val == 'DIVERGENTE':
                    styles[col_to_pos[st_col]] = 'background-color: #FADBD8; color: #9C0006; font-weight: 700;'
                elif st_val == 'OK':
                    styles[col_to_pos[st_col]] = 'background-color: #D4EFDF; color: #196F3D; font-weight: 700;'

            if st_val == 'DIVERGENTE':
                for dc in dep_cols:
                    if dc in col_to_pos:
                        styles[col_to_pos[dc]] = 'background-color: #FDEDEC; color: #9C0006; font-weight: 700;'

        return styles

    format_dict = {}
    for c in df_view.columns:
        if 'pRedBC' in c or 'red. Whintor' in c:
            format_dict[c] = lambda v: formatar_num_br(v, 4)
        elif 'Qtd' in c or 'ICMS' in c or 'Whintor' in c:
            format_dict[c] = lambda v: formatar_num_br(v, 2)

    return df_view.style.apply(style_rows, axis=1).format(format_dict)


def renderizar_card_produto(item):
    status_geral = str(item.get('Status_Geral', ''))
    erros = item.get('_Erros_Item', [])
    is_sobra = 'SOBRA' in status_geral

    desc_nf = html.escape(str(item.get('Descricao_NF', '-')))
    desc_win = html.escape(str(item.get('Descricao_WINTHOR', '-')))
    st.markdown(f"""
        <div class="desc-box">
            <div><strong>📄 Descrição no XML (Fornecedor):</strong> {desc_nf}</div>
            <div style="margin-top:4px;"><strong>🏢 Descrição no WinThor (ERP):</strong> {desc_win}</div>
        </div>
    """, unsafe_allow_html=True)

    if is_sobra:
        st.warning(f"⚠️ **{status_geral}** — Este item existe apenas em uma das pontas do cruzamento posicional e não possui par na linha correspondente.")

    c1, c2, c3, c4 = st.columns(4, gap="small")

    # 1. Card Código, EAN e NCM
    with c1:
        st_ean = item.get('Status_EAN', '-')
        st_ncm = item.get('Status_NCM', '-')
        tem_div_c1 = (st_ean == 'DIVERGENTE') or (st_ncm == 'DIVERGENTE')
        cls_c1 = "div-border" if tem_div_c1 else ("ok-border" if not is_sobra else "")
        cls_ean = "val-div" if st_ean == 'DIVERGENTE' else ("val-ok" if st_ean == 'OK' else "")
        cls_ncm = "val-div" if st_ncm == 'DIVERGENTE' else ("val-ok" if st_ncm == 'OK' else "")

        st.markdown(f"""
            <div class="prod-section {cls_c1}">
                <div class="prod-sec-header">
                    <span>🏷️ Código, EAN & NCM</span>
                    {badge_status_html('DIVERGENTE' if tem_div_c1 else ('OK' if not is_sobra else '-'))}
                </div>
                <div class="field-row">
                    <span class="field-label">Cód. WinThor:</span>
                    <span class="field-val">{html.escape(str(item.get('Cod_Produto_WINTHOR', '-')))}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">EAN (XML):</span>
                    <span class="field-val {cls_ean}">{html.escape(str(item.get('EAN_NF') or 'SEM GTIN'))}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">EAN (WinThor):</span>
                    <span class="field-val {cls_ean}">{html.escape(str(item.get('EAN_WINTHOR') or 'VAZIO'))}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Status EAN:</span>
                    <span class="field-val">{badge_status_html(st_ean)}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">NCM (XML vs ERP):</span>
                    <span class="field-val {cls_ncm}">{html.escape(str(item.get('NCM_NF', '-')))} | {html.escape(str(item.get('NCM_WINTHOR', '-')))}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Status NCM:</span>
                    <span class="field-val">{badge_status_html(st_ncm)}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

    # 2. Card Quantidades, Embalagem e Valores
    with c2:
        st_qtd = item.get('Status_Qtd', '-')
        cls_c2 = "div-border" if st_qtd == 'DIVERGENTE' else ("ok-border" if st_qtd == 'OK' else "")
        cls_qtd = "val-div" if st_qtd == 'DIVERGENTE' else ("val-ok" if st_qtd == 'OK' else "")
        fator_emb = item.get('_Fator_Embalagem', 1.0)
        unid_xml = html.escape(str(item.get('Unid_XML', '-')))

        st.markdown(f"""
            <div class="prod-section {cls_c2}">
                <div class="prod-sec-header">
                    <span>📦 Quantidades & Valores</span>
                    {badge_status_html(st_qtd)}
                </div>
                <div class="field-row">
                    <span class="field-label">Unid. / Fator XML:</span>
                    <span class="field-val">{unid_xml} (×{formatar_num_br(fator_emb, 0)})</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Qtd Pura (XML):</span>
                    <span class="field-val {cls_qtd}">{formatar_num_br(item.get('Qtd_NF'), 2)}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Qtd Convertida (XML):</span>
                    <span class="field-val {cls_qtd}">{formatar_num_br(item.get('_Qtd_Convertida_XML'), 2)}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Qtd WinThor:</span>
                    <span class="field-val {cls_qtd}">{formatar_num_br(item.get('Qtd_WINTHOR'), 2)}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Vl. Unitário NF:</span>
                    <span class="field-val">{formatar_moeda_br(item.get('Vl_Unit_NF'))}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Vl. Total NF:</span>
                    <span class="field-val">{formatar_moeda_br(item.get('Vl_Total_NF'))}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

    # 3. Card Tributação (ICMS, Origem e ST)
    with c3:
        st_icms = item.get('Status_ICMS_Aliq', '-')
        cls_c3 = "div-border" if st_icms == 'DIVERGENTE' else ("ok-border" if st_icms == 'OK' else "")
        cls_icms = "val-div" if st_icms == 'DIVERGENTE' else ("val-ok" if st_icms == 'OK' else "")

        st.markdown(f"""
            <div class="prod-section {cls_c3}">
                <div class="prod-sec-header">
                    <span>💰 ICMS, Origem & ST</span>
                    {badge_status_html(st_icms)}
                </div>
                <div class="field-row">
                    <span class="field-label">Origem (XML):</span>
                    <span class="field-val">{html.escape(str(item.get('Origem_XML', '-')))}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">ICMS NF (XML):</span>
                    <span class="field-val {cls_icms}">{formatar_num_br(item.get('ICMS_NF'), 2, '%')}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">%ICMS WinThor:</span>
                    <span class="field-val {cls_icms}">{formatar_num_br(item.get('ICMS_WINTHOR'), 2, '%')}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">%Alíq. Ext. ST Guia:</span>
                    <span class="field-val {cls_icms}">{formatar_num_br(item.get('%Alíq.Ext. ST Guia_Whintor'), 2, '%')}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Tipo Cálculo ST:</span>
                    <span class="field-val">{html.escape(str(item.get('Tipo_Calculo_ST', '-')))}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

    # 4. Card Redução de Base & Carga Efetiva
    with c4:
        st_red = item.get('Status_Red_BC', '-')
        cls_c4 = "div-border" if st_red == 'DIVERGENTE' else ("ok-border" if st_red == 'OK' else "")
        cls_red = "val-div" if st_red == 'DIVERGENTE' else ("val-ok" if st_red == 'OK' else "")

        st.markdown(f"""
            <div class="prod-section {cls_c4}">
                <div class="prod-sec-header">
                    <span>📉 Redução de Base</span>
                    {badge_status_html(st_red)}
                </div>
                <div class="field-row">
                    <span class="field-label">pRedBC (XML):</span>
                    <span class="field-val {cls_red}">{formatar_num_br(item.get('pRedBC_NF'), 4, '%')}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Carga Efetiva XML:</span>
                    <span class="field-val {cls_red}">{formatar_num_br(item.get('_Carga_Efetiva_XML'), 4, '%')}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">% ICMS Red. WinThor:</span>
                    <span class="field-val {cls_red}">{formatar_num_br(item.get('% ICMS red. Whintor'), 4, '%')}</span>
                </div>
                <div class="field-row">
                    <span class="field-label">Fórmula Aplicada:</span>
                    <span class="field-val" style="font-size:0.75rem; color:#566573;">ICMS × (1 - pRedBC/100)</span>
                </div>
            </div>
        """, unsafe_allow_html=True)


# ==========================================
# EXECUÇÃO E EXIBIÇÃO DOS RESULTADOS
# ==========================================
if arquivo_xml and arquivo_dados:
    assinatura_arquivos = f"{arquivo_xml.name}_{arquivo_xml.size}_{arquivo_dados.name}_{arquivo_dados.size}"

    col_run1, col_run2 = st.columns([3, 1])
    with col_run1:
        st.success("✅ Arquivos prontos para processamento cruzado.")
    with col_run2:
        processar = st.button("🚀 Processar Auditoria", use_container_width=True, type="primary")

    if processar or st.session_state.get('assinatura_arquivos') == assinatura_arquivos:
        if processar or 'resultado_auditoria' not in st.session_state or st.session_state.get('assinatura_arquivos') != assinatura_arquivos:
            with st.spinner("Analisando itens e cruzando tributações com base no XML..."):
                st.session_state['resultado_auditoria'] = executar_auditoria(
                    arquivo_xml.getvalue(),
                    arquivo_dados.getvalue(),
                    arquivo_dados.name
                )
                st.session_state['assinatura_arquivos'] = assinatura_arquivos

        res = st.session_state['resultado_auditoria']
        dados_relatorio = res['dados_relatorio']
        numero_nf = res['numero_nf']
        emit_nome = res['emit_nome']
        nome_arquivo_saida = res['nome_arquivo_saida']

        # Filtragem de colunas conforme os blocos selecionados
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

        super_grupos_def = []
        if inc_ident: super_grupos_def.append(("IDENTIFICAÇÃO", ['Linha_NF', 'Linha_WINTHOR']))
        if inc_ean: super_grupos_def.append(("CÓDIGO E EAN", ['Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN']))
        if inc_desc: super_grupos_def.append(("DESCRIÇÕES", ['Descricao_NF', 'Descricao_WINTHOR']))
        if inc_qtd: super_grupos_def.append(("QUANTIDADES", ['Qtd_NF', 'Qtd_WINTHOR', 'Status_Qtd']))
        if inc_ncm: super_grupos_def.append(("CLASSIFICAÇÃO FISCAL (NCM)", ['NCM_NF', 'NCM_WINTHOR', 'Status_NCM']))
        if inc_icms: super_grupos_def.append(("ALÍQUOTAS DE ICMS E ORIGEM", ['Origem_XML', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq']))
        if inc_red: super_grupos_def.append(("REDUÇÃO DE BASE", ['pRedBC_NF', '% ICMS red. Whintor', 'Status_Red_BC']))
        if inc_st: super_grupos_def.append(("TIPO CÁLCULO ST", ['Tipo_Calculo_ST']))

        df_completo = pd.DataFrame(dados_relatorio)
        df_final = df_completo[[c for c in colunas_selecionadas if c in df_completo.columns]].copy()

        # Geração do arquivo Excel em memória
        excel_bytes = gerar_excel_formatado(df_final, dados_relatorio, super_grupos_def)

        # Cálculo de métricas de resumo para os Cards de Topo
        total_linhas = len(dados_relatorio)
        itens_pareados = [d for d in dados_relatorio if d['Status_Geral'] == 'PAREADO']
        itens_100_ok = [d for d in itens_pareados if d['_Qtd_Erros'] == 0]
        itens_com_div = [d for d in itens_pareados if d['_Qtd_Erros'] > 0]
        itens_sobra = [d for d in dados_relatorio if 'SOBRA' in str(d['Status_Geral'])]

        div_ean = sum(1 for d in itens_pareados if d.get('Status_EAN') == 'DIVERGENTE')
        div_qtd = sum(1 for d in itens_pareados if d.get('Status_Qtd') == 'DIVERGENTE')
        div_ncm = sum(1 for d in itens_pareados if d.get('Status_NCM') == 'DIVERGENTE')
        div_icms = sum(1 for d in itens_pareados if d.get('Status_ICMS_Aliq') == 'DIVERGENTE')
        div_red = sum(1 for d in itens_pareados if d.get('Status_Red_BC') == 'DIVERGENTE')

        st.markdown("---")

        # Barra superior de cabeçalho da NF + Botão de Download Imediato
        top_res1, top_res2 = st.columns([3, 1.3])
        with top_res1:
            subtitulo_emit = f" — Emitente: **{emit_nome}**" if emit_nome else ""
            st.markdown(f"### 🧾 Resultado da Auditoria — NF Nº **{numero_nf}**{subtitulo_emit}")
        with top_res2:
            st.download_button(
                label="📥 Baixar Relatório Executivo em Excel",
                data=excel_bytes,
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_download_top"
            )

        # KPIs Executivos
        st.markdown(f"""
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Volumetria (XML vs ERP)</div>
                    <div class="kpi-value">{res['qtd_xml']} <span style="font-size:0.95rem;color:#7F8C8D;">NF</span> / {res['qtd_csv']} <span style="font-size:0.95rem;color:#7F8C8D;">WinThor</span></div>
                    <div class="kpi-sub">{len(itens_pareados)} itens pareados linha a linha</div>
                </div>
                <div class="kpi-card ok">
                    <div class="kpi-label">Itens 100% Conformes</div>
                    <div class="kpi-value" style="color:#196F3D;">{len(itens_100_ok)}</div>
                    <div class="kpi-sub">Sem nenhuma divergência fiscal ou cadastral</div>
                </div>
                <div class="kpi-card div">
                    <div class="kpi-label">Itens com Divergência</div>
                    <div class="kpi-value" style="color:#9C0006;">{len(itens_com_div)}</div>
                    <div class="kpi-sub">Total de {div_ean + div_qtd + div_ncm + div_icms + div_red} ocorrências detectadas</div>
                </div>
                <div class="kpi-card sobra">
                    <div class="kpi-label">Itens Sem Par (Sobras)</div>
                    <div class="kpi-value" style="color:#D68910;">{len(itens_sobra)}</div>
                    <div class="kpi-sub">Diferença de linhas entre XML e WinThor</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # Pílulas de resumo por categoria de divergência
        pills_html = [
            f'<span class="err-pill {"alert" if div_ean > 0 else "zero"}">Código/EAN: {div_ean}</span>',
            f'<span class="err-pill {"alert" if div_qtd > 0 else "zero"}">Quantidade: {div_qtd}</span>',
            f'<span class="err-pill {"alert" if div_ncm > 0 else "zero"}">NCM: {div_ncm}</span>',
            f'<span class="err-pill {"alert" if div_icms > 0 else "zero"}">ICMS Alíquota: {div_icms}</span>',
            f'<span class="err-pill {"alert" if div_red > 0 else "zero"}">Redução de Base: {div_red}</span>',
            f'<span class="err-pill {"alert" if len(itens_sobra) > 0 else "zero"}">Sobras: {len(itens_sobra)}</span>',
        ]
        st.markdown(f'<div class="pill-bar">{"".join(pills_html)}</div>', unsafe_allow_html=True)

        # ==========================================
        # ABAS INTERATIVAS DE VISUALIZAÇÃO NO FRONT
        # ==========================================
        tab_cards, tab_tabela, tab_raiox = st.tabs([
            f"🚨 Cards de Inspeção por Produto ({len(itens_com_div) + len(itens_sobra)} com alerta)",
            f"📋 Tabela Comparativa Completa ({total_linhas} linhas)",
            "🔍 Raio-X Individual do Produto"
        ])

        # ------------------------------------------
        # ABA 1: CARDS EXPANSÍVEIS COM DESTAQUE DE DIVERGÊNCIAS
        # ------------------------------------------
        with tab_cards:
            f_col1, f_col2, f_col3, f_col4 = st.columns([1.6, 1.8, 1.8, 1.1])
            with f_col1:
                padrao_idx = 0 if (len(itens_com_div) + len(itens_sobra)) > 0 else 1
                modo_visao = st.selectbox(
                    "Exibir itens:",
                    [
                        "Apenas Divergentes e Sobras",
                        "Todos os Itens da Nota",
                        "Apenas Itens 100% OK",
                        "Apenas Sobras (Sem Par)"
                    ],
                    index=padrao_idx
                )
            with f_col2:
                filtro_tipos = st.multiselect(
                    "Filtrar por tipo de divergência:",
                    ["EAN", "Quantidade", "NCM", "ICMS Alíquota", "Redução de Base", "Sobra NF", "Sobra WinThor"],
                    placeholder="Todos os tipos..."
                )
            with f_col3:
                termo_busca = st.text_input(
                    "Buscar produto (Cód. WinThor, EAN, NCM ou Descrição):",
                    placeholder="Digite para filtrar..."
                ).strip().lower()
            with f_col4:
                st.write("")
                expandir_todos = st.checkbox(
                    "Expandir todos",
                    value=(len(itens_com_div) + len(itens_sobra) <= 8 and modo_visao == "Apenas Divergentes e Sobras")
                )

            itens_exibidos = []
            for d in dados_relatorio:
                tem_erro = d['_Qtd_Erros'] > 0
                eh_sobra = 'SOBRA' in str(d['Status_Geral'])

                if modo_visao == "Apenas Divergentes e Sobras" and not tem_erro:
                    continue
                elif modo_visao == "Apenas Itens 100% OK" and (tem_erro or eh_sobra):
                    continue
                elif modo_visao == "Apenas Sobras (Sem Par)" and not eh_sobra:
                    continue

                if filtro_tipos:
                    if not any(tp in d['_Erros_Item'] for tp in filtro_tipos):
                        continue

                if termo_busca:
                    texto_item = " ".join([
                        str(d.get('Linha_NF', '')),
                        str(d.get('Cod_Produto_WINTHOR', '')),
                        str(d.get('EAN_NF', '')),
                        str(d.get('EAN_WINTHOR', '')),
                        str(d.get('NCM_NF', '')),
                        str(d.get('Descricao_NF', '')),
                        str(d.get('Descricao_WINTHOR', ''))
                    ]).lower()
                    if termo_busca not in texto_item:
                        continue

                itens_exibidos.append(d)

            if not itens_exibidos:
                st.info("🎉 Nenhum item encontrado para os filtros selecionados.")
            else:
                st.caption(f"Exibindo **{len(itens_exibidos)}** de **{total_linhas}** itens. Clique em qualquer card abaixo para expandir ou recolher os detalhes do produto:")

                for idx_card, item in enumerate(itens_exibidos):
                    erros = item['_Erros_Item']
                    eh_sobra = 'SOBRA' in str(item['Status_Geral'])
                    linha_nf = item.get('Linha_NF', '-')
                    linha_win = item.get('Linha_WINTHOR', '-')
                    cod_win = item.get('Cod_Produto_WINTHOR', '-')
                    nome_prod = item.get('Descricao_NF') if item.get('Descricao_NF') != '-' else item.get('Descricao_WINTHOR', '-')

                    if eh_sobra:
                        icone = "🟠"
                        tag_resumo = f"[{item['Status_Geral']}]"
                    elif len(erros) > 0:
                        icone = "🔴"
                        tag_resumo = f"⚠️ {len(erros)} Divergência(s): {', '.join(erros)}"
                    else:
                        icone = "🟢"
                        tag_resumo = "✓ 100% OK"

                    titulo_card = (
                        f"{icone} Linha NF #{linha_nf} (ERP #{linha_win}) | "
                        f"Cód. WinThor: {cod_win} — {nome_prod}   •   {tag_resumo}"
                    )

                    abrir_card = expandir_todos or (idx_card == 0 and len(erros) > 0 and len(itens_exibidos) <= 15)
                    with st.expander(titulo_card, expanded=abrir_card):
                        renderizar_card_produto(item)

        # ------------------------------------------
        # ABA 2: TABELA COMPARATIVA INTERATIVA COLORIDA
        # ------------------------------------------
        with tab_tabela:
            t_col1, t_col2 = st.columns([1.5, 2.5])
            with t_col1:
                filtro_tab = st.radio(
                    "Filtro rápido da tabela:",
                    ["Todos os Itens", "Somente com Divergência / Sobra", "Somente 100% OK"],
                    horizontal=True
                )
            with t_col2:
                busca_tab = st.text_input(
                    "Filtrar linhas na tabela (Código, EAN ou Descrição):",
                    placeholder="Digite para filtrar a tabela...",
                    key="busca_tabela"
                ).strip().lower()

            indices_filtrados = []
            for i, d in enumerate(dados_relatorio):
                if filtro_tab == "Somente com Divergência / Sobra" and d['_Qtd_Erros'] == 0:
                    continue
                elif filtro_tab == "Somente 100% OK" and (d['_Qtd_Erros'] > 0 or 'SOBRA' in str(d['Status_Geral'])):
                    continue

                if busca_tab:
                    blob = f"{d.get('Cod_Produto_WINTHOR','')} {d.get('EAN_NF','')} {d.get('EAN_WINTHOR','')} {d.get('Descricao_NF','')} {d.get('Descricao_WINTHOR','')}".lower()
                    if busca_tab not in blob:
                        continue
                indices_filtrados.append(i)

            if not indices_filtrados:
                st.info("Nenhuma linha corresponde ao filtro selecionado.")
            else:
                df_tab_view = df_final.iloc[indices_filtrados].reset_index(drop=True)
                dados_tab_filtrados = [dados_relatorio[i] for i in indices_filtrados]
                styled_df = estilizar_dataframe_streamlit(df_tab_view, dados_tab_filtrados)
                st.dataframe(styled_df, use_container_width=True, height=520)

        # ------------------------------------------
        # ABA 3: RAIO-X INDIVIDUAL DO PRODUTO (MEMÓRIA DE CÁLCULO)
        # ------------------------------------------
        with tab_raiox:
            opcoes_itens = []
            for i, d in enumerate(dados_relatorio):
                status_tag = f"🔴 {', '.join(d['_Erros_Item'])}" if d['_Qtd_Erros'] > 0 else "🟢 100% OK"
                nome_ref = d.get('Descricao_NF') if d.get('Descricao_NF') != '-' else d.get('Descricao_WINTHOR', '-')
                opcoes_itens.append(
                    (i, f"Linha NF #{d.get('Linha_NF')} | Cód. ERP: {d.get('Cod_Produto_WINTHOR')} — {nome_ref} ({status_tag})")
                )

            item_escolhido = st.selectbox(
                "Selecione um produto da Nota Fiscal para inspecionar todos os critérios e cálculos:",
                options=opcoes_itens,
                format_func=lambda x: x[1]
            )

            if item_escolhido is not None:
                item_sel = dados_relatorio[item_escolhido[0]]
                renderizar_card_produto(item_sel)

                st.markdown("#### 🧮 Memória de Cálculo e Critérios Aplicados no Item")
                linhas_memoria = [
                    {
                        "Critério Auditado": "1. Código de Barras (EAN)",
                        "Valor no XML (NF-e)": f"cEAN: {item_sel.get('EAN_Raw_NF', '-')} | cEANTrib: {item_sel.get('EANTrib_Raw_NF', '-')}",
                        "Valor no WinThor (ERP)": str(item_sel.get('EAN_WINTHOR', '-')),
                        "Regra / Cálculo Executado": "Compara EAN limpo (sem zeros à esquerda e ignorando SEM GTIN) contra cEAN ou cEANTrib",
                        "Resultado": item_sel.get('Status_EAN', '-')
                    },
                    {
                        "Critério Auditado": "2. Quantidade e Embalagem",
                        "Valor no XML (NF-e)": f"{formatar_num_br(item_sel.get('Qtd_NF'), 2)} ({item_sel.get('Unid_XML', '-')})",
                        "Valor no WinThor (ERP)": formatar_num_br(item_sel.get('Qtd_WINTHOR'), 2),
                        "Regra / Cálculo Executado": f"Fator extraído = {formatar_num_br(item_sel.get('_Fator_Embalagem'), 0)} → Qtd Convertida = {formatar_num_br(item_sel.get('_Qtd_Convertida_XML'), 2)}",
                        "Resultado": item_sel.get('Status_Qtd', '-')
                    },
                    {
                        "Critério Auditado": "3. Classificação Fiscal (NCM)",
                        "Valor no XML (NF-e)": str(item_sel.get('NCM_NF', '-')),
                        "Valor no WinThor (ERP)": str(item_sel.get('NCM_WINTHOR', '-')),
                        "Regra / Cálculo Executado": "Comparação exata dos 8 dígitos numéricos (sem pontuação)",
                        "Resultado": item_sel.get('Status_NCM', '-')
                    },
                    {
                        "Critério Auditado": "4. Alíquota de ICMS & ST Guia",
                        "Valor no XML (NF-e)": f"{formatar_num_br(item_sel.get('ICMS_NF'), 2, '%')} ({item_sel.get('Origem_XML', '-')})",
                        "Valor no WinThor (ERP)": f"%ICMS: {formatar_num_br(item_sel.get('ICMS_WINTHOR'), 2, '%')} | ST Guia: {formatar_num_br(item_sel.get('%Alíq.Ext. ST Guia_Whintor'), 2, '%')}",
                        "Regra / Cálculo Executado": f"Aprova se ICMS_NF bater com %ICMS ou com %Alíq.Ext.ST Guia (Tipo ST: {item_sel.get('Tipo_Calculo_ST', '-')})",
                        "Resultado": item_sel.get('Status_ICMS_Aliq', '-')
                    },
                    {
                        "Critério Auditado": "5. Redução de Base de Cálculo",
                        "Valor no XML (NF-e)": f"pRedBC: {formatar_num_br(item_sel.get('pRedBC_NF'), 4, '%')}",
                        "Valor no WinThor (ERP)": f"% ICMS red.: {formatar_num_br(item_sel.get('% ICMS red. Whintor'), 4, '%')}",
                        "Regra / Cálculo Executado": f"Carga Efetiva XML = {formatar_num_br(item_sel.get('_Carga_Efetiva_XML'), 4, '%')} (compara % direto ou Carga Efetiva em 2 e 4 casas)",
                        "Resultado": item_sel.get('Status_Red_BC', '-')
                    }
                ]
                df_mem = pd.DataFrame(linhas_memoria)
                st.table(df_mem)

        st.markdown("---")
        st.success("✨ Auditoria concluída! Você pode inspecionar os cards acima ou baixar a planilha executiva abaixo:")
        st.download_button(
            label="📥 Baixar Relatório Executivo em Excel (.xlsx)",
            data=excel_bytes,
            file_name=nome_arquivo_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_download_bottom"
        )
