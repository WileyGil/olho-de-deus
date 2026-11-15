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
        '5': '5 - Nacional (Imp. <= 70%)',
        '6': '6 - Estrangeira (Sem Similar Nacional)',
        '7': '7 - Estrangeira (Mercado Interno, Sem Similar)',
        '8': '8 - Nacional (Imp. > 70%)'
    }
    tabela_origem['5'] = '5 - Nacional (Imp. <= 40%)'
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


# ==========================================
# 2. CONFIGURAÇÃO DE PÁGINA E ESTILO (100% FOCO EM DIVERGÊNCIAS)
# ==========================================
st.set_page_config(
    page_title="Painel de Auditoria de Notas - Foco em Divergências",
    page_icon="🚨",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        .stApp {
            background-color: #F4F6F9;
            color: #1B2631;
        }
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 98%;
        }
        /* Top Header Executivo */
        .top-banner {
            background: linear-gradient(135deg, #1B2631 0%, #2C3E50 100%);
            padding: 14px 22px;
            border-radius: 8px;
            color: #FFFFFF;
            margin-bottom: 14px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 3px 10px rgba(0,0,0,0.08);
        }
        .top-banner-title {
            font-size: 1.35rem;
            font-weight: 700;
            margin: 0;
            color: #FFFFFF;
        }
        .top-banner-sub {
            font-size: 0.82rem;
            color: #BDC3C7;
            margin: 2px 0 0 0;
        }
        /* Barra de Comando da Nota */
        .nf-action-bar {
            background-color: #FFFFFF;
            border: 1px solid #E5E7E9;
            border-left: 5px solid #9C0006;
            border-radius: 8px;
            padding: 12px 18px;
            margin-bottom: 12px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        }
        .nf-title-line {
            font-size: 1.08rem;
            font-weight: 700;
            color: #1B2631;
            margin: 0;
        }
        .nf-meta-line {
            font-size: 0.83rem;
            color: #566573;
            margin-top: 3px;
        }
        /* Caixa de Divergência Vermelha (Destaque Principal no Card) */
        .div-alert-box {
            background-color: #FDEDEC;
            border: 1.5px solid #F5B7B1;
            border-left: 6px solid #9C0006;
            border-radius: 8px;
            padding: 14px 18px;
            margin-bottom: 12px;
        }
        .div-alert-header {
            font-size: 0.82rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #9C0006;
            margin-bottom: 10px;
            border-bottom: 1px solid #F5B7B1;
            padding-bottom: 6px;
        }
        .div-item-row {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            align-items: center;
            background-color: #FFFFFF;
            border: 1px solid #FADBD8;
            border-radius: 6px;
            padding: 9px 14px;
            margin-bottom: 8px;
        }
        .div-item-row:last-child {
            margin-bottom: 0;
        }
        .div-crit-name {
            font-size: 0.9rem;
            font-weight: 700;
            color: #9C0006;
            min-width: 190px;
        }
        .div-comp-vals {
            font-size: 0.9rem;
            font-weight: 600;
            color: #1B2631;
        }
        .val-xml-tag {
            background-color: #EBF5FB;
            color: #1B4F72;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: 700;
            border: 1px solid #AED6F1;
        }
        .val-erp-tag {
            background-color: #FADBD8;
            color: #9C0006;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: 700;
            border: 1px solid #F1948A;
        }
        .div-rule-note {
            font-size: 0.78rem;
            color: #641E16;
            background-color: #F9EBEA;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: 600;
        }
        /* Ficha Técnica do Produto (Neutra, sem selos OK) */
        .prod-info-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            background-color: #F8F9F9;
            border: 1px solid #E5E7E9;
            border-radius: 8px;
            padding: 12px 14px;
        }
        .info-col-title {
            font-size: 0.78rem;
            font-weight: 700;
            text-transform: uppercase;
            color: #566573;
            border-bottom: 1px solid #D5D8DC;
            padding-bottom: 4px;
            margin-bottom: 6px;
        }
        .info-kv {
            display: flex;
            justify-content: space-between;
            font-size: 0.81rem;
            padding: 2px 0;
        }
        .info-k {
            color: #566573;
        }
        .info-v {
            font-weight: 600;
            color: #1B2631;
            text-align: right;
        }
        .desc-compare-bar {
            background-color: #FFFFFF;
            border: 1px solid #E5E7E9;
            border-radius: 6px;
            padding: 8px 14px;
            margin-bottom: 10px;
            font-size: 0.83rem;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 14px;
        }
        /* Ajuste de botões no Streamlit */
        div.stButton > button {
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.82rem;
            padding: 0.35rem 0.75rem;
        }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 3. MOTOR DE AUDITORIA E GERAÇÃO DO EXCEL
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
    ocorrencias_de_para = []

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
        detalhes_erros_item = []

        if item_x and item_c:
            linha['Status_Geral'] = 'PAREADO'
            cod_prod = linha.get('Cod_Produto_WINTHOR') or '-'
            desc_ref = linha.get('Descricao_WINTHOR') or linha.get('Descricao_NF') or '-'

            # 1. Validação EAN
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
                val_xml_ean = linha['EAN_NF'] if linha['EAN_NF'] else 'SEM GTIN'
                val_win_ean = linha['EAN_WINTHOR'] if linha['EAN_WINTHOR'] else 'VAZIO'
                detalhes_erros_item.append({
                    'criterio': 'DIVERGÊNCIA DE EAN (CÓDIGO DE BARRAS)',
                    'tag': 'EAN',
                    'xml': val_xml_ean,
                    'winthor': val_win_ean,
                    'regra': f"cEAN ({item_x.get('EAN_Raw_NF') or '-'}) / cEANTrib ({item_x.get('EANTrib_Raw_NF') or '-'}) ≠ ERP"
                })
                ocorrencias_de_para.append({
                    'Linha NF': linha['Linha_NF'],
                    'Linha ERP': linha['Linha_WINTHOR'],
                    'Cód. WinThor': cod_prod,
                    'Produto': desc_ref,
                    'Critério Divergente': 'EAN (Cód. Barras)',
                    'Valor no XML (NF)': val_xml_ean,
                    'Valor no WinThor (ERP)': val_win_ean,
                    'Diagnóstico': 'Código de barras da NF não confere com o cadastro no WinThor'
                })

            # 2. Tipo Cálculo ST (Tradutor)
            tipo_calc = str(item_c.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
            if tipo_calc == 'N':
                linha['Tipo_Calculo_ST'] = 'NOTA FISCAL'
            elif tipo_calc == 'G':
                linha['Tipo_Calculo_ST'] = 'GUIA'
            else:
                linha['Tipo_Calculo_ST'] = item_c['Tipo_Calculo_WINTHOR']

            # 3. Validação ICMS Alíquota
            icms_xml = round(item_x['ICMS_NF'], 2)
            icms_csv = round(item_c['ICMS_WINTHOR'], 2)
            st_guia = round(item_c['%Alíq.Ext. ST Guia_Whintor'], 2)
            if icms_xml == icms_csv or icms_xml == st_guia:
                linha['Status_ICMS_Aliq'] = 'OK'
            else:
                linha['Status_ICMS_Aliq'] = 'DIVERGENTE'
                lista_erros_item.append('ICMS')
                val_xml_icms = formatar_num_br(icms_xml, 2, '%')
                val_win_icms = f"%ICMS: {formatar_num_br(icms_csv, 2, '%')} | ST Guia: {formatar_num_br(st_guia, 2, '%')}"
                detalhes_erros_item.append({
                    'criterio': 'DIVERGÊNCIA DE ALÍQUOTA DE ICMS',
                    'tag': 'ICMS',
                    'xml': val_xml_icms,
                    'winthor': val_win_icms,
                    'regra': f"Origem: {linha['Origem_XML']} | Tipo ST: {linha['Tipo_Calculo_ST']}"
                })
                ocorrencias_de_para.append({
                    'Linha NF': linha['Linha_NF'],
                    'Linha ERP': linha['Linha_WINTHOR'],
                    'Cód. WinThor': cod_prod,
                    'Produto': desc_ref,
                    'Critério Divergente': 'Alíquota ICMS',
                    'Valor no XML (NF)': val_xml_icms,
                    'Valor no WinThor (ERP)': val_win_icms,
                    'Diagnóstico': f"ICMS da NF diferente de %ICMS e de %Alíq.Ext.ST Guia (ST: {linha['Tipo_Calculo_ST']})"
                })

            # 4. Validação Redução de Base de Cálculo
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
                lista_erros_item.append('Redução BC')
                val_xml_red = f"pRedBC: {formatar_num_br(pred_xml, 4, '%')} (Carga Efetiva: {formatar_num_br(carga_efetiva_xml, 2, '%')})"
                val_win_red = formatar_num_br(red_csv, 4, '%')
                detalhes_erros_item.append({
                    'criterio': 'DIVERGÊNCIA DE REDUÇÃO DE BASE DE CÁLCULO',
                    'tag': 'Redução BC',
                    'xml': val_xml_red,
                    'winthor': val_win_red,
                    'regra': f"Carga Efetiva XML ({formatar_num_br(carga_efetiva_xml, 4, '%')}) ≠ WinThor ({val_win_red})"
                })
                ocorrencias_de_para.append({
                    'Linha NF': linha['Linha_NF'],
                    'Linha ERP': linha['Linha_WINTHOR'],
                    'Cód. WinThor': cod_prod,
                    'Produto': desc_ref,
                    'Critério Divergente': 'Redução de Base',
                    'Valor no XML (NF)': val_xml_red,
                    'Valor no WinThor (ERP)': val_win_red,
                    'Diagnóstico': 'Nem o % de redução nem a carga efetiva coincidem com o WinThor'
                })

            # 5. Validação NCM
            linha['Status_NCM'] = 'OK' if str(item_x['NCM_NF']) == str(item_c['NCM_WINTHOR']) else 'DIVERGENTE'
            if linha['Status_NCM'] == 'DIVERGENTE':
                lista_erros_item.append('NCM')
                val_xml_ncm = str(item_x['NCM_NF'] or '-')
                val_win_ncm = str(item_c['NCM_WINTHOR'] or '-')
                detalhes_erros_item.append({
                    'criterio': 'DIVERGÊNCIA DE CLASSIFICAÇÃO FISCAL (NCM)',
                    'tag': 'NCM',
                    'xml': val_xml_ncm,
                    'winthor': val_win_ncm,
                    'regra': 'Código NCM de 8 dígitos divergente entre NF e ERP'
                })
                ocorrencias_de_para.append({
                    'Linha NF': linha['Linha_NF'],
                    'Linha ERP': linha['Linha_WINTHOR'],
                    'Cód. WinThor': cod_prod,
                    'Produto': desc_ref,
                    'Critério Divergente': 'NCM (Fiscal)',
                    'Valor no XML (NF)': val_xml_ncm,
                    'Valor no WinThor (ERP)': val_win_ncm,
                    'Diagnóstico': 'Classificação fiscal NCM divergente'
                })

            # 6. Validação Quantidades
            qtd_xml_pura = item_x['Qtd_NF']
            qtd_win = item_c['Qtd_WINTHOR']
            fator = item_x['_Fator_Embalagem']
            qtd_xml_convertida = qtd_xml_pura * fator

            if round(qtd_xml_pura, 4) == round(qtd_win, 4) or round(qtd_xml_convertida, 4) == round(qtd_win, 4):
                linha['Status_Qtd'] = 'OK'
            else:
                linha['Status_Qtd'] = 'DIVERGENTE'
                lista_erros_item.append('Quantidade')
                val_xml_qtd = (
                    f"{formatar_num_br(qtd_xml_pura, 2)} {linha['Unid_XML']} "
                    f"(×{formatar_num_br(fator, 0)} = {formatar_num_br(qtd_xml_convertida, 2)})"
                    if fator != 1.0 else f"{formatar_num_br(qtd_xml_pura, 2)} {linha['Unid_XML']}"
                )
                val_win_qtd = formatar_num_br(qtd_win, 2)
                detalhes_erros_item.append({
                    'criterio': 'DIVERGÊNCIA DE QUANTIDADE',
                    'tag': 'Quantidade',
                    'xml': val_xml_qtd,
                    'winthor': val_win_qtd,
                    'regra': f"Qtd Pura ({formatar_num_br(qtd_xml_pura, 2)}) e Convertida ({formatar_num_br(qtd_xml_convertida, 2)}) ≠ ERP ({val_win_qtd})"
                })
                ocorrencias_de_para.append({
                    'Linha NF': linha['Linha_NF'],
                    'Linha ERP': linha['Linha_WINTHOR'],
                    'Cód. WinThor': cod_prod,
                    'Produto': desc_ref,
                    'Critério Divergente': 'Quantidade',
                    'Valor no XML (NF)': val_xml_qtd,
                    'Valor no WinThor (ERP)': val_win_qtd,
                    'Diagnóstico': 'Quantidade da NF (pura ou convertida pelo fator da embalagem) não bate com a Pré-Entrada'
                })

        elif item_x and not item_c:
            linha['Status_Geral'] = 'SOBRA'
            linha['_Tipo_Sobra'] = 'SOBRA NO XML (Ausente no WinThor)'
            linha['_Carga_Efetiva_XML'] = '-'
            for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                linha[st_col] = '-'
            lista_erros_item.append('Sobras')
            detalhes_erros_item.append({
                'criterio': 'ITEM SOBRANDO NO XML (SEM CORRESPONDENTE NO WINTHOR)',
                'tag': 'Sobras',
                'xml': f"Linha NF #{linha['Linha_NF']} — {linha['Descricao_NF']}",
                'winthor': 'AUSENTE NO ERP',
                'regra': 'Faturado na Nota Fiscal, mas sem linha correspondente na planilha do WinThor'
            })
            ocorrencias_de_para.append({
                'Linha NF': linha['Linha_NF'],
                'Linha ERP': '-',
                'Cód. WinThor': '-',
                'Produto': linha['Descricao_NF'],
                'Critério Divergente': 'Sobra no XML',
                'Valor no XML (NF)': f"Qtd: {formatar_num_br(linha['Qtd_NF'], 2)}",
                'Valor no WinThor (ERP)': 'AUSENTE',
                'Diagnóstico': 'Item presente no XML, mas sem par na planilha WinThor'
            })

        else:
            linha['Status_Geral'] = 'SOBRA'
            linha['_Tipo_Sobra'] = 'SOBRA NO WINTHOR (Ausente no XML)'
            linha['_Carga_Efetiva_XML'] = '-'
            for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                linha[st_col] = '-'
            lista_erros_item.append('Sobras')
            detalhes_erros_item.append({
                'criterio': 'ITEM SOBRANDO NO WINTHOR (SEM CORRESPONDENTE NO XML)',
                'tag': 'Sobras',
                'xml': 'AUSENTE NO XML',
                'winthor': f"Linha ERP #{linha['Linha_WINTHOR']} (Cód. {linha['Cod_Produto_WINTHOR']}) — {linha['Descricao_WINTHOR']}",
                'regra': 'Presente na Pré-Entrada do WinThor, mas sem item correspondente no XML'
            })
            ocorrencias_de_para.append({
                'Linha NF': '-',
                'Linha ERP': linha['Linha_WINTHOR'],
                'Cód. WinThor': linha['Cod_Produto_WINTHOR'],
                'Produto': linha['Descricao_WINTHOR'],
                'Critério Divergente': 'Sobra no WinThor',
                'Valor no XML (NF)': 'AUSENTE',
                'Valor no WinThor (ERP)': f"Qtd: {formatar_num_br(linha['Qtd_WINTHOR'], 2)}",
                'Diagnóstico': 'Item esperado no WinThor, mas ausente no XML'
            })

        linha['_Erros_Item'] = lista_erros_item
        linha['_Detalhes_Erros'] = detalhes_erros_item
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
        'dados_relatorio': dados_relatorio,
        'ocorrencias_de_para': ocorrencias_de_para
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


def renderizar_card_divergente(item):
    detalhes_erros = item.get('_Detalhes_Erros', [])

    # 1. Bloco Principal em Destaque Vermelho: SOMENTE AS DIVERGÊNCIAS ENCONTRADAS
    linhas_html_erros = []
    for err in detalhes_erros:
        crit_esc = html.escape(str(err.get('criterio', '')))
        xml_esc = html.escape(str(err.get('xml', '-')))
        win_esc = html.escape(str(err.get('winthor', '-')))
        regra_esc = html.escape(str(err.get('regra', '')))

        linhas_html_erros.append(f"""
            <div class="div-item-row">
                <div class="div-crit-name">🔴 {crit_esc}</div>
                <div class="div-comp-vals">
                    XML: <span class="val-xml-tag">{xml_esc}</span>
                    &nbsp;&nbsp;➔&nbsp;&nbsp;
                    WINTHOR: <span class="val-erp-tag">{win_esc}</span>
                </div>
                <div class="div-rule-note">{regra_esc}</div>
            </div>
        """)

    st.markdown(f"""
        <div class="div-alert-box">
            <div class="div-alert-header">🚨 Ocorrências que exigem correção neste produto ({len(detalhes_erros)})</div>
            {"".join(linhas_html_erros)}
        </div>
    """, unsafe_allow_html=True)

    # 2. Ficha Completa do Produto para Análise (Neutra, sem selos OK)
    desc_nf = html.escape(str(item.get('Descricao_NF', '-')))
    desc_win = html.escape(str(item.get('Descricao_WINTHOR', '-')))
    fator_emb = item.get('_Fator_Embalagem', 1.0)
    unid_xml = html.escape(str(item.get('Unid_XML', '-')))

    st.markdown(f"""
        <div class="desc-compare-bar">
            <div><strong>📄 Descrição no XML:</strong> {desc_nf}</div>
            <div><strong>🏢 Descrição no WinThor:</strong> {desc_win}</div>
        </div>
        <div class="prod-info-grid">
            <div>
                <div class="info-col-title">🏷️ Dados de Identificação</div>
                <div class="info-kv"><span class="info-k">Linha NF / ERP:</span><span class="info-v">#{item.get('Linha_NF', '-')} / #{item.get('Linha_WINTHOR', '-')}</span></div>
                <div class="info-kv"><span class="info-k">Cód. WinThor:</span><span class="info-v">{html.escape(str(item.get('Cod_Produto_WINTHOR', '-')))}</span></div>
                <div class="info-kv"><span class="info-k">EAN XML:</span><span class="info-v">{html.escape(str(item.get('EAN_NF') or 'SEM GTIN'))}</span></div>
                <div class="info-kv"><span class="info-k">EAN WinThor:</span><span class="info-v">{html.escape(str(item.get('EAN_WINTHOR') or 'VAZIO'))}</span></div>
            </div>
            <div>
                <div class="info-col-title">📦 Quantidade & Valores</div>
                <div class="info-kv"><span class="info-k">Unid. / Fator XML:</span><span class="info-v">{unid_xml} (×{formatar_num_br(fator_emb, 0)})</span></div>
                <div class="info-kv"><span class="info-k">Qtd XML (Pura / Conv.):</span><span class="info-v">{formatar_num_br(item.get('Qtd_NF'), 2)} / {formatar_num_br(item.get('_Qtd_Convertida_XML'), 2)}</span></div>
                <div class="info-kv"><span class="info-k">Qtd WinThor:</span><span class="info-v">{formatar_num_br(item.get('Qtd_WINTHOR'), 2)}</span></div>
                <div class="info-kv"><span class="info-k">Vl. Unit. / Total NF:</span><span class="info-v">{formatar_moeda_br(item.get('Vl_Unit_NF'))} | {formatar_moeda_br(item.get('Vl_Total_NF'))}</span></div>
            </div>
            <div>
                <div class="info-col-title">💰 Tributação ICMS & ST</div>
                <div class="info-kv"><span class="info-k">Origem XML:</span><span class="info-v">{html.escape(str(item.get('Origem_XML', '-')))}</span></div>
                <div class="info-kv"><span class="info-k">ICMS XML:</span><span class="info-v">{formatar_num_br(item.get('ICMS_NF'), 2, '%')}</span></div>
                <div class="info-kv"><span class="info-k">%ICMS WinThor:</span><span class="info-v">{formatar_num_br(item.get('ICMS_WINTHOR'), 2, '%')}</span></div>
                <div class="info-kv"><span class="info-k">ST Guia / Tipo ST:</span><span class="info-v">{formatar_num_br(item.get('%Alíq.Ext. ST Guia_Whintor'), 2, '%')} ({html.escape(str(item.get('Tipo_Calculo_ST', '-')))})</span></div>
            </div>
            <div>
                <div class="info-col-title">📉 Redução & Fiscal</div>
                <div class="info-kv"><span class="info-k">NCM (XML | ERP):</span><span class="info-v">{html.escape(str(item.get('NCM_NF', '-')))} | {html.escape(str(item.get('NCM_WINTHOR', '-')))}</span></div>
                <div class="info-kv"><span class="info-k">pRedBC XML:</span><span class="info-v">{formatar_num_br(item.get('pRedBC_NF'), 4, '%')}</span></div>
                <div class="info-kv"><span class="info-k">Carga Efetiva XML:</span><span class="info-v">{formatar_num_br(item.get('_Carga_Efetiva_XML'), 4, '%')}</span></div>
                <div class="info-kv"><span class="info-k">% ICMS Red. ERP:</span><span class="info-v">{formatar_num_br(item.get('% ICMS red. Whintor'), 4, '%')}</span></div>
            </div>
        </div>
    """, unsafe_allow_html=True)


# ==========================================
# 4. BARRA LATERAL (UPLOAD E COLUNAS DO EXCEL)
# ==========================================
with st.sidebar:
    st.markdown("### 📂 Arquivos da Auditoria")
    arquivo_xml = st.file_uploader("1. XML Nota Fiscal", type=["xml"])
    arquivo_dados = st.file_uploader("2. WinThor (.csv ou .xlsx)", type=["csv", "xlsx"])

    st.markdown("---")
    st.markdown("### ⚙️ Colunas do Relatório Excel")
    selecionar_todos = st.checkbox("Selecionar Todos os Blocos", value=True)

    inc_ident = st.checkbox("Identificação (Linhas)", value=selecionar_todos)
    inc_ean = st.checkbox("Código & EAN", value=selecionar_todos)
    inc_desc = st.checkbox("Descrições", value=selecionar_todos)
    inc_qtd = st.checkbox("Quantidades", value=selecionar_todos)
    inc_ncm = st.checkbox("NCM (Fiscal)", value=selecionar_todos)
    inc_icms = st.checkbox("ICMS & Origem", value=selecionar_todos)
    inc_red = st.checkbox("Redução de Base", value=selecionar_todos)
    inc_st = st.checkbox("Tipo Cálculo ST", value=selecionar_todos)


# ==========================================
# 5. PAINEL PRINCIPAL — 100% FOCADO EM DIVERGÊNCIAS
# ==========================================
st.markdown("""
    <div class="top-banner">
        <div>
            <p class="top-banner-title">🚨 Painel de Auditoria de Notas Fiscais — Foco em Divergências</p>
            <p class="top-banner-sub">Exibindo exclusivamente inconsistências fiscais, cadastrais e sobras entre o XML do fornecedor e o WinThor.</p>
        </div>
    </div>
""", unsafe_allow_html=True)

if not (arquivo_xml and arquivo_dados):
    st.info("👈 **Envie o arquivo XML da Nota Fiscal e a planilha do WinThor (.csv ou .xlsx) no menu lateral esquerdo** para carregar automaticamente todas as divergências.")
else:
    assinatura_arquivos = f"{arquivo_xml.name}_{arquivo_xml.size}_{arquivo_dados.name}_{arquivo_dados.size}"

    if 'resultado_auditoria' not in st.session_state or st.session_state.get('assinatura_arquivos') != assinatura_arquivos:
        with st.spinner("Cruzando itens e filtrando divergências..."):
            st.session_state['resultado_auditoria'] = executar_auditoria(
                arquivo_xml.getvalue(),
                arquivo_dados.getvalue(),
                arquivo_dados.name
            )
            st.session_state['assinatura_arquivos'] = assinatura_arquivos
            st.session_state['filtro_erro_ativo'] = 'TODOS'

    res = st.session_state['resultado_auditoria']
    dados_relatorio = res['dados_relatorio']
    ocorrencias_de_para = res['ocorrencias_de_para']
    numero_nf = res['numero_nf']
    emit_nome = res['emit_nome']
    nome_arquivo_saida = res['nome_arquivo_saida']

    # Preparação do DataFrame para o Excel (mantém a funcionalidade original completa)
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
    excel_bytes = gerar_excel_formatado(df_final, dados_relatorio, super_grupos_def)

    # Filtra EXCLUSIVAMENTE itens com divergência ou sobra (tudo que está 100% OK é descartado da tela)
    itens_divergentes = [d for d in dados_relatorio if d['_Qtd_Erros'] > 0]
    # Ordena os produtos com mais divergências primeiro
    itens_divergentes.sort(key=lambda x: ('SOBRA' in str(x['Status_Geral']), -x['_Qtd_Erros']))

    div_ean = sum(1 for d in itens_divergentes if 'EAN' in d['_Erros_Item'])
    div_qtd = sum(1 for d in itens_divergentes if 'Quantidade' in d['_Erros_Item'])
    div_ncm = sum(1 for d in itens_divergentes if 'NCM' in d['_Erros_Item'])
    div_icms = sum(1 for d in itens_divergentes if 'ICMS' in d['_Erros_Item'])
    div_red = sum(1 for d in itens_divergentes if 'Redução BC' in d['_Erros_Item'])
    div_sobras = sum(1 for d in itens_divergentes if 'Sobras' in d['_Erros_Item'])
    total_ocorrencias = len(ocorrencias_de_para)

    # ==========================================
    # BARRA DE COMANDO DO TOPO (NF + BUSCA + DOWNLOAD EXCEL + REPROCESSAR)
    # ==========================================
    bar_col1, bar_col2, bar_col3, bar_col4 = st.columns([2.6, 1.8, 1.4, 0.9], gap="small")

    with bar_col1:
        emit_txt = f" • {html.escape(emit_nome)}" if emit_nome else ""
        st.markdown(f"""
            <div class="nf-action-bar">
                <p class="nf-title-line">NF Nº {html.escape(str(numero_nf))} — {len(itens_divergentes)} Produto(s) com Divergência ({total_ocorrencias} erros)</p>
                <p class="nf-meta-line">Total auditado: {res['qtd_xml']} itens XML vs {res['qtd_csv']} itens WinThor{emit_txt}</p>
            </div>
        """, unsafe_allow_html=True)

    with bar_col2:
        termo_busca = st.text_input(
            "Buscar nas divergências:",
            placeholder="🔍 Buscar Cód. WinThor, EAN, NCM ou Produto...",
            label_visibility="collapsed"
        ).strip().lower()

    with bar_col3:
        st.download_button(
            label="📥 Baixar Excel Executivo",
            data=excel_bytes,
            file_name=nome_arquivo_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

    with bar_col4:
        if st.button("🔄 Reprocessar", use_container_width=True):
            st.session_state.pop('resultado_auditoria', None)
            st.rerun()

    # ==========================================
    # BOTÕES DE FILTRO RÁPIDO DE DIVERGÊNCIAS (SEM BOTÃO DE ITENS OK)
    # ==========================================
    if 'filtro_erro_ativo' not in st.session_state:
        st.session_state['filtro_erro_ativo'] = 'TODOS'

    f_ativo = st.session_state['filtro_erro_ativo']

    btn_cols = st.columns(7, gap="small")
    botoes_config = [
        ('TODOS', f"🔴 Todos os Erros ({len(itens_divergentes)})", len(itens_divergentes)),
        ('EAN', f"🏷️ EAN ({div_ean})", div_ean),
        ('Quantidade', f"📦 Quantidade ({div_qtd})", div_qtd),
        ('NCM', f"🏛️ NCM ({div_ncm})", div_ncm),
        ('ICMS', f"💰 ICMS ({div_icms})", div_icms),
        ('Redução BC', f"📉 Redução BC ({div_red})", div_red),
        ('Sobras', f"⚠️ Sobras ({div_sobras})", div_sobras),
    ]

    for idx_b, (chave_f, rotulo_f, contagem_f) in enumerate(botoes_config):
        with btn_cols[idx_b]:
            tipo_btn = "primary" if f_ativo == chave_f else "secondary"
            desabilitado = (contagem_f == 0 and chave_f != 'TODOS')
            if st.button(rotulo_f, key=f"btn_flt_{chave_f}", use_container_width=True, type=tipo_btn, disabled=desabilitado):
                st.session_state['filtro_erro_ativo'] = chave_f
                st.rerun()

    # Aplicando filtro selecionado e busca textual sobre os itens divergentes
    itens_filtrados = []
    for item in itens_divergentes:
        if f_ativo != 'TODOS' and f_ativo not in item['_Erros_Item']:
            continue

        if termo_busca:
            blob_busca = " ".join([
                str(item.get('Linha_NF', '')),
                str(item.get('Cod_Produto_WINTHOR', '')),
                str(item.get('EAN_NF', '')),
                str(item.get('EAN_WINTHOR', '')),
                str(item.get('NCM_NF', '')),
                str(item.get('NCM_WINTHOR', '')),
                str(item.get('Descricao_NF', '')),
                str(item.get('Descricao_WINTHOR', ''))
            ]).lower()
            if termo_busca not in blob_busca:
                continue

        itens_filtrados.append(item)

    if len(itens_divergentes) == 0:
        st.warning("Nenhuma divergência encontrada entre o XML e a planilha do WinThor.")
    elif len(itens_filtrados) == 0:
        st.info("Nenhuma divergência encontrada para o filtro ou termo pesquisado.")
    else:
        tab_cards, tab_de_para = st.tabs([
            f"🚨 Cards de Produtos com Divergência ({len(itens_filtrados)})",
            f"⚡ Tabela De ➔ Para de Correção ({total_ocorrencias} ocorrências)"
        ])

        # ------------------------------------------
        # ABA 1: CARDS DE PRODUTOS COM DIVERGÊNCIA
        # ------------------------------------------
        with tab_cards:
            ctrl_col1, ctrl_col2 = st.columns([4, 1])
            with ctrl_col1:
                st.caption(
                    f"Exibindo **{len(itens_filtrados)}** produto(s) com divergência "
                    f"(filtro ativo: **{f_ativo}**). Os itens com mais erros aparecem no topo:"
                )
            with ctrl_col2:
                expandir_todos = st.checkbox("Expandir todos os cards", value=(len(itens_filtrados) <= 5))

            for idx_c, item in enumerate(itens_filtrados):
                erros = item['_Erros_Item']
                eh_sobra = 'SOBRA' in str(item['Status_Geral'])
                linha_nf = item.get('Linha_NF', '-')
                linha_win = item.get('Linha_WINTHOR', '-')
                cod_win = item.get('Cod_Produto_WINTHOR', '-')
                nome_prod = item.get('Descricao_WINTHOR') if item.get('Descricao_WINTHOR') not in ('', '-') else item.get('Descricao_NF', '-')

                if eh_sobra:
                    titulo_card = f"⚠️ {item.get('_Tipo_Sobra', 'SOBRA')} | Linha NF #{linha_nf} / ERP #{linha_win} — {nome_prod}"
                else:
                    plural = "Divergências" if len(erros) > 1 else "Divergência"
                    titulo_card = (
                        f"🔴 Linha NF #{linha_nf} | Cód. WinThor: {cod_win} — {nome_prod}   "
                        f"•   [{len(erros)} {plural}: {', '.join(erros)}]"
                    )

                abrir_card = expandir_todos or (idx_c == 0)
                with st.expander(titulo_card, expanded=abrir_card):
                    renderizar_card_divergente(item)

        # ------------------------------------------
        # ABA 2: TABELA DIRETA DE -> PARA (APENAS ERROS)
        # ------------------------------------------
        with tab_de_para:
            ocorrencias_filtradas = []
            for oc in ocorrencias_de_para:
                if f_ativo != 'TODOS':
                    mapa_crit = {
                        'EAN': 'EAN',
                        'Quantidade': 'Quantidade',
                        'NCM': 'NCM',
                        'ICMS': 'ICMS',
                        'Redução BC': 'Redução',
                        'Sobras': 'Sobra'
                    }
                    termo_crit = mapa_crit.get(f_ativo, f_ativo)
                    if termo_crit.lower() not in oc['Critério Divergente'].lower():
                        continue

                if termo_busca:
                    blob_oc = f"{oc['Linha NF']} {oc['Cód. WinThor']} {oc['Produto']} {oc['Valor no XML (NF)']} {oc['Valor no WinThor (ERP)']}".lower()
                    if termo_busca not in blob_oc:
                        continue

                ocorrencias_filtradas.append(oc)

            df_de_para = pd.DataFrame(ocorrencias_filtradas)

            def estilizar_de_para(df_dp):
                def _style_cols(col):
                    if col.name == 'Critério Divergente':
                        return ['background-color: #FADBD8; color: #9C0006; font-weight: 700;' for _ in col]
                    elif col.name == 'Valor no XML (NF)':
                        return ['background-color: #EBF5FB; color: #1B4F72; font-weight: 700;' for _ in col]
                    elif col.name == 'Valor no WinThor (ERP)':
                        return ['background-color: #FDEDEC; color: #9C0006; font-weight: 700;' for _ in col]
                    return ['' for _ in col]
                return df_dp.style.apply(_style_cols, axis=0)

            if df_de_para.empty:
                st.info("Nenhuma ocorrência para o filtro selecionado.")
            else:
                st.dataframe(
                    estilizar_de_para(df_de_para),
                    use_container_width=True,
                    hide_index=True,
                    height=min(120 + len(df_de_para) * 38, 620)
                )
