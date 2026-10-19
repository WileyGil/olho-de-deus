import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from itertools import zip_longest
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
import re
from datetime import datetime
import io

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
    
    # Tratamento caso venha notação científica do Excel ou ERP
    if 'E+' in val_str or 'E-' in val_str:
        try:
            val_limpo = val_str.replace(',', '.')
            val_str = str(int(float(val_limpo)))
        except:
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
    except:
        return 0.0

def obter_icms_xml(det_node, ns):
    pICMS = det_node.find('.//ns:pICMS', ns)
    if pICMS is not None and pICMS.text:
        return float(pICMS.text)
    return 0.0

# ==========================================
# 2. INTERFACE STREAMLIT
# ==========================================
st.set_page_config(page_title="Auditoria Olho de Deus", layout="wide")
st.title("Auditoria de Notas")
st.markdown("Faça o upload do arquivo XML da nota e do arquivo do WinThor (.csv ou .xlsx) para cruzar as informações com base absoluta na ordem do XML.")

col1, col2 = st.columns(2)
with col1:
    arquivo_xml = st.file_uploader("Upload do XML (Fornecedor)", type=["xml"])
with col2:
    arquivo_dados = st.file_uploader("Upload do Relatório WinThor", type=["csv", "xlsx"])

if arquivo_xml and arquivo_dados:
    if st.button("Processar Auditoria"):
        with st.spinner("Analisando com base fixa no XML..."):
            # ==========================================
            # 3. EXTRAÇÃO DO XML (NF - O CENTRO DE TUDO)
            # ==========================================
            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
            tree = ET.parse(arquivo_xml)
            root = tree.getroot()

            numero_nf = "DESCONHECIDO"
            nNF_node = root.find('.//ns:ide/ns:nNF', ns)
            if nNF_node is not None:
                numero_nf = nNF_node.text

            itens_xml = []
            for det in root.findall('.//ns:det', ns):
                prod = det.find('ns:prod', ns)

                try:
                    nItem = int(det.attrib.get('nItem', '0'))
                except:
                    nItem = det.attrib.get('nItem', '')

                desc = prod.find('ns:xProd', ns).text if prod.find('ns:xProd', ns) is not None else ''
                ean_raw = prod.find('ns:cEAN', ns).text if prod.find('ns:cEAN', ns) is not None else ''
                ean_trib_raw = prod.find('ns:cEANTrib', ns).text if prod.find('ns:cEANTrib', ns) is not None else ''

                ean_exibicao = limpar_numero(ean_raw) if limpar_ean(ean_raw) != '' else limpar_numero(ean_trib_raw)
                ean_comp = limpar_ean(ean_raw)
                ean_trib_comp = limpar_ean(ean_trib_raw)

                ncm = limpar_numero(prod.find('ns:NCM', ns).text) if prod.find('ns:NCM', ns) is not None else ''
                qtd = float(prod.find('ns:qCom', ns).text) if prod.find('ns:qCom', ns) is not None else 0.0
                vUn = float(prod.find('ns:vUnCom', ns).text) if prod.find('ns:vUnCom', ns) is not None else 0.0
                vTot = float(prod.find('ns:vProd', ns).text) if prod.find('ns:vProd', ns) is not None else 0.0
                icms = obter_icms_xml(det, ns)

                pRedBC_node = det.find('.//ns:pRedBC', ns)
                pRedBC = float(pRedBC_node.text) if pRedBC_node is not None and pRedBC_node.text else 0.0

                itens_xml.append({
                    'Linha_NF': nItem,
                    'Descricao_NF': desc,
                    'EAN_NF': ean_exibicao,
                    'EAN_Comp_NF': ean_comp,
                    'EANTrib_Comp_NF': ean_trib_comp,
                    'NCM_NF': ncm,
                    'Qtd_NF': qtd,
                    'Vl_Unit_NF': vUn,
                    'Vl_Total_NF': vTot,
                    'ICMS_NF': icms,
                    'pRedBC_NF': pRedBC
                })

            data_simples = datetime.now().strftime('%Y%m%d')
            nome_arquivo_saida = f'relatorio_NF_{numero_nf}_{data_simples}.xlsx'

            # ==========================================
            # 4. EXTRAÇÃO DA PLANILHA (WINTHOR)
            # ==========================================
            if arquivo_dados.name.lower().endswith('.xlsx'):
                df_csv = pd.read_excel(arquivo_dados, dtype=str)
            else:
                df_csv = pd.read_csv(arquivo_dados, sep=';', encoding='latin1', dtype=str)

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
                    'Descricao_WINTHOR': str(row.get('Descrição', '')).strip(),
                    'Qtd_WINTHOR': converte_float_br(row.get('Qtde Pré-Entrada', 0.0)),
                    'NCM_WINTHOR': limpar_numero(row.get('NCM', '')),
                    'ICMS_WINTHOR': converte_float_br(row.get('%ICMS', 0.0)),
                    '% ICMS red. Whintor': converte_float_br(row.get('% ICMS red.', 0.0)),
                    '%Alíq.Ext. ST Guia_Whintor': converte_float_br(row.get('%Alíq.Ext.ST Guia', 0.0)),
                    'Tipo_Calculo_WINTHOR': str(row.get(col_tipo_calc, '')).strip()
                })

            # ==========================================
            # 5. CRUZAMENTO BASEADO ESTRITAMENTE NA ORDEM DO XML
            # ==========================================
            dados_relatorio = []

            for item_x, item_c in zip_longest(itens_xml, itens_csv, fillvalue=None):
                linha = {}

                if item_x:
                    linha.update(item_x)
                else:
                    linha.update({
                        'Linha_NF': '-', 'Descricao_NF': '-', 'EAN_NF': '-', 'NCM_NF': '-',
                        'Qtd_NF': '-', 'Vl_Unit_NF': '-', 'Vl_Total_NF': '-', 'ICMS_NF': '-', 'pRedBC_NF': '-'
                    })

                if item_c:
                    linha.update(item_c)
                else:
                    linha.update({
                        'Linha_WINTHOR': '-', 'Cod_Produto_WINTHOR': '-', 'EAN_WINTHOR': '-', 'Descricao_WINTHOR': '-',
                        'Qtd_WINTHOR': '-', 'NCM_WINTHOR': '-', 'ICMS_WINTHOR': '-', '% ICMS red. Whintor': '-',
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
                    if pred_xml == red_csv:
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
                    linha['Status_Qtd'] = 'OK' if round(item_x['Qtd_NF'], 4) == round(item_c['Qtd_WINTHOR'], 4
