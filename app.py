import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from itertools import zip_longest
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
from datetime import datetime
import io

# ==========================================
# 1. FUNÇÕES AUXILIARES
# ==========================================
def limpar_numero(valor):
    if pd.isna(valor): return ""
    return re.sub(r'[^0-9]', '', str(valor))

def converte_float_br(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return 0.0
    try:
        return float(str(valor).replace(',', '.'))
    except:
        return 0.0

def obter_icms_xml(det_node, ns):
    pICMS = det_node.find('.//ns:pICMS', ns)
    if pICMS is not None:
        return float(pICMS.text)
    return 0.0

# ==========================================
# 2. INTERFACE STREAMLIT
# ==========================================
st.set_page_config(page_title="Auditoria Olho de Deus", layout="wide")
st.title("PROJETO OLHO DE DEUS - Auditoria de Compras")
st.markdown("Faça o upload do arquivo XML da nota e do arquivo do WinThor (.csv ou .xlsx) para cruzar as informações.")

col1, col2 = st.columns(2)
with col1:
    arquivo_xml = st.file_uploader("Upload do XML (Fornecedor)", type=["xml"])
with col2:
    arquivo_dados = st.file_uploader("Upload do Relatório WinThor", type=["csv", "xlsx"])

if arquivo_xml is not None and arquivo_dados is not None:
    if st.button("Processar Auditoria"):
        with st.spinner("Analisando e cruzando dados..."):
            # ==========================================
            # 3. EXTRAÇÃO DO XML (NF)
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
                ean = prod.find('ns:cEAN', ns).text if prod.find('ns:cEAN', ns) is not None else ''
                ncm = prod.find('ns:NCM', ns).text if prod.find('ns:NCM', ns) is not None else ''
                qtd = float(prod.find('ns:qCom', ns).text) if prod.find('ns:qCom', ns) is not None else 0.0
                vUn = float(prod.find('ns:vUnCom', ns).text) if prod.find('ns:vUnCom', ns) is not None else 0.0
                vTot = float(prod.find('ns:vProd', ns).text) if prod.find('ns:vProd', ns) is not None else 0.0
                icms = obter_icms_xml(det, ns)

                pRedBC_node = det.find('.//ns:pRedBC', ns)
                pRedBC = float(pRedBC_node.text) if pRedBC_node is not None else 0.0

                itens_xml.append({
                    'Linha_NF': nItem, 'Descricao_NF': desc, 'EAN_NF': ean,
                    'NCM_NF': ncm, 'Qtd_NF': qtd, 'Vl_Unit_NF': vUn,
                    'Vl_Total_NF': vTot, 'ICMS_NF': icms, 'pRedBC_NF': pRedBC
                })

            data_simples = datetime.now().strftime('%Y%m%d')
            nome_arquivo_saida = f'relatorio_NF_{numero_nf}_{data_simples}.xlsx'

            # ==========================================
            # 4. EXTRAÇÃO DA PLANILHA (WINTHOR)
            # ==========================================
            if arquivo_dados.name.lower().endswith('.xlsx'):
                df_csv = pd.read_excel(arquivo_dados)
            else:
                df_csv = pd.read_csv(arquivo_dados, sep=';', encoding='latin1')

            df_csv.columns = df_csv.columns.str.strip()
            col_tipo_calc = next((c for c in df_csv.columns if 'tipo' in c.lower() and ('cálculo' in c.lower() or 'calculo' in c.lower())), None)
            if not col_tipo_calc:
                col_tipo_calc = df_csv.columns[-2]

            itens_csv = []
            for index, row in df_csv.iterrows():
                itens_csv.append({
                    'Linha_WINTHOR': index + 1,
                    'Cod_Produto_WINTHOR': str(row.get('Cód. produto', '')),
                    'EAN_WINTHOR': str(row.get('Cód. barra', '')),
                    'Descricao_WINTHOR': str(row.get('Descrição', '')),
                    'Qtd_WINTHOR': converte_float_br(row.get('Qtde Pré-Entrada', 0.0)),
                    'NCM_WINTHOR': limpar_numero(row.get('NCM', '')),
                    'ICMS_WINTHOR': converte_float_br(row.get('%ICMS', 0.0)),
                    '% ICMS red. Whintor': converte_float_br(row.get('% ICMS red.', 0.0)),
                    '%Alíq.Ext. ST Guia_Whintor': converte_float_br(row.get('%Alíq.Ext.ST Guia', 0.0)),
                    'Tipo_Calculo_WINTHOR': str(row.get(col_tipo_calc, '')).strip()
                })

            # ==========================================
            # 5. CRUZAMENTO E LÓGICA DE AUDITORIA
            # ==========================================
            dados_relatorio = []
            alertas_resumo = []
            contagem_erros_produto = {}
            total_pareados = 0
            total_sobras_nf = 0
            total_sobras_winthor = 0
            divergencias_ean = 0
            divergencias_icms = 0
            divergencias_red = 0
            divergencias_ncm = 0
            divergencias_qtd = 0

            for item_x, item_c in zip_longest(itens_xml, itens_csv, fillvalue=None):
                linha = {}
                if item_x:
                    linha.update(item_x)
                else:
                    linha.update({'Linha_NF': '-', 'Descricao_NF': '-', 'EAN_NF': '-', 'NCM_NF': '-', 'Qtd_NF': '-', 'Vl_Unit_NF': '-', 'Vl_Total_NF': '-', 'ICMS_NF': '-', 'pRedBC_NF': '-'})

                if item_c:
                    linha.update(item_c)
                else:
                    linha.update({'Linha_WINTHOR': '-', 'Cod_Produto_WINTHOR': '-', 'EAN_WINTHOR': '-', 'Descricao_WINTHOR': '-', 'Qtd_WINTHOR': '-', 'NCM_WINTHOR': '-', 'ICMS_WINTHOR': '-', '% ICMS red. Whintor': '-', '%Alíq.Ext. ST Guia_Whintor': '-', 'Tipo_Calculo_WINTHOR': '-'})

                if item_x and item_c:
                    linha['Status_Geral'] = 'PAREDO'
                    total_pareados += 1
                    linha['Status_EAN'] = 'OK' if str(item_x['EAN_NF']) == str(item_c['EAN_WINTHOR']) else 'DIVERGENTE'
                    if linha['Status_EAN'] == 'DIVERGENTE': divergencias_ean += 1

                    icms_xml = item_x['ICMS_NF']
                    icms_csv = item_c['ICMS_WINTHOR']
                    st_guia = item_c['%Alíq.Ext. ST Guia_Whintor']

                    if icms_xml == icms_csv or icms_xml == st_guia:
                        linha['Status_ICMS_Aliq'] = 'OK'
                    else:
                        linha['Status_ICMS_Aliq'] = 'DIVERGENTE'
                        divergencias_icms += 1

                    pred_xml = item_x['pRedBC_NF']
                    red_csv = item_c['% ICMS red. Whintor']
                    if pred_xml == red_csv:
                        linha['Status_Red_BC'] = 'OK'
                    else:
                        linha['Status_Red_BC'] = 'DIVERGENTE'
                        divergencias_red += 1

                    tipo_calc = str(item_c.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
                    if tipo_calc == 'N': linha['Tipo_Calculo_ST'] = 'NOTA FISCAL'
                    elif tipo_calc == 'G': linha['Tipo_Calculo_ST'] = 'GUIA'
                    else: linha['Tipo_Calculo_ST'] = item_c['Tipo_Calculo_WINTHOR']

                    linha['Status_NCM'] = 'OK' if str(item_x['NCM_NF']) == str(item_c['NCM_WINTHOR']) else 'DIVERGENTE'
                    if linha['Status_NCM'] == 'DIVERGENTE': divergencias_ncm += 1

                    linha['Status_Qtd'] = 'OK' if item_x['Qtd_NF'] == item_c['Qtd_WINTHOR'] else 'DIVERGENTE'
                    if linha['Status_Qtd'] == 'DIVERGENTE': divergencias_qtd += 1

                    cod_prod = item_c['Cod_Produto_WINTHOR']
                    alertas_atuais = []
                    if linha['Status_EAN'] == 'DIVERGENTE': alertas_atuais.append({"Tipo": "EAN", "Cod. ERP": cod_prod, "Descrição": item_c['Descricao_WINTHOR'], "Detalhe": f"XML: {item_x['EAN_NF']} vs ERP: {item_c['EAN_WINTHOR']}"})
                    if linha['Status_ICMS_Aliq'] == 'DIVERGENTE': alertas_atuais.append({"Tipo": "ICMS Alíquota", "Cod. ERP": cod_prod, "Descrição": item_c['Descricao_WINTHOR'], "Detalhe": f"XML: {icms_xml}% vs ERP: {icms_csv}% / Guia: {st_guia}%"})
                    if linha['Status_Red_BC'] == 'DIVERGENTE': alertas_atuais.append({"Tipo": "Redução BC", "Cod. ERP": cod_prod, "Descrição": item_c['Descricao_WINTHOR'], "Detalhe": f"XML: {pred_xml}% vs ERP: {red_csv}%"})
                    if linha['Status_NCM'] == 'DIVERGENTE': alertas_atuais.append({"Tipo": "NCM", "Cod. ERP": cod_prod, "Descrição": item_c['Descricao_WINTHOR'], "Detalhe": f"XML: {item_x['NCM_NF']} vs ERP: {item_c['NCM_WINTHOR']}"})
                    if linha['Status_Qtd'] == 'DIVERGENTE': alertas_atuais.append({"Tipo": "Quantidade", "Cod. ERP": cod_prod, "Descrição": item_c['Descricao_WINTHOR'], "Detalhe": f"XML: {item_x['Qtd_NF']} vs ERP: {item_c['Qtd_WINTHOR']}"})

                    if alertas_atuais:
                        contagem_erros_produto[cod_prod] = contagem_erros_produto.get(cod_prod, 0) + len(alertas_atuais)
                        alertas_resumo.extend(alertas_atuais)

                elif item_x and not item_c:
                    linha['Status_Geral'] = 'SOBRA NF (SEM PAR)'
                    total_sobras_nf += 1
                    for st in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']: linha[st] = '-'
                    alertas_resumo.append({"Tipo": "Sobra NF", "Cod. ERP": "-", "Descrição": item_x['Descricao_NF'], "Detalhe": f"Faturado na linha {item_x['Linha_NF']}, mas ausente do ERP"})

                elif item_c and not item_x:
                    linha['Status_Geral'] = 'SOBRA WINTHOR (SEM PAR)'
                    total_sobras_winthor += 1
                    for st in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']: linha[st] = '-'
                    alertas_resumo.append({"Tipo": "Sobra ERP", "Cod. ERP": item_c['Cod_Produto_WINTHOR'], "Descrição": item_c['Descricao_WINTHOR'], "Detalhe": "Esperado no ERP, mas ausente do XML"})

                dados_relatorio.append(linha)

            def obter_prioridade_alerta(alerta):
                cod = alerta["Cod. ERP"]
                if cod == "-" or "Sobra" in alerta["Tipo"]: return (1, 0)
                num_erros = contagem_erros_produto.get(cod, 0)
                return (0, -num_erros, hash(cod) % 10000)

            alertas_resumo.sort(key=obter_prioridade_alerta)

            # ==========================================
            # 6. CRIANDO EXCEL NA MEMÓRIA (BUFFER)
            # ==========================================
            df_final = pd.DataFrame(dados_relatorio)
            colunas_ordem = [
                'Linha_NF', 'Linha_WINTHOR', 'Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN',
                'Descricao_NF', 'Descricao_WINTHOR', 'Qtd_NF', 'Qtd_WINTHOR', 'Status_Qtd',
                'NCM_NF', 'NCM_WINTHOR', 'Status_NCM', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor', 'Status_ICMS_Aliq',
                'pRedBC_NF', '% ICMS red. Whintor', 'Status_Red_BC', 'Tipo_Calculo_ST'
            ]
            df_final = df_final[colunas_ordem]

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Comparativo_ICMS_e_EAN', index=False, startrow=2)
                workbook = writer.book
                
                # Resumo Analítico
                ws_resumo = workbook.create_sheet(title='Resumo_Analitico', index=0)
                ws_resumo.sheet_view.showGridLines = False
                ws_resumo["B2"] = "PROJETO OLHO DE DEUS - AUDITORIA DE COMPRAS"
                ws_resumo["B2"].font = Font(bold=True, size=15, color="1F4E78")
                ws_resumo["B3"] = f"Relatório consolidado referente à Nota Fiscal Nº: {numero_nf}"
                
                # Métricas
                ws_resumo["B5"] = "MÉTRICA"
                ws_resumo["C5"] = "QUANTIDADE DE ITENS"
                metricas = [
                    ("Total de Itens na Nota (XML)", len(itens_xml)), ("Total de Itens no WinThor (ERP)", len(itens_csv)),
                    ("Itens Pareados com Sucesso", total_pareados), ("Itens Sobrando no XML", total_sobras_nf),
                    ("Itens Sobrando no ERP", total_sobras_winthor)
                ]
                for idx, (m_name, m_val) in enumerate(metricas, start=6):
                    ws_resumo.cell(row=idx, column=2, value=m_name)
                    ws_resumo.cell(row=idx, column=3, value=m_val)

            excel_buffer.seek(0)
            
            st.success("Auditoria finalizada com sucesso!")
            st.download_button(
                label="📥 Baixar Relatório Executivo (Excel)",
                data=excel_buffer,
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
