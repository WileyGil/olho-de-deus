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

def obter_v_icms_xml(det_node, ns):
    vICMS = det_node.find('.//ns:vICMS', ns)
    if vICMS is not None and vICMS.text:
        return float(vICMS.text)
    return 0.0

# ==========================================
# 2. INTERFACE STREAMLIT
# ==========================================
st.set_page_config(page_title="Auditoria Olho de Deus", layout="wide")
st.title("Auditoria de Notas")
st.markdown("Faça o upload do arquivo XML da nota e do arquivo do WinThor (.csv ou .xlsx) para cruzar as informações com foco na divergência de ICMS.")

col1, col2 = st.columns(2)
with col1:
    arquivo_xml = st.file_uploader("Upload do XML (Fornecedor)", type=["xml"])
with col2:
    arquivo_dados = st.file_uploader("Upload do Relatório WinThor", type=["csv", "xlsx"])

if arquivo_xml and arquivo_dados:
    if st.button("Processar Auditoria"):
        with st.spinner("Analisando e rastreando diferenças de ICMS..."):
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
                icms_aliq = obter_icms_xml(det, ns)
                v_icms_item = obter_v_icms_xml(det, ns)

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
                    'ICMS_NF': icms_aliq,
                    'vICMS_NF': v_icms_item,
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
                v_prod_win = converte_float_br(row.get('Vl. Total', 0.0)) # Caso exista coluna de valor total no ERP
                p_icms_win = converte_float_br(row.get('%ICMS', 0.0))
                
                itens_csv.append({
                    'Linha_WINTHOR': index + 1,
                    'Cod_Produto_WINTHOR': limpar_numero(row.get('Cód. produto', '')),
                    'EAN_WINTHOR': limpar_numero(ean_win_raw),
                    'EAN_Comp_WINTHOR': limpar_ean(ean_win_raw),
                    'Descricao_WINTHOR': str(row.get('Descrição', '')).strip(),
                    'Qtd_WINTHOR': converte_float_br(row.get('Qtde Pré-Entrada', 0.0)),
                    'NCM_WINTHOR': limpar_numero(row.get('NCM', '')),
                    'ICMS_WINTHOR': p_icms_win,
                    '% ICMS red. Whintor': converte_float_br(row.get('% ICMS red.', 0.0)),
                    '%Alíq.Ext. ST Guia_Whintor': converte_float_br(row.get('%Alíq.Ext.ST Guia', 0.0)),
                    'Tipo_Calculo_WINTHOR': str(row.get(col_tipo_calc, '')).strip()
                })

            # ==========================================
            # 5. CRUZAMENTO E AUDITORIA DE ICMS ITEM A ITEM
            # ==========================================
            dados_relatorio = []

            for item_x, item_c in zip_longest(itens_xml, itens_csv, fillvalue=None):
                linha = {}

                if item_x:
                    linha.update(item_x)
                else:
                    linha.update({
                        'Linha_NF': '-', 'Descricao_NF': '-', 'EAN_NF': '-', 'NCM_NF': '-',
                        'Qtd_NF': '-', 'Vl_Unit_NF': '-', 'Vl_Total_NF': '-', 'ICMS_NF': '-', 'vICMS_NF': 0.0, 'pRedBC_NF': '-'
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

                    # Cálculo do Valor do ICMS no ERP estimado para comparar com o XML
                    base_calc_item = item_x['Vl_Total_NF'] # Base aproximada pelo total do produto no XML
                    icms_erp_aplicado = icms_csv if icms_csv > 0 else icms_xml
                    v_icms_erp_estimado = round(base_calc_item * (icms_erp_aplicado / 100.0), 2)
                    v_icms_xml_item = round(item_x['vICMS_NF'], 2)

                    linha['vICMS_ERP_Estimado'] = v_icms_erp_estimado
                    
                    if abs(v_icms_xml_item - v_icms_erp_estimado) <= 0.05:
                        linha['Status_vICMS'] = 'OK'
                        linha['Dif_vICMS'] = 0.0
                    else:
                        linha['Status_vICMS'] = 'DIVERGENTE'
                        linha['Dif_vICMS'] = round(v_icms_xml_item - v_icms_erp_estimado, 2)

                    pred_xml = round(item_x['pRedBC_NF'], 4)
                    red_csv = round(item_c['% ICMS red. Whintor'], 4)
                    linha['Status_Red_BC'] = 'OK' if pred_xml == red_csv else 'DIVERGENTE'

                    tipo_calc = str(item_c.get('Tipo_Calculo_WINTHOR', '')).strip().upper()
                    if tipo_calc == 'N':
                        linha['Tipo_Calculo_ST'] = 'NOTA FISCAL'
                    elif tipo_calc == 'G':
                        linha['Tipo_Calculo_ST'] = 'GUIA'
                    else:
                        linha['Tipo_Calculo_ST'] = item_c['Tipo_Calculo_WINTHOR']

                    linha['Status_NCM'] = 'OK' if str(item_x['NCM_NF']) == str(item_c['NCM_WINTHOR']) else 'DIVERGENTE'
                    linha['Status_Qtd'] = 'OK' if round(item_x['Qtd_NF'], 4) == round(item_c['Qtd_WINTHOR'], 4) else 'DIVERGENTE'

                    # Diagnóstico detalhado de onde está a diferença do ICMS
                    diagnosticos = []
                    if linha['Status_EAN'] == 'DIVERGENTE':
                        diagnosticos.append("EAN Divergente")
                    if linha['Status_ICMS_Aliq'] == 'DIVERGENTE':
                        diagnosticos.append(f"Alíq. ICMS (XML {icms_xml}% vs ERP {icms_csv}%)")
                    if linha['Status_vICMS'] == 'DIVERGENTE':
                        diagnosticos.append(f"Diferença Vl. ICMS (Δ R$ {linha['Dif_vICMS']:+.2f})")
                    if linha['Status_NCM'] == 'DIVERGENTE':
                        diagnosticos.append("NCM Divergente")
                    if linha['Status_Qtd'] == 'DIVERGENTE':
                        diagnosticos.append("Qtd Divergente")

                    linha['Diagnostico_Divergencia'] = " | ".join(diagnosticos) if diagnosticos else "TUDO OK (1 = 1)"

                elif item_x and not item_c:
                    linha['Status_Geral'] = 'SOBRA NF (SEM PAR)'
                    linha['vICMS_ERP_Estimado'] = 0.0
                    linha['Dif_vICMS'] = item_x['vICMS_NF']
                    for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_vICMS', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                        linha[st_col] = '-'
                    linha['Diagnostico_Divergencia'] = "SOBRA NO XML (Ausente no ERP)"

                elif item_c and not item_x:
                    linha['Status_Geral'] = 'SOBRA WINTHOR (SEM PAR)'
                    linha['vICMS_ERP_Estimado'] = 0.0
                    linha['Dif_vICMS'] = 0.0
                    for st_col in ['Status_EAN', 'Status_ICMS_Aliq', 'Status_vICMS', 'Status_Red_BC', 'Tipo_Calculo_ST', 'Status_NCM', 'Status_Qtd']:
                        linha[st_col] = '-'
                    linha['Diagnostico_Divergencia'] = "SOBRA NO ERP (Ausente no XML)"

                dados_relatorio.append(linha)

            # ==========================================
            # 6. CRIANDO EXCEL NA MEMÓRIA (BUFFER)
            # ==========================================
            df_final = pd.DataFrame(dados_relatorio)
            colunas_ordem = [
                'Linha_NF', 'Linha_WINTHOR',
                'Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Status_EAN',
                'Descricao_NF', 'Descricao_WINTHOR',
                'Qtd_NF', 'Qtd_WINTHOR', 'Status_Qtd',
                'NCM_NF', 'NCM_WINTHOR', 'Status_NCM',
                'ICMS_NF', 'ICMS_WINTHOR', 'Status_ICMS_Aliq',
                'vICMS_NF', 'vICMS_ERP_Estimado', 'Dif_vICMS', 'Status_vICMS',
                'pRedBC_NF', '% ICMS red. Whintor', 'Status_Red_BC',
                'Tipo_Calculo_ST', 'Diagnostico_Divergencia'
            ]
            df_final = df_final[colunas_ordem]

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Auditoria_ICMS', index=False, startrow=2)
                
                workbook = writer.book
                worksheet = workbook['Auditoria_ICMS']
                
                worksheet.sheet_view.showGridLines = False
                worksheet.sheet_view.zoomScale = 80

                worksheet.freeze_panes = 'A4'
                worksheet.auto_filter.ref = f"A3:Z{len(df_final) + 3}"

                fill_zebrado = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
                fill_vermelho_claro = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                fill_verde_oliva_claro = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                fill_cinza = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

                font_divergente = Font(color="9C0006", bold=True)
                font_normal = Font(color="000000", bold=False)

                color_destaque_xml = "1F4E78"
                color_neutro_geral = "E0E2E5"

                fill_cab_xml = PatternFill(start_color=color_destaque_xml, end_color=color_destaque_xml, fill_type="solid")
                fill_cab_neutro = PatternFill(start_color=color_neutro_geral, end_color=color_neutro_geral, fill_type="solid")

                font_super_header_xml = Font(bold=True, size=11, color="FFFFFF")
                font_super_header_erp = Font(bold=True, size=11, color="333333")
                font_header_default_xml = Font(bold=True, size=10, color="FFFFFF")
                font_header_default_erp = Font(bold=True, size=10, color="333333")

                borda_branca = Border(
                    left=Side(style='thin', color='FFFFFF'), right=Side(style='thin', color='FFFFFF'),
                    top=Side(style='thin', color='FFFFFF'), bottom=Side(style='thin', color='FFFFFF')
                )
                borda_divisoria_direita = Border(right=Side(style='medium', color='D9D9D9'))

                super_grupos = [
                    ("IDENTIFICAÇÃO", 1, 2, fill_cab_neutro, font_super_header_erp),
                    ("CÓDIGO E CÓDIGO DE BARRAS (EAN)", 3, 6, fill_cab_xml, font_super_header_xml),
                    ("DESCRIÇÕES", 7, 8, fill_cab_neutro, font_super_header_erp),
                    ("QUANTIDADES", 9, 11, fill_cab_xml, font_super_header_xml),
                    ("NCM", 12, 14, fill_cab_xml, font_super_header_xml),
                    ("ALÍQUOTA ICMS (%)", 15, 17, fill_cab_xml, font_super_header_xml),
                    ("VALOR TOTAL DO ICMS (R$) - AUDITORIA DE DIFERENÇA", 18, 21, fill_cab_xml, font_super_header_xml),
                    ("REDUÇÃO BC", 22, 24, fill_cab_xml, font_super_header_xml),
                    ("OUTROS", 25, 26, fill_cab_neutro, font_super_header_erp)
                ]

                for label, col_start, col_end, fill_obj, font_obj in super_grupos:
                    worksheet.merge_cells(start_row=1, start_column=col_start, end_row=2, end_column=col_end)
                    cell = worksheet.cell(row=1, column=col_start)
                    cell.value = label
                    cell.font = font_obj
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                    for r in (1, 2):
                        for c in range(col_start, col_end + 1):
                            cell_target = worksheet.cell(row=r, column=c)
                            cell_target.fill = fill_obj
                            cell_target.border = borda_branca

                worksheet.row_dimensions[1].height = 18
                worksheet.row_dimensions[2].height = 18
                worksheet.row_dimensions[3].height = 25

                headers = colunas_ordem
                for col_idx, col_name in enumerate(headers, 1):
                    cell = worksheet.cell(row=3, column=col_idx)
                    cell.value = col_name
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.border = borda_divisoria_direita if col_idx in (2, 6, 8, 11, 14, 17, 21, 24, 26) else borda_branca

                    if 'NF' in col_name or col_name in ('EAN_NF', 'Qtd_NF', 'NCM_NF', 'ICMS_NF', 'vICMS_NF', 'pRedBC_NF'):
                        cell.fill = fill_cab_xml
                        cell.font = Font(bold=True, size=10, color="FFFFFF")
                    else:
                        cell.fill = fill_cab_neutro
                        cell.font = Font(bold=True, size=10, color="333333")

                align_center_cols = [
                    'Linha_NF', 'Linha_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Qtd_NF', 'Qtd_WINTHOR',
                    'NCM_NF', 'NCM_WINTHOR', 'ICMS_NF', 'ICMS_WINTHOR', 'ICMS_NF', 'vICMS_NF', 'vICMS_ERP_Estimado', 'Dif_vICMS'
                ]

                for row in range(4, len(df_final) + 4):
                    linha_dict = dados_relatorio[row-4]
                    is_even = (row % 2 == 0)

                    for col_idx in range(1, len(headers) + 1):
                        cell = worksheet.cell(row=row, column=col_idx)

                        if is_even:
                            cell.fill = fill_zebrado

                        cell.border = borda_divisoria_direita if col_idx in (2, 6, 8, 11, 14, 17, 21, 24, 26) else borda_branca
                        col_name = headers[col_idx - 1]

                        if col_name == 'Cod_Produto_WINTHOR' and str(cell.value).strip() not in ('', '-'):
                            try:
                                cell.value = int(float(cell.value))
                                cell.number_format = '0'
                            except:
                                pass
                        elif col_name in ('NCM_NF', 'NCM_WINTHOR') and str(cell.value).strip() not in ('', '-'):
                            try:
                                cell.value = int(float(cell.value))
                                cell.number_format = '00000000'
                            except:
                                pass
                        elif col_name in ('EAN_NF', 'EAN_WINTHOR') and str(cell.value).strip() not in ('', '-'):
                            cell.value = str(cell.value).strip()
                            cell.number_format = '@'
                        elif isinstance(cell.value, (int, float)) and cell.value != '-':
                            if 'vICMS' in col_name or 'Dif_vICMS' in col_name:
                                cell.number_format = 'R$ #,##0.00'
                            elif 'pRedBC' in col_name:
                                cell.number_format = '#,##0.0000'
                            elif 'Qtd' in col_name or 'ICMS' in col_name:
                                cell.number_format = '#,##0.00'

                        if col_name == 'Cod_Produto_WINTHOR':
                            cell.alignment = Alignment(horizontal='right', vertical='center')
                        elif col_name in align_center_cols or 'Status' in col_name:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        elif col_name == 'Diagnostico_Divergencia':
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                            if "TUDO OK" in str(cell.value):
                                cell.font = Font(color="006100", bold=True)
                            else:
                                cell.font = Font(color="9C0006", bold=True)

                    if 'SOBRA' in str(linha_dict.get('Status_Geral', '')):
                        for col_idx in range(1, len(headers) + 1):
                            worksheet.cell(row=row, column=col_idx).fill = fill_cinza
                        continue

                    # Pintar divergências de ICMS em vermelho
                    if linha_dict.get('Status_vICMS') == 'DIVERGENTE':
                        idx_vicms = headers.index('vICMS_ERP_Estimado') + 1
                        worksheet.cell(row=row, column=idx_vicms).fill = fill_vermelho_claro
                        worksheet.cell(row=row, column=idx_vicms).font = font_divergente
                        
                        idx_dif = headers.index('Dif_vICMS') + 1
                        worksheet.cell(row=row, column=idx_dif).fill = fill_vermelho_claro
                        worksheet.cell(row=row, column=idx_dif).font = font_divergente

                for col in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(col[0].column)
                    for cell in col:
                        try:
                            if cell.value is not None:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                        except:
                            pass
                    worksheet.column_dimensions[column_letter].width = max(max_length + 4, 12)

                # ==========================================
                # 7. BLOQUEIO DA PLANILHA (PERMITIR APENAS FILTROS)
                # ==========================================
                worksheet.protection.sheet = True
                worksheet.protection.enableAutoFilter = True
                worksheet.protection.enableSort = True

            excel_buffer.seek(0)
            
            st.success("Auditoria finalizada com sucesso!")
            st.download_button(
                label="📥 Baixar Relatório Executivo (Excel)",
                data=excel_buffer.getvalue(),
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
