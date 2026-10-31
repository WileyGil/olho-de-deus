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
    
    # Tratamento contra notação científica do Excel ou ERP
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
st.set_page_config(page_title="Espelho NF x WinThor - Olho de Deus", layout="wide")
st.title("Comparativo de Informações: Nota Fiscal x WinThor")
st.markdown("Faça o upload do arquivo XML da nota e do arquivo do WinThor (.csv ou .xlsx) para gerar o espelho lado a lado das informações seguindo a ordem do XML.")

col1, col2 = st.columns(2)
with col1:
    arquivo_xml = st.file_uploader("Upload do XML (Fornecedor)", type=["xml"])
with col2:
    arquivo_dados = st.file_uploader("Upload do Relatório WinThor", type=["csv", "xlsx"])

if arquivo_xml and arquivo_dados:
    if st.button("Gerar Relatório Comparativo"):
        with st.spinner("Extraindo e alinhando informações do XML e do WinThor..."):
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
                ean_raw = prod.find('ns:cEAN', ns).text if prod.find('ns:cEAN', ns) is not None else ''
                ean_trib_raw = prod.find('ns:cEANTrib', ns).text if prod.find('ns:cEANTrib', ns) is not None else ''

                ean_exibicao = limpar_numero(ean_raw) if limpar_ean(ean_raw) != '' else limpar_numero(ean_trib_raw)

                ncm_raw = limpar_numero(prod.find('ns:NCM', ns).text) if prod.find('ns:NCM', ns) is not None else ''
                ncm = ncm_raw.zfill(8) if ncm_raw else ''

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
                    'NCM_NF': ncm,
                    'Qtd_NF': qtd,
                    'Vl_Unit_NF': vUn,
                    'Vl_Total_NF': vTot,
                    'ICMS_NF': icms,
                    'pRedBC_NF': pRedBC
                })

            data_simples = datetime.now().strftime('%Y%m%d')
            nome_arquivo_saida = f'espelho_NF_{numero_nf}_{data_simples}.xlsx'

            # ==========================================
            # 4. EXTRAÇÃO DA PLANILHA (WINTHOR)
            # ==========================================
            if arquivo_dados.name.lower().endswith('.xlsx'):
                df_csv = pd.read_excel(arquivo_dados, dtype=str)
            else:
                df_csv = pd.read_csv(arquivo_dados, sep=';', encoding='latin1', dtype=str)
                if len(df_csv.columns) == 1:
                    arquivo_dados.seek(0)
                    df_csv = pd.read_csv(arquivo_dados, sep=',', encoding='latin1', dtype=str)

            df_csv.columns = df_csv.columns.str.strip()

            col_tipo_calc = next((c for c in df_csv.columns if 'tipo' in c.lower() and ('cálculo' in c.lower() or 'calculo' in c.lower())), None)
            if not col_tipo_calc and len(df_csv.columns) >= 2:
                col_tipo_calc = df_csv.columns[-2]

            itens_csv = []
            for index, row in df_csv.iterrows():
                ean_win_raw = row.get('Cód. barra', '')
                ncm_win_raw = limpar_numero(row.get('NCM', ''))
                ncm_win = ncm_win_raw.zfill(8) if ncm_win_raw else ''

                tipo_calc_raw = str(row.get(col_tipo_calc, '')).strip().upper() if col_tipo_calc else ''
                if tipo_calc_raw == 'N':
                    tipo_calc_fmt = 'NOTA FISCAL'
                elif tipo_calc_raw == 'G':
                    tipo_calc_fmt = 'GUIA'
                else:
                    tipo_calc_fmt = str(row.get(col_tipo_calc, '')).strip() if col_tipo_calc else ''

                itens_csv.append({
                    'Linha_WINTHOR': index + 1,
                    'Cod_Produto_WINTHOR': limpar_numero(row.get('Cód. produto', '')),
                    'EAN_WINTHOR': limpar_numero(ean_win_raw),
                    'Descricao_WINTHOR': str(row.get('Descrição', '')).strip(),
                    'Qtd_WINTHOR': converte_float_br(row.get('Qtde Pré-Entrada', 0.0)),
                    'NCM_WINTHOR': ncm_win,
                    'ICMS_WINTHOR': converte_float_br(row.get('%ICMS', 0.0)),
                    '% ICMS red. Whintor': converte_float_br(row.get('% ICMS red.', 0.0)),
                    '%Alíq.Ext. ST Guia_Whintor': converte_float_br(row.get('%Alíq.Ext.ST Guia', 0.0)),
                    'Tipo_Calculo_ST': tipo_calc_fmt
                })

            # ==========================================
            # 5. PAREAMENTO LADO A LADO (BASEADO NA ORDEM)
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
                        '%Alíq.Ext. ST Guia_Whintor': '-', 'Tipo_Calculo_ST': '-'
                    })

                linha['Sobra'] = not (item_x and item_c)
                dados_relatorio.append(linha)

            # ==========================================
            # 6. CRIANDO EXCEL NA MEMÓRIA (BUFFER)
            # ==========================================
            df_final = pd.DataFrame(dados_relatorio)
            colunas_ordem = [
                'Linha_NF', 'Linha_WINTHOR',
                'Cod_Produto_WINTHOR', 'EAN_NF', 'EAN_WINTHOR',
                'Descricao_NF', 'Descricao_WINTHOR',
                'Qtd_NF', 'Qtd_WINTHOR',
                'NCM_NF', 'NCM_WINTHOR',
                'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor',
                'pRedBC_NF', '% ICMS red. Whintor',
                'Tipo_Calculo_ST'
            ]
            df_final = df_final[colunas_ordem]

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Espelho_NF_x_WinThor', index=False, startrow=2)
                
                workbook = writer.book
                worksheet = workbook['Espelho_NF_x_WinThor']
                
                worksheet.sheet_view.showGridLines = False
                worksheet.sheet_view.zoomScale = 80
                worksheet.freeze_panes = 'A4'
                
                last_col_letter = get_column_letter(len(colunas_ordem))
                worksheet.auto_filter.ref = f"A3:{last_col_letter}{len(df_final) + 3}"

                # Estilos de Preenchimento e Fontes
                color_destaque_xml = "1F4E78"  # Azul Escuro para NF/XML
                color_neutro_geral = "E0E2E5"  # Cinza Claro para WinThor/ERP

                fill_cab_xml = PatternFill(start_color=color_destaque_xml, end_color=color_destaque_xml, fill_type="solid")
                fill_cab_neutro = PatternFill(start_color=color_neutro_geral, end_color=color_neutro_geral, fill_type="solid")
                fill_zebrado = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
                fill_cinza_sobra = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

                font_super_header_xml = Font(bold=True, size=11, color="FFFFFF")
                font_super_header_erp = Font(bold=True, size=11, color="333333")
                font_header_default_xml = Font(bold=True, size=10, color="FFFFFF")
                font_header_default_erp = Font(bold=True, size=10, color="333333")
                font_normal = Font(color="000000", bold=False)

                borda_branca = Border(
                    left=Side(style='thin', color='FFFFFF'), right=Side(style='thin', color='FFFFFF'),
                    top=Side(style='thin', color='FFFFFF'), bottom=Side(style='thin', color='FFFFFF')
                )
                borda_divisoria_direita = Border(
                    left=Side(style='thin', color='FFFFFF'), right=Side(style='medium', color='BFBFBF'),
                    top=Side(style='thin', color='FFFFFF'), bottom=Side(style='thin', color='FFFFFF')
                )

                # Super-grupos ajustados para as 17 colunas informativas (Nota x CSV)
                super_grupos = [
                    ("ORDEM DOS ITENS", 1, 2, fill_cab_neutro, font_super_header_erp),
                    ("CÓDIGOS E EAN (NF x WINTHOR)", 3, 5, fill_cab_xml, font_super_header_xml),
                    ("DESCRIÇÕES (NF x WINTHOR)", 6, 7, fill_cab_neutro, font_super_header_erp),
                    ("QUANTIDADES (NF x WINTHOR)", 8, 9, fill_cab_xml, font_super_header_xml),
                    ("CLASSIFICAÇÃO NCM (NF x WINTHOR)", 10, 11, fill_cab_neutro, font_super_header_erp),
                    ("ALÍQUOTAS DE ICMS (NF x WINTHOR)", 12, 14, fill_cab_xml, font_super_header_xml),
                    ("REDUÇÃO BASE DE CÁLCULO (NF x WINTHOR)", 15, 16, fill_cab_neutro, font_super_header_erp),
                    ("CÁLCULO ST", 17, 17, fill_cab_xml, font_super_header_xml)
                ]

                colunas_fim_bloco = (2, 5, 7, 9, 11, 14, 16, 17)

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
                            if c == col_end and col_end != len(colunas_ordem):
                                cell_target.border = borda_divisoria_direita
                            else:
                                cell_target.border = borda_branca

                worksheet.row_dimensions[1].height = 18
                worksheet.row_dimensions[2].height = 18
                worksheet.row_dimensions[3].height = 25

                # Formatação da Linha 3 (Cabeçalhos das Colunas: Azul = NF, Cinza = WinThor)
                headers = colunas_ordem
                for col_idx, col_name in enumerate(headers, 1):
                    cell = worksheet.cell(row=3, column=col_idx)
                    cell.value = col_name
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.border = borda_divisoria_direita if col_idx in colunas_fim_bloco else borda_branca

                    if 'NF' in col_name:
                        cell.fill = fill_cab_xml
                        cell.font = font_header_default_xml
                    else:
                        cell.fill = fill_cab_neutro
                        cell.font = font_header_default_erp

                align_center_cols = [
                    'Linha_NF', 'Linha_WINTHOR', 'EAN_NF', 'EAN_WINTHOR', 'Qtd_NF', 'Qtd_WINTHOR',
                    'NCM_NF', 'NCM_WINTHOR', 'ICMS_NF', 'ICMS_WINTHOR', '%Alíq.Ext. ST Guia_Whintor',
                    'pRedBC_NF', '% ICMS red. Whintor', 'Tipo_Calculo_ST'
                ]

                # Formatação das Linhas de Dados
                for row in range(4, len(df_final) + 4):
                    linha_dict = dados_relatorio[row - 4]
                    is_even = (row % 2 == 0)
                    is_sobra = linha_dict.get('Sobra', False)

                    for col_idx in range(1, len(headers) + 1):
                        cell = worksheet.cell(row=row, column=col_idx)
                        col_name = headers[col_idx - 1]

                        if is_sobra:
                            cell.fill = fill_cinza_sobra
                        elif is_even:
                            cell.fill = fill_zebrado

                        cell.border = borda_divisoria_direita if col_idx in colunas_fim_bloco else borda_branca
                        cell.font = font_normal

                        if col_name == 'Cod_Produto_WINTHOR' and str(cell.value).strip() not in ('', '-'):
                            try:
                                cell.value = int(float(cell.value))
                                cell.number_format = '0'
                            except ValueError:
                                pass
                        elif col_name in ('NCM_NF', 'NCM_WINTHOR') and str(cell.value).strip() not in ('', '-'):
                            cell.value = str(cell.value).strip().zfill(8)
                            cell.number_format = '@'
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
                        elif col_name in align_center_cols:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            cell.alignment = Alignment(horizontal='left', vertical='center')

                # Ajuste de largura das colunas (ignorando as linhas 1 e 2 mescladas)
                for col in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(col[0].column)
                    for cell in col[2:]:
                        try:
                            if cell.value is not None:
                                max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                    worksheet.column_dimensions[column_letter].width = max(max_length + 4, 12)

            excel_buffer.seek(0)
            
            # Resumo informativo na tela antes do download
            m1, m2, m3 = st.columns(3)
            m1.metric("Nota Fiscal", numero_nf)
            m2.metric("Itens no XML (NF)", len(itens_xml))
            m3.metric("Itens no CSV (WinThor)", len(itens_csv))

            st.success("Relatório comparativo gerado com sucesso!")
            st.download_button(
                label="📥 Baixar Espelho NF x WinThor (Excel)",
                data=excel_buffer.getvalue(),
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
