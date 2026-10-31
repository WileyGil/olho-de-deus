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
        except:
            pass
    match_nums = re.findall(r'\d+', str(uCom))
    if match_nums:
        try:
            return float(match_nums[-1])
        except:
            pass
    return 1.0

# ==========================================
# 2. INTERFACE STREAMLIT (DASHBOARD COMPACTO)
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
    </style>
""", unsafe_allow_html=True)

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

# Seção de Seleção de Blocos com a opção "Marcar/Desmarcar Todos"
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
    col_run1, col_run2 = st.columns([3, 1])
    with col_run1:
        st.success("✅ Arquivos prontos para processamento cruzado.")
    with col_run2:
        processar = st.button("🚀 Processar", use_container_width=True, type="primary")

    if processar:
        with st.spinner("Analisando itens e cruzando tributações com base no XML..."):
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
                    'EAN_Comp_NF': ean_comp,
                    'EANTrib_Comp_NF': ean_trib_comp,
                    'NCM_NF': ncm,
                    'Qtd_NF': qCom_xml,
                    '_Fator_Embalagem': fator,
                    'Vl_Unit_NF': vUn,
                    'Vl_Total_NF': vTot,
                    'Origem_XML': origem_desc,
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
                        'Qtd_NF': '-', '_Fator_Embalagem': 1.0, 'Vl_Unit_NF': '-', 'Vl_Total_NF': '-', 'Origem_XML': '-', 'ICMS_NF': '-', 'pRedBC_NF': '-'
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

            # ==========================================
            # FILTRAGEM DE COLUNAS CONFORME OS BLOCOS SELECIONADOS
            # ==========================================
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

            # ==========================================
            # 6. CRIANDO EXCEL NA MEMÓRIA (BUFFER)
            # ==========================================
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
                    linha_dict = dados_relatorio[row-4]
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
                    rule = FormulaRule(formula=[f'MOD(ROW(),2)=0'], stopIfTrue=True, fill=fill_zebrado)
                    worksheet.conditional_formatting.add(f'{l_col}4:{l_col}{len(df_final) + 3}', rule)

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

            excel_buffer.seek(0)
            
            st.markdown("---")
            st.success("✨ Relatório gerado com sucesso com base nas suas preferências!")
            st.download_button(
                label="📥 Baixar Relatório Executivo em Excel",
                data=excel_buffer.getvalue(),
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
