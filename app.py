import re
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter
from decimal import Decimal, InvalidOperation

import pandas as pd

NS = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

# Limite mínimo de itens com ao menos uma palavra de marca/produto em comum
LIMITE_MARCA = 0.70


# ==========================================================
# 1. EAN
# ==========================================================
def _dv_gtin_ok(digitos: str) -> bool:
    """Valida o dígito verificador GS1 (GTIN-8/12/13/14), completando com zeros à esquerda."""
    d = digitos.zfill(14)
    if len(d) != 14 or not d.isdigit():
        return False
    soma = sum(int(c) * (3 if i % 2 == 0 else 1) for i, c in enumerate(d[:13]))
    return (10 - soma % 10) % 10 == int(d[13])


def normalizar_ean(valor):
    """
    Devolve (ean, status):
      ean    -> só dígitos, SEM zeros à esquerda ('' se vazio / sem GTIN)
      status -> 'ok' | 'vazio' | 'ilegivel'

    'ilegivel' = veio em notação científica e perdeu dígitos (ex.: 7,89862E+12),
    sem como recuperar o EAN real.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return '', 'vazio'
    s = str(valor).strip().upper()
    if s in ('', 'NAN', 'NONE', 'SEM GTIN', 'S/GTIN'):
        return '', 'vazio'

    if 'E' in s and re.fullmatch(r'[0-9.,]+E[+\-]?[0-9]+', s):
        try:
            dec = Decimal(s.replace(',', '.'))
        except InvalidOperation:
            return '', 'ilegivel'
        inteiro = format(dec, 'f').split('.')[0]
        mantissa, expo = re.split(r'E', s)
        digitos_mantissa = len(re.sub(r'[^0-9]', '', mantissa))
        total_digitos = int(expo) + 1
        # Excel só mostra parte dos dígitos: aceita apenas se o DV GS1 confirmar
        if digitos_mantissa < total_digitos and not _dv_gtin_ok(inteiro):
            return '', 'ilegivel'
        return inteiro.lstrip('0'), 'ok'

    if s.endswith('.0'):
        s = s[:-2]
    apenas = re.sub(r'[^0-9]', '', s)
    return apenas.lstrip('0'), ('ok' if apenas.strip('0') else 'vazio')


# ==========================================================
# 2. LEITURA DO CSV/XLSX DO WINTHOR + ORDENAÇÃO
# ==========================================================
def _sem_acento(txt: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFKD', str(txt)) if not unicodedata.combining(c))


def _achar_coluna_seq(colunas):
    for c in colunas:
        n = re.sub(r'[^a-z]', '', _sem_acento(c).lower())
        if 'seq' in n or n in ('item', 'nitem', 'noitem'):
            return c
    return None


def ler_csv_winthor(arquivo):
    """Lê o arquivo, ordena por 'Nº seq.' (numérico) e devolve (df, avisos)."""
    avisos = []
    nome = getattr(arquivo, 'name', '').lower()
    if nome.endswith('.xlsx'):
        df = pd.read_excel(arquivo, dtype=str)
    else:
        bruto = arquivo.read() if hasattr(arquivo, 'read') else open(arquivo, 'rb').read()
        for enc in ('utf-8-sig', 'latin1'):
            try:
                texto = bruto.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        import io
        df = pd.read_csv(io.StringIO(texto), sep=';', dtype=str)
    df.columns = df.columns.str.strip()

    # remove linhas totalmente vazias (rodapés, separadores)
    df = df.dropna(how='all').reset_index(drop=True)

    col_seq = _achar_coluna_seq(df.columns)
    if col_seq is None:
        avisos.append("Coluna de sequência ('Nº seq.') não encontrada: mantida a ordem original do arquivo.")
        df['_seq'] = range(1, len(df) + 1)
    else:
        seq = pd.to_numeric(df[col_seq].astype(str).str.strip(), errors='coerce')
        invalidas = seq.isna()
        if invalidas.any():
            avisos.append(f"{int(invalidas.sum())} linha(s) sem número de sequência válido foram ignoradas (rodapé/total?).")
            df = df[~invalidas].copy()
            seq = seq[~invalidas]
        df['_seq'] = seq.astype(int)
        df = df.sort_values('_seq', kind='stable').reset_index(drop=True)

    return df, avisos


def montar_itens_csv(df):
    """Converte o DataFrame ordenado na lista de itens usada pelo cruzamento."""
    def num(v):
        return re.sub(r'[^0-9]', '', str(v).strip().removesuffix('.0')) if not pd.isna(v) else ''

    def flt(v):
        if pd.isna(v) or str(v).strip() == '':
            return 0.0
        s = str(v).strip()
        s = s.replace('.', '').replace(',', '.') if ('.' in s and ',' in s) else s.replace(',', '.')
        try:
            return float(s)
        except ValueError:
            return 0.0

    col_tipo = next((c for c in df.columns if 'tipo' in c.lower() and ('cálculo' in c.lower() or 'calculo' in c.lower())), None)
    itens = []
    for _, row in df.iterrows():
        ean, st_ean = normalizar_ean(row.get('Cód. barra', ''))
        itens.append({
            'Linha_WINTHOR': int(row['_seq']),
            'Cod_Produto_WINTHOR': num(row.get('Cód. produto', '')),
            'EAN_WINTHOR': ean,
            'EAN_Comp_WINTHOR': ean,
            'EAN_Status_WINTHOR': st_ean,
            'Descricao_WINTHOR': str(row.get('Descrição', '')).strip(),
            'Qtd_WINTHOR': flt(row.get('Qtde Pré-Entrada', 0.0)),
            'NCM_WINTHOR': num(row.get('NCM', '')),
            'ICMS_WINTHOR': flt(row.get('%ICMS', 0.0)),
            '% ICMS red. Whintor': flt(row.get('% ICMS red.', 0.0)),
            '%Alíq.Ext. ST Guia_Whintor': flt(row.get('%Alíq.Ext.ST Guia', 0.0)),
            'Tipo_Calculo_WINTHOR': str(row.get(col_tipo, '')).strip() if col_tipo else '',
        })
    return itens


# ==========================================================
# 3. LEITURA DO XML (mesma lógica que já existia no app)
# ==========================================================
def ler_xml_nfe(arquivo, extrair_fator_embalagem, obter_descricao_origem, limpar_numero, limpar_ean):
    """Devolve (numero_nf, itens_xml). As funções auxiliares do app são injetadas."""
    root = ET.parse(arquivo).getroot()
    nNF = root.find('.//ns:ide/ns:nNF', NS)
    numero_nf = nNF.text.strip() if nNF is not None and nNF.text else 'DESCONHECIDO'

    def txt(no, tag):
        e = no.find(f'ns:{tag}', NS)
        return e.text if e is not None and e.text else ''

    itens = []
    for det in root.findall('.//ns:det', NS):
        prod = det.find('ns:prod', NS)
        try:
            n_item = int(det.attrib.get('nItem', '0'))
        except ValueError:
            n_item = det.attrib.get('nItem', '')

        ean_raw, ean_trib_raw = txt(prod, 'cEAN'), txt(prod, 'cEANTrib')
        ean_comp, ean_trib_comp = limpar_ean(ean_raw), limpar_ean(ean_trib_raw)
        ean_exib = limpar_numero(ean_raw) if ean_comp != '' else limpar_numero(ean_trib_raw)
        u_com = txt(prod, 'uCom') or 'UN'

        icms, origem = 0.0, '0 - Nacional'
        icms_node = det.find('.//ns:ICMS', NS)
        if icms_node is not None and len(icms_node) > 0:
            sub = icms_node[0]
            p = sub.find('ns:pICMS', NS)
            if p is not None and p.text:
                icms = float(p.text)
            o = sub.find('ns:orig', NS)
            if o is not None and o.text:
                origem = obter_descricao_origem(o.text)
        # busca pRedBC dentro do próprio nó do ICMS (e não em todo o <det>)
        pred_node = icms_node[0].find('ns:pRedBC', NS) if icms_node is not None and len(icms_node) > 0 else None
        pred = float(pred_node.text) if pred_node is not None and pred_node.text else 0.0

        itens.append({
            'Linha_NF': n_item,
            'Descricao_NF': txt(prod, 'xProd'),
            'EAN_NF': ean_exib,
            'EAN_Comp_NF': ean_comp,
            'EANTrib_Comp_NF': ean_trib_comp,
            'NCM_NF': limpar_numero(txt(prod, 'NCM')),
            'Qtd_NF': float(txt(prod, 'qCom') or 0.0),
            '_Fator_Embalagem': extrair_fator_embalagem(u_com),
            'Vl_Unit_NF': float(txt(prod, 'vUnCom') or 0.0),
            'Vl_Total_NF': float(txt(prod, 'vProd') or 0.0),
            'Origem_XML': origem,
            'ICMS_NF': icms,
            'pRedBC_NF': pred,
        })
    return numero_nf, itens


# ==========================================================
# 4. PORTÃO DE VALIDAÇÃO DO UPLOAD
# ==========================================================
_IGNORAR = {'COM', 'PARA', 'DOS', 'DAS', 'UNID', 'UND', 'CX', 'KIT', 'PCT', 'FR', 'REF', 'ESC', 'ATIV'}


def _tokens(desc):
    t = re.findall(r'[A-Z0-9]+', _sem_acento(desc).upper())
    return {x for x in t if len(x) >= 3 and not any(ch.isdigit() for ch in x) and x not in _IGNORAR}


def _numeros_do_nome(nome):
    stem = re.sub(r'\.[A-Za-z0-9]+$', '', str(nome))
    return {g.lstrip('0') for g in re.findall(r'\d{3,}', stem)}


def validar_upload(itens_xml, itens_csv, numero_nf, nome_arquivo_csv):
    """
    Devolve dict:
      bloqueios      -> motivos que impedem a comparação (lista vazia = liberado)
      bloqueio_nf    -> True se o ÚNICO motivo é o número da NF do nome do arquivo (liberável pelo usuário)
      avisos / infos -> mensagens informativas
      marca_em_comum -> palavras de marca/produto comuns aos dois lados
    """
    bloqueios, avisos, infos = [], [], []
    r = {'bloqueios': bloqueios, 'avisos': avisos, 'infos': infos, 'marca_em_comum': [], 'bloqueio_nf': False}

    # (a) quantidade de itens
    n_x, n_c = len(itens_xml), len(itens_csv)
    if n_x != n_c:
        bloqueios.append(f"Quantidade de itens diferente: XML tem {n_x} e o arquivo do WinThor tem {n_c}.")
    else:
        infos.append(f"Quantidade de itens confere: {n_x}.")

    # (b) EAN legível (notação científica truncada impede comparação exata)
    ilegiveis = [i['Linha_WINTHOR'] for i in itens_csv if i['EAN_Status_WINTHOR'] == 'ilegivel']
    if ilegiveis:
        lista = ', '.join(map(str, ilegiveis[:10])) + ('…' if len(ilegiveis) > 10 else '')
        bloqueios.append(
            f"{len(ilegiveis)} EAN(s) do WinThor estão em notação científica truncada e não podem ser recuperados "
            f"(Nº seq.: {lista}). Reexporte com a coluna 'Cód. barra' formatada como texto.")

    # (c) mesma marca/produto, comparando item a item na ordem (CSV já está ordenado por Nº seq.)
    if n_x == n_c and n_x > 0:
        compat = 0
        cont_x, cont_c = Counter(), Counter()
        for ix, ic in zip(itens_xml, itens_csv):
            tx, tc = _tokens(ix['Descricao_NF']), _tokens(ic['Descricao_WINTHOR'])
            compat += bool(tx & tc)
            cont_x.update(tx)
            cont_c.update(tc)
        pct = compat / n_x
        comum = [t for t, _ in sorted(((t, min(cont_x[t], cont_c[t])) for t in cont_x.keys() & cont_c.keys()),
                                      key=lambda p: -p[1])[:4]]
        r['marca_em_comum'] = comum
        if pct < LIMITE_MARCA:
            bloqueios.append(
                f"Os produtos não parecem ser da mesma nota: só {compat} de {n_x} itens ({pct:.0%}) "
                f"têm descrição compatível entre XML e WinThor.")
        else:
            infos.append(f"Produtos compatíveis em {compat} de {n_x} itens ({pct:.0%})"
                         + (f"; termos em comum: {', '.join(comum)}." if comum else '.'))

    # (d) número da NF no nome do arquivo (crivo extra)
    candidatos = _numeros_do_nome(nome_arquivo_csv)
    nf_norm = str(numero_nf).strip().lstrip('0')
    if candidatos:
        if nf_norm in candidatos:
            infos.append(f"Número da NF confere com o nome do arquivo ({numero_nf}).")
        else:
            bloqueios.append(
                f"O nome do arquivo ({nome_arquivo_csv}) não contém o número da NF do XML ({numero_nf}).")
            r['bloqueio_nf'] = len(bloqueios) == 1
    else:
        avisos.append("O nome do arquivo não traz o número da NF; a conferência foi feita só por itens e produtos.")

    return r
