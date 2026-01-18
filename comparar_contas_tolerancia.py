# -*- coding: utf-8 -*-
"""
Script para comparar contas médicas entre Excel (TASY) e arquivos XML (TISS)
VERSÃO COM TOLERÂNCIA DE 1 CENTAVO para evitar falsos positivos de arredondamento
"""

import os
import pandas as pd
from collections import defaultdict
import xml.etree.ElementTree as ET
import warnings
warnings.filterwarnings('ignore')

# Configurações
PASTA_XML = r"C:\Users\AMH\Desktop\meu-site\xml"
ARQUIVO_EXCEL = r"C:\Users\AMH\Desktop\meu-site\Unimed conta recalculadas.xlsx"
ARQUIVO_SAIDA = r"C:\Users\AMH\Desktop\meu-site\Relatorio_Comparacao_Tolerancia1Centavo.xlsx"
CODIGO_PRESTADOR_VALIDO = "110020"
CONTAS_IGNORAR = {74078, 75059, 60282}
TOLERANCIA_PRECO = 0.01  # Tolerância de 1 centavo

# Namespace TISS
NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}


def extrair_texto(elemento, xpath):
    """Extrai texto de um elemento XML usando xpath"""
    el = elemento.find(xpath, NS)
    return el.text.strip() if el is not None and el.text else None


def processar_procedimento(proc, numero_lote, numero_guia, arquivo_xml):
    """Processa um procedimento executado e retorna dict com dados"""
    cod_prestador = extrair_texto(proc, './/ans:codigoPrestadorNaOperadora')

    codigo = extrair_texto(proc, './/ans:codigoProcedimento')
    if not codigo:
        return None

    qtd_str = extrair_texto(proc, './/ans:quantidadeExecutada')
    valor_unit_str = extrair_texto(proc, './/ans:valorUnitario')
    valor_total_str = extrair_texto(proc, './/ans:valorTotal')

    if not all([qtd_str, valor_unit_str]):
        return None

    try:
        qtd = float(qtd_str)
        valor_unit = float(valor_unit_str)
        valor_total = float(valor_total_str) if valor_total_str else qtd * valor_unit
    except:
        return None

    return {
        'NR_SEQ_PROTOCOLO': numero_lote,
        'NR_INTERNO_CONTA': numero_guia,
        'ITEM_CD_CONVENIO': codigo,
        'QT_ITEM': qtd,
        'PRECO_UNITARIO': valor_unit,
        'PRECO_TOTAL': valor_total,
        'ARQUIVO_XML': arquivo_xml,
        'COD_PRESTADOR': cod_prestador
    }


def processar_guia(guia, numero_lote, arquivo_xml):
    """Processa uma guia e extrai todos os procedimentos"""
    itens = []

    numero_guia = extrair_texto(guia, './/ans:numeroGuiaPrestador')
    if not numero_guia:
        return itens

    try:
        if int(numero_guia) in CONTAS_IGNORAR:
            return itens
    except:
        pass

    for proc in guia.findall('.//ans:procedimentoExecutado', NS):
        item = processar_procedimento(proc, numero_lote, numero_guia, arquivo_xml)
        if item:
            itens.append(item)

    for serv in guia.findall('.//ans:servicosExecutados', NS):
        item = processar_procedimento(serv, numero_lote, numero_guia, arquivo_xml)
        if item:
            itens.append(item)

    return itens


def processar_arquivo_xml(caminho_xml):
    """Processa um arquivo XML e extrai todos os itens"""
    itens = []
    nome_arquivo = os.path.basename(caminho_xml)

    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

        numero_lote = extrair_texto(root, './/ans:numeroLote')

        for guia in root.findall('.//ans:guiaSP-SADT', NS):
            itens.extend(processar_guia(guia, numero_lote, nome_arquivo))

        for guia in root.findall('.//ans:guiaConsulta', NS):
            itens.extend(processar_guia(guia, numero_lote, nome_arquivo))

        for guia in root.findall('.//ans:guiaResumoInternacao', NS):
            itens.extend(processar_guia(guia, numero_lote, nome_arquivo))

    except Exception as e:
        print(f"Erro em {nome_arquivo}: {e}")

    return itens


def agrupar_com_tolerancia(df, colunas_grupo, tolerancia):
    """
    Agrupa itens considerando tolerância de preço.
    Itens com mesmo código e conta, com preços dentro da tolerância, são agrupados juntos.
    """
    # Ordenar por conta, código e preço
    df = df.sort_values(['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'PRECO_UNITARIO']).copy()

    # Criar grupo baseado em tolerância
    df['GRUPO_PRECO'] = 0
    grupo_atual = 0
    ultimo_preco = None
    ultima_chave = None

    for idx, row in df.iterrows():
        chave_atual = f"{row['NR_INTERNO_CONTA']}_{row['ITEM_CD_CONVENIO']}"

        if chave_atual != ultima_chave:
            grupo_atual += 1
            ultimo_preco = row['PRECO_UNITARIO']
        elif abs(row['PRECO_UNITARIO'] - ultimo_preco) > tolerancia:
            grupo_atual += 1
            ultimo_preco = row['PRECO_UNITARIO']

        df.at[idx, 'GRUPO_PRECO'] = grupo_atual
        ultima_chave = chave_atual

    # Agrupar
    resultado = df.groupby(['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'GRUPO_PRECO'], as_index=False).agg({
        'NR_SEQ_PROTOCOLO': 'first',
        'QT_ITEM': 'sum',
        'PRECO_UNITARIO': 'mean',  # Média dos preços no grupo
        'PRECO_TOTAL': 'sum',
        'DS_ITEM': 'first' if 'DS_ITEM' in df.columns else lambda x: '',
        'ARQUIVO_XML': lambda x: ', '.join(set(x)) if 'ARQUIVO_XML' in df.columns else ''
    })

    return resultado


def main():
    print("\n" + "=" * 70)
    print("COMPARAÇÃO COM TOLERÂNCIA DE 1 CENTAVO")
    print("=" * 70)

    # Etapa 1: Extrair dados dos XMLs
    print("\nETAPA 1: Extraindo dados dos arquivos XML...")
    todos_itens = []
    arquivos_processados = 0

    for root_dir, dirs, files in os.walk(PASTA_XML):
        for file in files:
            if file.endswith('.xml'):
                caminho = os.path.join(root_dir, file)
                itens = processar_arquivo_xml(caminho)
                for item in itens:
                    if item.get('COD_PRESTADOR') and item['COD_PRESTADOR'] != CODIGO_PRESTADOR_VALIDO:
                        continue
                    todos_itens.append(item)
                arquivos_processados += 1
                if arquivos_processados % 200 == 0:
                    print(f"  Processados {arquivos_processados} arquivos...")

    print(f"  Total: {arquivos_processados} arquivos, {len(todos_itens)} itens")

    # Etapa 2: Processar Excel
    print("\nETAPA 2: Processando arquivo Excel...")
    df_excel = pd.read_excel(ARQUIVO_EXCEL)
    df_excel = df_excel[~df_excel['NR_INTERNO_CONTA'].isin(CONTAS_IGNORAR)]

    df_excel['NR_SEQ_PROTOCOLO'] = df_excel['NR_SEQ_PROTOCOLO'].astype(str)
    df_excel['NR_INTERNO_CONTA'] = df_excel['NR_INTERNO_CONTA'].astype(str)
    df_excel['ITEM_CD_CONVENIO'] = df_excel['ITEM_CD_CONVENIO'].astype(str)
    df_excel['QT_ITEM'] = pd.to_numeric(df_excel['QT_ITEM'], errors='coerce').fillna(0)
    df_excel['PRECO_UNITARIO'] = pd.to_numeric(df_excel['PRECO_UNITARIO'], errors='coerce').fillna(0)
    df_excel['PRECO_TOTAL'] = pd.to_numeric(df_excel['PRECO_TOTAL'], errors='coerce').fillna(0)

    print(f"  Total: {len(df_excel)} linhas")

    # Etapa 3: Criar DataFrames e agrupar com tolerância
    print("\nETAPA 3: Agrupando com tolerância de 1 centavo...")

    df_xml = pd.DataFrame(todos_itens)
    df_xml['NR_SEQ_PROTOCOLO'] = df_xml['NR_SEQ_PROTOCOLO'].astype(str)
    df_xml['NR_INTERNO_CONTA'] = df_xml['NR_INTERNO_CONTA'].astype(str)
    df_xml['ITEM_CD_CONVENIO'] = df_xml['ITEM_CD_CONVENIO'].astype(str)

    # Agrupar Excel por conta + item (somando todas as quantidades, independente do preço)
    df_excel_agrupado = df_excel.groupby(
        ['NR_SEQ_PROTOCOLO', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO'],
        as_index=False
    ).agg({
        'QT_ITEM': 'sum',
        'PRECO_UNITARIO': 'mean',
        'PRECO_TOTAL': 'sum',
        'DS_ITEM': 'first'
    })

    # Agrupar XML por conta + item
    df_xml_agrupado = df_xml.groupby(
        ['NR_SEQ_PROTOCOLO', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO'],
        as_index=False
    ).agg({
        'QT_ITEM': 'sum',
        'PRECO_UNITARIO': 'mean',
        'PRECO_TOTAL': 'sum',
        'ARQUIVO_XML': lambda x: ', '.join(set(x))
    })

    print(f"  Excel agrupado: {len(df_excel_agrupado)} itens")
    print(f"  XML agrupado: {len(df_xml_agrupado)} itens")

    # Etapa 4: Comparar
    print("\nETAPA 4: Comparando dados...")

    df_comparacao = pd.merge(
        df_excel_agrupado,
        df_xml_agrupado,
        on=['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO'],
        how='outer',
        suffixes=('_EXCEL', '_XML')
    )

    # Diferenças de quantidade
    df_dif_qtd = df_comparacao[
        (df_comparacao['QT_ITEM_EXCEL'].notna()) &
        (df_comparacao['QT_ITEM_XML'].notna()) &
        (abs(df_comparacao['QT_ITEM_EXCEL'] - df_comparacao['QT_ITEM_XML']) > 0.001)
    ].copy()

    if len(df_dif_qtd) > 0:
        df_dif_qtd['DIFERENCA_QTD'] = df_dif_qtd['QT_ITEM_EXCEL'] - df_dif_qtd['QT_ITEM_XML']
        df_dif_qtd = df_dif_qtd[[
            'NR_SEQ_PROTOCOLO_EXCEL', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'DS_ITEM',
            'QT_ITEM_EXCEL', 'QT_ITEM_XML', 'DIFERENCA_QTD',
            'PRECO_UNITARIO_EXCEL', 'PRECO_UNITARIO_XML', 'ARQUIVO_XML'
        ]].rename(columns={'NR_SEQ_PROTOCOLO_EXCEL': 'NR_SEQ_PROTOCOLO'})

    # Diferenças de preço (> 1 centavo)
    df_dif_preco = df_comparacao[
        (df_comparacao['PRECO_UNITARIO_EXCEL'].notna()) &
        (df_comparacao['PRECO_UNITARIO_XML'].notna()) &
        (abs(df_comparacao['PRECO_UNITARIO_EXCEL'] - df_comparacao['PRECO_UNITARIO_XML']) > TOLERANCIA_PRECO)
    ].copy()

    if len(df_dif_preco) > 0:
        df_dif_preco['DIFERENCA_PRECO'] = df_dif_preco['PRECO_UNITARIO_EXCEL'] - df_dif_preco['PRECO_UNITARIO_XML']
        df_dif_preco = df_dif_preco[[
            'NR_SEQ_PROTOCOLO_EXCEL', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'DS_ITEM',
            'PRECO_UNITARIO_EXCEL', 'PRECO_UNITARIO_XML', 'DIFERENCA_PRECO',
            'QT_ITEM_EXCEL', 'QT_ITEM_XML', 'ARQUIVO_XML'
        ]].rename(columns={'NR_SEQ_PROTOCOLO_EXCEL': 'NR_SEQ_PROTOCOLO'})

    # Itens apenas no Excel
    df_apenas_excel = df_comparacao[df_comparacao['QT_ITEM_XML'].isna()].copy()
    if len(df_apenas_excel) > 0:
        df_apenas_excel = df_apenas_excel[[
            'NR_SEQ_PROTOCOLO_EXCEL', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'DS_ITEM',
            'QT_ITEM_EXCEL', 'PRECO_UNITARIO_EXCEL', 'PRECO_TOTAL_EXCEL'
        ]].rename(columns={
            'NR_SEQ_PROTOCOLO_EXCEL': 'NR_SEQ_PROTOCOLO',
            'QT_ITEM_EXCEL': 'QT_ITEM',
            'PRECO_UNITARIO_EXCEL': 'PRECO_UNITARIO',
            'PRECO_TOTAL_EXCEL': 'PRECO_TOTAL'
        })

    # Itens apenas no XML
    df_apenas_xml = df_comparacao[df_comparacao['QT_ITEM_EXCEL'].isna()].copy()
    if len(df_apenas_xml) > 0:
        df_apenas_xml = df_apenas_xml[[
            'NR_SEQ_PROTOCOLO_XML', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO',
            'QT_ITEM_XML', 'PRECO_UNITARIO_XML', 'PRECO_TOTAL_XML', 'ARQUIVO_XML'
        ]].rename(columns={
            'NR_SEQ_PROTOCOLO_XML': 'NR_SEQ_PROTOCOLO',
            'QT_ITEM_XML': 'QT_ITEM',
            'PRECO_UNITARIO_XML': 'PRECO_UNITARIO',
            'PRECO_TOTAL_XML': 'PRECO_TOTAL'
        })

    # Resumo
    print(f"\n  RESUMO (com tolerância de 1 centavo):")
    print(f"  - Itens com diferença de quantidade: {len(df_dif_qtd)}")
    print(f"  - Itens com diferença de preço > 1 centavo: {len(df_dif_preco)}")
    print(f"  - Itens apenas no Excel: {len(df_apenas_excel)}")
    print(f"  - Itens apenas no XML: {len(df_apenas_xml)}")

    # Etapa 5: Gerar relatório
    print("\nETAPA 5: Gerando relatório Excel...")

    with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:
        if len(df_dif_qtd) > 0:
            df_dif_qtd.to_excel(writer, sheet_name='1-Diferenca Quantidade', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhuma diferença encontrada']}).to_excel(
                writer, sheet_name='1-Diferenca Quantidade', index=False)

        if len(df_dif_preco) > 0:
            df_dif_preco.to_excel(writer, sheet_name='2-Diferenca Preco >1cent', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhuma diferença encontrada']}).to_excel(
                writer, sheet_name='2-Diferenca Preco >1cent', index=False)

        if len(df_apenas_excel) > 0:
            df_apenas_excel.to_excel(writer, sheet_name='3-Apenas Excel', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhum item exclusivo']}).to_excel(
                writer, sheet_name='3-Apenas Excel', index=False)

        if len(df_apenas_xml) > 0:
            df_apenas_xml.to_excel(writer, sheet_name='4-Apenas XML', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhum item exclusivo']}).to_excel(
                writer, sheet_name='4-Apenas XML', index=False)

    print(f"\n  Relatório salvo em: {ARQUIVO_SAIDA}")
    print("\n" + "=" * 70)
    print("PROCESSAMENTO CONCLUÍDO!")
    print("=" * 70)


if __name__ == "__main__":
    main()
