# -*- coding: utf-8 -*-
"""
Script para comparar contas médicas entre Excel (TASY) e arquivos XML (TISS)
"""

import os
import re
import pandas as pd
from pathlib import Path
from collections import defaultdict
import xml.etree.ElementTree as ET
from decimal import Decimal, ROUND_HALF_UP
import warnings
warnings.filterwarnings('ignore')

# Configurações
PASTA_XML = r"C:\Users\AMH\Desktop\meu-site\xml"
ARQUIVO_EXCEL = r"C:\Users\AMH\Desktop\meu-site\Unimed conta recalculadas.xlsx"
ARQUIVO_SAIDA = r"C:\Users\AMH\Desktop\meu-site\Relatorio_Comparacao.xlsx"
CODIGO_PRESTADOR_VALIDO = "110020"
CONTAS_IGNORAR = {74078, 75059, 60282}

# Namespace TISS
NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}


def extrair_texto(elemento, xpath):
    """Extrai texto de um elemento XML usando xpath"""
    el = elemento.find(xpath, NS)
    return el.text.strip() if el is not None and el.text else None


def processar_procedimento(proc, numero_lote, numero_guia, arquivo_xml):
    """Processa um procedimento executado e retorna dict com dados"""
    # Verificar se o código do prestador é válido (110020)
    cod_prestador = extrair_texto(proc, './/ans:codigoPrestadorNaOperadora')

    codigo = extrair_texto(proc, './/ans:codigoProcedimento')
    if not codigo:
        return None

    qtd_str = extrair_texto(proc, './/ans:quantidadeExecutada')
    valor_unit_str = extrair_texto(proc, './/ans:valorUnitario')
    valor_total_str = extrair_texto(proc, './/ans:valorTotal')

    if not all([qtd_str, valor_unit_str]):
        return None

    # Converter quantidade (pode vir como 1.0000 ou 1)
    try:
        qtd = float(qtd_str)
    except:
        return None

    try:
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
    """Processa uma guia SP-SADT e extrai todos os procedimentos"""
    itens = []

    numero_guia = extrair_texto(guia, './/ans:numeroGuiaPrestador')
    if not numero_guia:
        return itens

    # Ignorar contas específicas
    try:
        if int(numero_guia) in CONTAS_IGNORAR:
            return itens
    except:
        pass

    # Processar procedimentos executados
    for proc in guia.findall('.//ans:procedimentoExecutado', NS):
        item = processar_procedimento(proc, numero_lote, numero_guia, arquivo_xml)
        if item:
            itens.append(item)

    # Processar serviços executados (dentro de outrasDespesas)
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

        # Verificar código do prestador no cabeçalho
        cod_prestador_cabecalho = extrair_texto(root, './/ans:cabecalho//ans:codigoPrestadorNaOperadora')

        # Extrair número do lote (protocolo)
        numero_lote = extrair_texto(root, './/ans:numeroLote')

        # Processar todas as guias SP-SADT
        for guia in root.findall('.//ans:guiaSP-SADT', NS):
            itens_guia = processar_guia(guia, numero_lote, nome_arquivo)
            itens.extend(itens_guia)

        # Processar guias de consulta
        for guia in root.findall('.//ans:guiaConsulta', NS):
            itens_guia = processar_guia(guia, numero_lote, nome_arquivo)
            itens.extend(itens_guia)

        # Processar guias de internação (ADICIONADO)
        for guia in root.findall('.//ans:guiaResumoInternacao', NS):
            itens_guia = processar_guia(guia, numero_lote, nome_arquivo)
            itens.extend(itens_guia)

    except ET.ParseError as e:
        print(f"Erro ao processar {nome_arquivo}: {e}")
    except Exception as e:
        print(f"Erro inesperado em {nome_arquivo}: {e}")

    return itens


def extrair_dados_xmls():
    """Extrai dados de todos os arquivos XML"""
    print("=" * 60)
    print("ETAPA 1: Extraindo dados dos arquivos XML")
    print("=" * 60)

    todos_itens = []
    arquivos_processados = 0
    protocolos_xml = set()
    contas_xml = set()
    arquivos_por_protocolo = defaultdict(list)

    # Percorrer todos os XMLs
    for root_dir, dirs, files in os.walk(PASTA_XML):
        for file in files:
            if file.endswith('.xml'):
                caminho = os.path.join(root_dir, file)
                itens = processar_arquivo_xml(caminho)

                for item in itens:
                    # Filtrar por código do prestador válido
                    if item.get('COD_PRESTADOR') and item['COD_PRESTADOR'] != CODIGO_PRESTADOR_VALIDO:
                        continue

                    todos_itens.append(item)
                    if item['NR_SEQ_PROTOCOLO']:
                        protocolos_xml.add(item['NR_SEQ_PROTOCOLO'])
                        arquivos_por_protocolo[item['NR_SEQ_PROTOCOLO']].append(file)
                    if item['NR_INTERNO_CONTA']:
                        contas_xml.add(item['NR_INTERNO_CONTA'])

                arquivos_processados += 1
                if arquivos_processados % 100 == 0:
                    print(f"  Processados {arquivos_processados} arquivos...")

    print(f"\n  Total de arquivos XML processados: {arquivos_processados}")
    print(f"  Total de itens extraídos: {len(todos_itens)}")
    print(f"  Protocolos únicos encontrados: {len(protocolos_xml)}")
    print(f"  Contas únicas encontradas: {len(contas_xml)}")

    # Identificar protocolos duplicados (mais de um arquivo XML)
    protocolos_duplicados = {p: arquivos for p, arquivos in arquivos_por_protocolo.items()
                            if len(set(arquivos)) > 1}

    return todos_itens, protocolos_xml, contas_xml, arquivos_por_protocolo, protocolos_duplicados


def processar_excel():
    """Processa o arquivo Excel e aplica regras de agrupamento"""
    print("\n" + "=" * 60)
    print("ETAPA 2: Processando arquivo Excel")
    print("=" * 60)

    df = pd.read_excel(ARQUIVO_EXCEL)
    print(f"  Linhas no Excel: {len(df)}")

    # Filtrar contas a ignorar
    df = df[~df['NR_INTERNO_CONTA'].isin(CONTAS_IGNORAR)]
    print(f"  Linhas após filtrar contas ignoradas: {len(df)}")

    # Converter tipos
    df['NR_SEQ_PROTOCOLO'] = df['NR_SEQ_PROTOCOLO'].astype(str)
    df['NR_INTERNO_CONTA'] = df['NR_INTERNO_CONTA'].astype(str)
    df['ITEM_CD_CONVENIO'] = df['ITEM_CD_CONVENIO'].astype(str)
    df['QT_ITEM'] = pd.to_numeric(df['QT_ITEM'], errors='coerce').fillna(0)
    df['PRECO_UNITARIO'] = pd.to_numeric(df['PRECO_UNITARIO'], errors='coerce').fillna(0)
    df['PRECO_TOTAL'] = pd.to_numeric(df['PRECO_TOTAL'], errors='coerce').fillna(0)

    # Agrupar itens iguais (mesmo código + conta + preço unitário) e somar quantidades
    df_agrupado = df.groupby(
        ['NR_SEQ_PROTOCOLO', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'PRECO_UNITARIO'],
        as_index=False
    ).agg({
        'QT_ITEM': 'sum',
        'PRECO_TOTAL': 'sum',
        'DS_ITEM': 'first'
    })

    print(f"  Linhas após agrupamento: {len(df_agrupado)}")

    protocolos_excel = set(df_agrupado['NR_SEQ_PROTOCOLO'].unique())
    contas_excel = set(df_agrupado['NR_INTERNO_CONTA'].unique())

    print(f"  Protocolos únicos: {len(protocolos_excel)}")
    print(f"  Contas únicas: {len(contas_excel)}")

    return df_agrupado, df, protocolos_excel, contas_excel


def comparar_dados(df_excel, itens_xml, protocolos_excel, protocolos_xml,
                   contas_excel, contas_xml, arquivos_por_protocolo, protocolos_duplicados):
    """Compara dados do Excel com XML e identifica divergências"""
    print("\n" + "=" * 60)
    print("ETAPA 3: Comparando dados Excel vs XML")
    print("=" * 60)

    # Criar DataFrame dos XMLs
    df_xml = pd.DataFrame(itens_xml)

    if len(df_xml) == 0:
        print("  AVISO: Nenhum item válido encontrado nos XMLs!")
        return None, None, None, None, None, None

    # Converter tipos no XML
    df_xml['NR_SEQ_PROTOCOLO'] = df_xml['NR_SEQ_PROTOCOLO'].astype(str)
    df_xml['NR_INTERNO_CONTA'] = df_xml['NR_INTERNO_CONTA'].astype(str)
    df_xml['ITEM_CD_CONVENIO'] = df_xml['ITEM_CD_CONVENIO'].astype(str)

    # Agrupar itens do XML (mesmo código + conta + preço unitário)
    df_xml_agrupado = df_xml.groupby(
        ['NR_SEQ_PROTOCOLO', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'PRECO_UNITARIO'],
        as_index=False
    ).agg({
        'QT_ITEM': 'sum',
        'PRECO_TOTAL': 'sum',
        'ARQUIVO_XML': lambda x: ', '.join(set(x))
    })

    print(f"  Itens agrupados no XML: {len(df_xml_agrupado)}")

    # Chave de comparação
    df_excel['CHAVE'] = (df_excel['NR_INTERNO_CONTA'] + '_' +
                         df_excel['ITEM_CD_CONVENIO'] + '_' +
                         df_excel['PRECO_UNITARIO'].round(2).astype(str))

    df_xml_agrupado['CHAVE'] = (df_xml_agrupado['NR_INTERNO_CONTA'] + '_' +
                                 df_xml_agrupado['ITEM_CD_CONVENIO'] + '_' +
                                 df_xml_agrupado['PRECO_UNITARIO'].round(2).astype(str))

    # 1. Resumo de Protocolos
    print("\n  Analisando protocolos...")
    resumo_protocolos = []
    todos_protocolos = protocolos_excel | protocolos_xml

    for protocolo in todos_protocolos:
        tem_excel = protocolo in protocolos_excel
        tem_xml = protocolo in protocolos_xml
        duplicado = protocolo in protocolos_duplicados
        arquivos = ', '.join(set(arquivos_por_protocolo.get(protocolo, [])))

        resumo_protocolos.append({
            'NR_SEQ_PROTOCOLO': protocolo,
            'NO_EXCEL': 'Sim' if tem_excel else 'Não',
            'NO_XML': 'Sim' if tem_xml else 'Não',
            'XML_DUPLICADO': 'Sim' if duplicado else 'Não',
            'ARQUIVOS_XML': arquivos,
            'STATUS': 'OK' if tem_excel and tem_xml and not duplicado else
                      'DUPLICADO' if duplicado else
                      'APENAS EXCEL' if tem_excel and not tem_xml else
                      'APENAS XML'
        })

    df_resumo_protocolos = pd.DataFrame(resumo_protocolos)

    # 2. Resumo de Contas
    print("  Analisando contas...")
    resumo_contas = []
    todas_contas = contas_excel | contas_xml

    contas_por_arquivo = defaultdict(set)
    for item in itens_xml:
        if item['NR_INTERNO_CONTA']:
            contas_por_arquivo[item['NR_INTERNO_CONTA']].add(item['ARQUIVO_XML'])

    for conta in todas_contas:
        tem_excel = conta in contas_excel
        tem_xml = conta in contas_xml
        arquivos = contas_por_arquivo.get(conta, set())
        duplicado = len(arquivos) > 1

        resumo_contas.append({
            'NR_INTERNO_CONTA': conta,
            'NO_EXCEL': 'Sim' if tem_excel else 'Não',
            'NO_XML': 'Sim' if tem_xml else 'Não',
            'XML_DUPLICADO': 'Sim' if duplicado else 'Não',
            'ARQUIVOS_XML': ', '.join(arquivos),
            'STATUS': 'OK' if tem_excel and tem_xml and not duplicado else
                      'DUPLICADO' if duplicado else
                      'APENAS EXCEL' if tem_excel and not tem_xml else
                      'APENAS XML'
        })

    df_resumo_contas = pd.DataFrame(resumo_contas)

    # 3 e 4. Comparar itens - encontrar diferenças de quantidade e preço
    print("  Comparando itens...")

    # Arredondar preços para evitar problemas de precisão de ponto flutuante
    df_excel['PRECO_UNITARIO_ROUND'] = df_excel['PRECO_UNITARIO'].round(2)
    df_xml_agrupado['PRECO_UNITARIO_ROUND'] = df_xml_agrupado['PRECO_UNITARIO'].round(2)

    # Merge para comparação de QUANTIDADE (incluindo preço na chave!)
    df_comparacao = pd.merge(
        df_excel,
        df_xml_agrupado,
        on=['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'PRECO_UNITARIO_ROUND'],
        how='outer',
        suffixes=('_EXCEL', '_XML')
    )

    # Itens com diferença de quantidade (mesmo item, mesma conta, mesmo preço)
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
            'PRECO_UNITARIO_ROUND', 'ARQUIVO_XML'
        ]].rename(columns={'NR_SEQ_PROTOCOLO_EXCEL': 'NR_SEQ_PROTOCOLO',
                          'PRECO_UNITARIO_ROUND': 'PRECO_UNITARIO'})

    # Itens com diferença de preço unitário
    # (mesmo código e conta existem em ambos, mas com preços diferentes que não casam)
    # Primeiro, encontrar combinações conta+item que existem em ambos
    excel_conta_item = df_excel.groupby(['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO'])['PRECO_UNITARIO_ROUND'].apply(set).reset_index()
    xml_conta_item = df_xml_agrupado.groupby(['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO'])['PRECO_UNITARIO_ROUND'].apply(set).reset_index()

    merged_precos = pd.merge(excel_conta_item, xml_conta_item,
                             on=['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO'],
                             suffixes=('_EXCEL', '_XML'))

    # Encontrar onde os conjuntos de preços são diferentes
    merged_precos['PRECOS_IGUAIS'] = merged_precos.apply(
        lambda row: row['PRECO_UNITARIO_ROUND_EXCEL'] == row['PRECO_UNITARIO_ROUND_XML'], axis=1)

    dif_precos_contas = merged_precos[~merged_precos['PRECOS_IGUAIS']][['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO']]

    # Pegar detalhes dos itens com diferença de preço
    if len(dif_precos_contas) > 0:
        df_dif_preco = pd.merge(
            df_excel[['NR_SEQ_PROTOCOLO', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'DS_ITEM',
                      'PRECO_UNITARIO', 'QT_ITEM', 'PRECO_TOTAL']],
            dif_precos_contas,
            on=['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO']
        )
        df_dif_preco = df_dif_preco.drop_duplicates()
    else:
        df_dif_preco = pd.DataFrame()

    # 5. Itens do Excel que não estão no XML
    print("  Identificando itens exclusivos do Excel...")
    chaves_xml = set(df_xml_agrupado['CHAVE'])
    df_apenas_excel = df_excel[~df_excel['CHAVE'].isin(chaves_xml)].copy()
    df_apenas_excel = df_apenas_excel.drop(columns=['CHAVE'])

    # 6. Itens do XML que não estão no Excel
    print("  Identificando itens exclusivos do XML...")
    chaves_excel = set(df_excel['CHAVE'])
    df_apenas_xml = df_xml_agrupado[~df_xml_agrupado['CHAVE'].isin(chaves_excel)].copy()
    df_apenas_xml = df_apenas_xml.drop(columns=['CHAVE'])

    # Estatísticas
    print(f"\n  RESUMO:")
    print(f"  - Protocolos apenas no Excel: {len(df_resumo_protocolos[df_resumo_protocolos['STATUS'] == 'APENAS EXCEL'])}")
    print(f"  - Protocolos apenas no XML: {len(df_resumo_protocolos[df_resumo_protocolos['STATUS'] == 'APENAS XML'])}")
    print(f"  - Protocolos duplicados (múltiplos XMLs): {len(df_resumo_protocolos[df_resumo_protocolos['STATUS'] == 'DUPLICADO'])}")
    print(f"  - Contas apenas no Excel: {len(df_resumo_contas[df_resumo_contas['STATUS'] == 'APENAS EXCEL'])}")
    print(f"  - Contas apenas no XML: {len(df_resumo_contas[df_resumo_contas['STATUS'] == 'APENAS XML'])}")
    print(f"  - Itens com diferença de quantidade: {len(df_dif_qtd)}")
    print(f"  - Itens com diferença de preço: {len(df_dif_preco)}")
    print(f"  - Itens apenas no Excel: {len(df_apenas_excel)}")
    print(f"  - Itens apenas no XML: {len(df_apenas_xml)}")

    return df_resumo_protocolos, df_resumo_contas, df_dif_qtd, df_dif_preco, df_apenas_excel, df_apenas_xml


def gerar_relatorio(df_resumo_protocolos, df_resumo_contas, df_dif_qtd,
                    df_dif_preco, df_apenas_excel, df_apenas_xml):
    """Gera o relatório Excel final com todas as abas"""
    print("\n" + "=" * 60)
    print("ETAPA 4: Gerando relatório Excel")
    print("=" * 60)

    with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:
        # Aba 1: Resumo de Protocolos
        df_resumo_protocolos.to_excel(writer, sheet_name='1-Resumo Protocolos', index=False)
        print(f"  Aba 1 criada: Resumo Protocolos ({len(df_resumo_protocolos)} linhas)")

        # Aba 2: Resumo de Contas
        df_resumo_contas.to_excel(writer, sheet_name='2-Resumo Contas', index=False)
        print(f"  Aba 2 criada: Resumo Contas ({len(df_resumo_contas)} linhas)")

        # Aba 3: Diferenças de Quantidade
        if df_dif_qtd is not None and len(df_dif_qtd) > 0:
            df_dif_qtd.to_excel(writer, sheet_name='3-Diferenca Quantidade', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhuma diferença de quantidade encontrada']}).to_excel(
                writer, sheet_name='3-Diferenca Quantidade', index=False)
        print(f"  Aba 3 criada: Diferença Quantidade ({len(df_dif_qtd) if df_dif_qtd is not None else 0} linhas)")

        # Aba 4: Diferenças de Preço
        if df_dif_preco is not None and len(df_dif_preco) > 0:
            df_dif_preco.to_excel(writer, sheet_name='4-Diferenca Preco', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhuma diferença de preço encontrada']}).to_excel(
                writer, sheet_name='4-Diferenca Preco', index=False)
        print(f"  Aba 4 criada: Diferença Preço ({len(df_dif_preco) if df_dif_preco is not None else 0} linhas)")

        # Aba 5: Itens apenas no Excel
        if df_apenas_excel is not None and len(df_apenas_excel) > 0:
            df_apenas_excel.to_excel(writer, sheet_name='5-Apenas Excel', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhum item exclusivo do Excel']}).to_excel(
                writer, sheet_name='5-Apenas Excel', index=False)
        print(f"  Aba 5 criada: Apenas Excel ({len(df_apenas_excel) if df_apenas_excel is not None else 0} linhas)")

        # Aba 6: Itens apenas no XML
        if df_apenas_xml is not None and len(df_apenas_xml) > 0:
            df_apenas_xml.to_excel(writer, sheet_name='6-Apenas XML', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhum item exclusivo do XML']}).to_excel(
                writer, sheet_name='6-Apenas XML', index=False)
        print(f"  Aba 6 criada: Apenas XML ({len(df_apenas_xml) if df_apenas_xml is not None else 0} linhas)")

    print(f"\n  Relatório salvo em: {ARQUIVO_SAIDA}")


def main():
    """Função principal"""
    print("\n" + "=" * 60)
    print("COMPARAÇÃO DE CONTAS MÉDICAS - EXCEL vs XML")
    print("=" * 60)

    # Etapa 1: Extrair dados dos XMLs
    itens_xml, protocolos_xml, contas_xml, arquivos_por_protocolo, protocolos_duplicados = extrair_dados_xmls()

    # Etapa 2: Processar Excel
    df_excel_agrupado, df_excel_original, protocolos_excel, contas_excel = processar_excel()

    # Etapa 3: Comparar dados
    resultados = comparar_dados(
        df_excel_agrupado, itens_xml,
        protocolos_excel, protocolos_xml,
        contas_excel, contas_xml,
        arquivos_por_protocolo, protocolos_duplicados
    )

    if resultados[0] is None:
        print("\nErro: Não foi possível realizar a comparação.")
        return

    # Etapa 4: Gerar relatório
    gerar_relatorio(*resultados)

    print("\n" + "=" * 60)
    print("PROCESSAMENTO CONCLUÍDO!")
    print("=" * 60)


if __name__ == "__main__":
    main()
