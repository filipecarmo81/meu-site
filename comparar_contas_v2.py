# -*- coding: utf-8 -*-
"""
Script para comparar contas médicas entre Excel (TASY) e arquivos XML (TISS)
VERSÃO 2: Com tolerância de 1 centavo para arredondamento de preços
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
ARQUIVO_SAIDA = r"C:\Users\AMH\Desktop\meu-site\Relatorio_Comparacao_v2.xlsx"
CODIGO_PRESTADOR_VALIDO = "110020"
CONTAS_IGNORAR = {74078, 75059, 60282}
TOLERANCIA_PRECO = 0.01  # Tolerância de 1 centavo

# Namespace TISS
NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}


def extrair_texto(elemento, xpath):
    """Extrai texto de um elemento XML usando xpath"""
    el = elemento.find(xpath, NS)
    return el.text.strip() if el is not None and el.text else None


def arredondar_com_tolerancia(preco):
    """
    Arredonda preço para agrupar valores com diferença de 1 centavo.
    Exemplo: 0.25 e 0.26 viram o mesmo valor (0.26)
    Arredonda para cima para o próximo múltiplo de 0.02
    """
    import math
    return round(math.ceil(preco / 0.02) * 0.02, 2)


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


def extrair_dados_xmls():
    """Extrai dados de todos os arquivos XML"""
    print("=" * 70)
    print("ETAPA 1: Extraindo dados dos arquivos XML")
    print("=" * 70)

    todos_itens = []
    arquivos_processados = 0
    protocolos_xml = set()
    contas_xml = set()
    arquivos_por_protocolo = defaultdict(list)
    protocolo_por_conta = {}  # Para guardar qual protocolo cada conta pertence

    for root_dir, dirs, files in os.walk(PASTA_XML):
        for file in files:
            if file.endswith('.xml'):
                caminho = os.path.join(root_dir, file)
                itens = processar_arquivo_xml(caminho)

                for item in itens:
                    if item.get('COD_PRESTADOR') and item['COD_PRESTADOR'] != CODIGO_PRESTADOR_VALIDO:
                        continue

                    todos_itens.append(item)
                    if item['NR_SEQ_PROTOCOLO']:
                        protocolos_xml.add(item['NR_SEQ_PROTOCOLO'])
                        arquivos_por_protocolo[item['NR_SEQ_PROTOCOLO']].append(file)
                    if item['NR_INTERNO_CONTA']:
                        contas_xml.add(item['NR_INTERNO_CONTA'])
                        # Guardar o protocolo da conta
                        if item['NR_INTERNO_CONTA'] not in protocolo_por_conta:
                            protocolo_por_conta[item['NR_INTERNO_CONTA']] = item['NR_SEQ_PROTOCOLO']

                arquivos_processados += 1
                if arquivos_processados % 200 == 0:
                    print(f"  Processados {arquivos_processados} arquivos...")

    print(f"\n  Total de arquivos XML processados: {arquivos_processados}")
    print(f"  Total de itens extraidos: {len(todos_itens)}")
    print(f"  Protocolos unicos encontrados: {len(protocolos_xml)}")
    print(f"  Contas unicas encontradas: {len(contas_xml)}")

    protocolos_duplicados = {p: arquivos for p, arquivos in arquivos_por_protocolo.items()
                            if len(set(arquivos)) > 1}

    return todos_itens, protocolos_xml, contas_xml, arquivos_por_protocolo, protocolos_duplicados, protocolo_por_conta


def processar_excel():
    """Processa o arquivo Excel e aplica regras de agrupamento"""
    print("\n" + "=" * 70)
    print("ETAPA 2: Processando arquivo Excel")
    print("=" * 70)

    df = pd.read_excel(ARQUIVO_EXCEL)
    print(f"  Linhas no Excel: {len(df)}")

    df = df[~df['NR_INTERNO_CONTA'].isin(CONTAS_IGNORAR)]
    print(f"  Linhas apos filtrar contas ignoradas: {len(df)}")

    df['NR_SEQ_PROTOCOLO'] = df['NR_SEQ_PROTOCOLO'].astype(str)
    df['NR_INTERNO_CONTA'] = df['NR_INTERNO_CONTA'].astype(str)
    df['ITEM_CD_CONVENIO'] = df['ITEM_CD_CONVENIO'].astype(str)
    df['QT_ITEM'] = pd.to_numeric(df['QT_ITEM'], errors='coerce').fillna(0)
    df['PRECO_UNITARIO'] = pd.to_numeric(df['PRECO_UNITARIO'], errors='coerce').fillna(0)
    df['PRECO_TOTAL'] = pd.to_numeric(df['PRECO_TOTAL'], errors='coerce').fillna(0)

    # Criar preço arredondado com tolerância
    df['PRECO_TOLERANCIA'] = df['PRECO_UNITARIO'].apply(arredondar_com_tolerancia)

    # Agrupar por conta + código + preço (com tolerância)
    df_agrupado = df.groupby(
        ['NR_SEQ_PROTOCOLO', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'PRECO_TOLERANCIA'],
        as_index=False
    ).agg({
        'QT_ITEM': 'sum',
        'PRECO_TOTAL': 'sum',
        'PRECO_UNITARIO': 'mean',  # Média dos preços originais
        'DS_ITEM': 'first'
    })

    print(f"  Linhas apos agrupamento: {len(df_agrupado)}")

    protocolos_excel = set(df_agrupado['NR_SEQ_PROTOCOLO'].unique())
    contas_excel = set(df_agrupado['NR_INTERNO_CONTA'].unique())

    # Mapear protocolo por conta do Excel
    protocolo_por_conta_excel = df.groupby('NR_INTERNO_CONTA')['NR_SEQ_PROTOCOLO'].first().to_dict()

    print(f"  Protocolos unicos: {len(protocolos_excel)}")
    print(f"  Contas unicas: {len(contas_excel)}")

    return df_agrupado, df, protocolos_excel, contas_excel, protocolo_por_conta_excel


def comparar_dados(df_excel, itens_xml, protocolos_excel, protocolos_xml,
                   contas_excel, contas_xml, arquivos_por_protocolo,
                   protocolos_duplicados, protocolo_por_conta_xml, protocolo_por_conta_excel):
    """Compara dados do Excel com XML e identifica divergências"""
    print("\n" + "=" * 70)
    print("ETAPA 3: Comparando dados Excel vs XML")
    print("=" * 70)

    # Criar DataFrame dos XMLs
    df_xml = pd.DataFrame(itens_xml)

    if len(df_xml) == 0:
        print("  AVISO: Nenhum item valido encontrado nos XMLs!")
        return None, None, None, None, None, None

    df_xml['NR_SEQ_PROTOCOLO'] = df_xml['NR_SEQ_PROTOCOLO'].astype(str)
    df_xml['NR_INTERNO_CONTA'] = df_xml['NR_INTERNO_CONTA'].astype(str)
    df_xml['ITEM_CD_CONVENIO'] = df_xml['ITEM_CD_CONVENIO'].astype(str)

    # Criar preço arredondado com tolerância
    df_xml['PRECO_TOLERANCIA'] = df_xml['PRECO_UNITARIO'].apply(arredondar_com_tolerancia)

    # Agrupar XML por conta + código + preço (com tolerância)
    df_xml_agrupado = df_xml.groupby(
        ['NR_SEQ_PROTOCOLO', 'NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'PRECO_TOLERANCIA'],
        as_index=False
    ).agg({
        'QT_ITEM': 'sum',
        'PRECO_TOTAL': 'sum',
        'PRECO_UNITARIO': 'mean',
        'ARQUIVO_XML': lambda x: ', '.join(set(x))
    })

    print(f"  Itens agrupados no XML: {len(df_xml_agrupado)}")

    # =====================================================
    # ABA 1: Resumo de Protocolos
    # =====================================================
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
            'NO_EXCEL': 'Sim' if tem_excel else 'Nao',
            'NO_XML': 'Sim' if tem_xml else 'Nao',
            'XML_DUPLICADO': 'Sim' if duplicado else 'Nao',
            'ARQUIVOS_XML': arquivos,
            'STATUS': 'OK' if tem_excel and tem_xml and not duplicado else
                      'DUPLICADO' if duplicado else
                      'APENAS EXCEL' if tem_excel and not tem_xml else
                      'APENAS XML'
        })

    df_resumo_protocolos = pd.DataFrame(resumo_protocolos)

    # =====================================================
    # ABA 2: Resumo de Contas (com número do protocolo)
    # =====================================================
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

        # Pegar o número do protocolo
        protocolo = protocolo_por_conta_excel.get(conta) or protocolo_por_conta_xml.get(conta) or ''

        resumo_contas.append({
            'NR_SEQ_PROTOCOLO': protocolo,
            'NR_INTERNO_CONTA': conta,
            'NO_EXCEL': 'Sim' if tem_excel else 'Nao',
            'NO_XML': 'Sim' if tem_xml else 'Nao',
            'XML_DUPLICADO': 'Sim' if duplicado else 'Nao',
            'ARQUIVOS_XML': ', '.join(arquivos),
            'STATUS': 'OK' if tem_excel and tem_xml and not duplicado else
                      'DUPLICADO' if duplicado else
                      'APENAS EXCEL' if tem_excel and not tem_xml else
                      'APENAS XML'
        })

    df_resumo_contas = pd.DataFrame(resumo_contas)

    # =====================================================
    # ABA 3 e 4: Comparar itens
    # =====================================================
    print("  Comparando itens...")

    # Chave de comparação (com tolerância de preço)
    df_excel['CHAVE'] = (df_excel['NR_INTERNO_CONTA'] + '_' +
                         df_excel['ITEM_CD_CONVENIO'] + '_' +
                         df_excel['PRECO_TOLERANCIA'].astype(str))

    df_xml_agrupado['CHAVE'] = (df_xml_agrupado['NR_INTERNO_CONTA'] + '_' +
                                 df_xml_agrupado['ITEM_CD_CONVENIO'] + '_' +
                                 df_xml_agrupado['PRECO_TOLERANCIA'].astype(str))

    # Merge para comparação de quantidade
    df_comparacao = pd.merge(
        df_excel,
        df_xml_agrupado,
        on=['NR_INTERNO_CONTA', 'ITEM_CD_CONVENIO', 'PRECO_TOLERANCIA'],
        how='outer',
        suffixes=('_EXCEL', '_XML')
    )

    # ABA 3: Diferença de quantidade
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

    # ABA 4: Diferença de preço (> 1 centavo)
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

    # =====================================================
    # ABA 5: Itens apenas no Excel
    # =====================================================
    print("  Identificando itens exclusivos do Excel...")
    chaves_xml = set(df_xml_agrupado['CHAVE'])
    df_apenas_excel = df_excel[~df_excel['CHAVE'].isin(chaves_xml)].copy()
    df_apenas_excel = df_apenas_excel.drop(columns=['CHAVE', 'PRECO_TOLERANCIA'])

    # =====================================================
    # ABA 6: Itens apenas no XML
    # =====================================================
    print("  Identificando itens exclusivos do XML...")
    chaves_excel = set(df_excel['CHAVE'])
    df_apenas_xml = df_xml_agrupado[~df_xml_agrupado['CHAVE'].isin(chaves_excel)].copy()
    df_apenas_xml = df_apenas_xml.drop(columns=['CHAVE', 'PRECO_TOLERANCIA'])

    # Estatísticas
    print(f"\n  RESUMO:")
    print(f"  - Protocolos apenas no Excel: {len(df_resumo_protocolos[df_resumo_protocolos['STATUS'] == 'APENAS EXCEL'])}")
    print(f"  - Protocolos apenas no XML: {len(df_resumo_protocolos[df_resumo_protocolos['STATUS'] == 'APENAS XML'])}")
    print(f"  - Protocolos duplicados: {len(df_resumo_protocolos[df_resumo_protocolos['STATUS'] == 'DUPLICADO'])}")
    print(f"  - Contas apenas no Excel: {len(df_resumo_contas[df_resumo_contas['STATUS'] == 'APENAS EXCEL'])}")
    print(f"  - Contas apenas no XML: {len(df_resumo_contas[df_resumo_contas['STATUS'] == 'APENAS XML'])}")
    print(f"  - Itens com diferenca de quantidade: {len(df_dif_qtd)}")
    print(f"  - Itens com diferenca de preco (>1 centavo): {len(df_dif_preco)}")
    print(f"  - Itens apenas no Excel: {len(df_apenas_excel)}")
    print(f"  - Itens apenas no XML: {len(df_apenas_xml)}")

    return df_resumo_protocolos, df_resumo_contas, df_dif_qtd, df_dif_preco, df_apenas_excel, df_apenas_xml


def gerar_relatorio(df_resumo_protocolos, df_resumo_contas, df_dif_qtd,
                    df_dif_preco, df_apenas_excel, df_apenas_xml):
    """Gera o relatório Excel final com todas as abas"""
    print("\n" + "=" * 70)
    print("ETAPA 4: Gerando relatorio Excel")
    print("=" * 70)

    with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:
        # Aba 1: Resumo de Protocolos
        df_resumo_protocolos.to_excel(writer, sheet_name='1-Resumo Protocolos', index=False)
        print(f"  Aba 1: Resumo Protocolos ({len(df_resumo_protocolos)} linhas)")

        # Aba 2: Resumo de Contas
        df_resumo_contas.to_excel(writer, sheet_name='2-Resumo Contas', index=False)
        print(f"  Aba 2: Resumo Contas ({len(df_resumo_contas)} linhas)")

        # Aba 3: Diferenças de Quantidade
        if df_dif_qtd is not None and len(df_dif_qtd) > 0:
            df_dif_qtd.to_excel(writer, sheet_name='3-Diferenca Quantidade', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhuma diferenca de quantidade encontrada']}).to_excel(
                writer, sheet_name='3-Diferenca Quantidade', index=False)
        print(f"  Aba 3: Diferenca Quantidade ({len(df_dif_qtd) if df_dif_qtd is not None else 0} linhas)")

        # Aba 4: Diferenças de Preço
        if df_dif_preco is not None and len(df_dif_preco) > 0:
            df_dif_preco.to_excel(writer, sheet_name='4-Diferenca Preco', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhuma diferenca de preco encontrada']}).to_excel(
                writer, sheet_name='4-Diferenca Preco', index=False)
        print(f"  Aba 4: Diferenca Preco ({len(df_dif_preco) if df_dif_preco is not None else 0} linhas)")

        # Aba 5: Itens apenas no Excel
        if df_apenas_excel is not None and len(df_apenas_excel) > 0:
            df_apenas_excel.to_excel(writer, sheet_name='5-Apenas Excel', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhum item exclusivo do Excel']}).to_excel(
                writer, sheet_name='5-Apenas Excel', index=False)
        print(f"  Aba 5: Apenas Excel ({len(df_apenas_excel) if df_apenas_excel is not None else 0} linhas)")

        # Aba 6: Itens apenas no XML
        if df_apenas_xml is not None and len(df_apenas_xml) > 0:
            df_apenas_xml.to_excel(writer, sheet_name='6-Apenas XML', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhum item exclusivo do XML']}).to_excel(
                writer, sheet_name='6-Apenas XML', index=False)
        print(f"  Aba 6: Apenas XML ({len(df_apenas_xml) if df_apenas_xml is not None else 0} linhas)")

    print(f"\n  Relatorio salvo em: {ARQUIVO_SAIDA}")


def main():
    """Função principal"""
    print("\n" + "=" * 70)
    print("COMPARACAO DE CONTAS MEDICAS - EXCEL vs XML")
    print("Versao 2: Com tolerancia de 1 centavo para arredondamento")
    print("=" * 70)

    # Etapa 1: Extrair dados dos XMLs
    resultado_xml = extrair_dados_xmls()
    itens_xml, protocolos_xml, contas_xml, arquivos_por_protocolo, protocolos_duplicados, protocolo_por_conta_xml = resultado_xml

    # Etapa 2: Processar Excel
    resultado_excel = processar_excel()
    df_excel_agrupado, df_excel_original, protocolos_excel, contas_excel, protocolo_por_conta_excel = resultado_excel

    # Etapa 3: Comparar dados
    resultados = comparar_dados(
        df_excel_agrupado, itens_xml,
        protocolos_excel, protocolos_xml,
        contas_excel, contas_xml,
        arquivos_por_protocolo, protocolos_duplicados,
        protocolo_por_conta_xml, protocolo_por_conta_excel
    )

    if resultados[0] is None:
        print("\nErro: Nao foi possivel realizar a comparacao.")
        return

    # Etapa 4: Gerar relatório
    gerar_relatorio(*resultados)

    print("\n" + "=" * 70)
    print("PROCESSAMENTO CONCLUIDO!")
    print("=" * 70)


if __name__ == "__main__":
    main()
