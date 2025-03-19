import os
import PyPDF2
import pandas as pd
from pandas import ExcelWriter

# Função para ler e exibir o texto do PDF, filtrando apenas as linhas relevantes por fornecedor
def ler_texto_pdf_por_fornecedor(nome_arquivo):
    print(f"Lendo o arquivo PDF: {nome_arquivo}")
    with open(nome_arquivo, 'rb') as file:
        leitor = PyPDF2.PdfReader(file)
        linhas_relevantes = []
        fornecedor_atual = None
        
        for pagina_numero in range(len(leitor.pages)):
            texto_pagina = leitor.pages[pagina_numero].extract_text()
            linhas_pagina = texto_pagina.split('\n')
            
            # Identificar e separar linhas por fornecedor
            for linha in linhas_pagina:
                if linha.strip().startswith("NOME:"):
                    fornecedor_atual = linha.strip().replace("NOME:", "").strip()
                    continue
                if fornecedor_atual and linha.strip() and any(char.isdigit() for char in linha.split()[0]):
                    linhas_relevantes.append((fornecedor_atual, linha.strip()))
            
            print(f"--- Página {pagina_numero + 1} ---")
            print('\n'.join([f"{fornecedor}: {linha}" for fornecedor, linha in linhas_relevantes]))  # Imprime as linhas relevantes extraídas de cada página
        
        return linhas_relevantes

# Função para salvar as linhas relevantes em arquivos Excel separados por fornecedor
def salvar_em_excel_por_fornecedor(linhas, prefixo_arquivo):
    colunas_esperadas = ['Fornecedor', 'Codigo', 'Descricao', 'Quantidade', 'Valor Unitario', 'Valor Total']
    dados_por_fornecedor = {}

    for fornecedor, linha in linhas:
        partes = linha.split()
        if len(partes) >= 5:
            codigo = partes[0]
            quantidade = partes[-3]
            valor_unitario = partes[-2]
            valor_total = partes[-1]
            descricao = ' '.join(partes[1:-3])
            if fornecedor not in dados_por_fornecedor:
                dados_por_fornecedor[fornecedor] = []
            dados_por_fornecedor[fornecedor].append([fornecedor, codigo, descricao, quantidade, valor_unitario, valor_total])
        else:
            print(f"Linha ignorada (dados insuficientes): {linha}")

    # Salvar em arquivos Excel separados por fornecedor
    pasta_destino = 'DISTRIBUICAO'
    if not os.path.exists(pasta_destino):
        print(f"Criando o diretório: {pasta_destino}")
        os.makedirs(pasta_destino)

    for fornecedor, dados in dados_por_fornecedor.items():
        df = pd.DataFrame(dados, columns=colunas_esperadas)
        nome_arquivo = f"{pasta_destino}/{prefixo_arquivo}_{fornecedor.replace(' ', '_')}.xlsx"
        df.to_excel(nome_arquivo, index=False)
        print(f"Dados salvos no arquivo Excel: {nome_arquivo}")

# Função para ler os arquivos Excel e fazer a correspondência e cálculo
def processar_dados(arquivo_pdf_excel, arquivo_distribuicao_excel):
    # Ler os arquivos Excel
    pdf_df = pd.read_excel(arquivo_pdf_excel)
    distribuicao_df = pd.read_excel(arquivo_distribuicao_excel)

    # Garantir que CODIGO_1 e CODIGO_2 são strings
    distribuicao_df['CODIGO_1'] = distribuicao_df['CODIGO_1'].astype(str).fillna('')
    distribuicao_df['CODIGO_2'] = distribuicao_df['CODIGO_2'].astype(str).fillna('')

    # Inicializar lista para armazenar os resultados
    resultados = []

    # Iterar sobre as linhas do DataFrame do PDF
    for index, row in pdf_df.iterrows():
        codigo_pdf = str(row['Codigo'])
        descricao_pdf = row['Descricao']
        quantidade_pdf = float(row['Quantidade'])
        
        # Encontrar linha correspondente no DataFrame de Distribuição considerando CODIGO_1 e CODIGO_2
        distribuicao_row = distribuicao_df[
            (distribuicao_df['CODIGO_1'].str.contains(codigo_pdf, na=False)) | 
            (distribuicao_df['CODIGO_2'].str.contains(codigo_pdf, na=False)) |
            (distribuicao_df['CODIGO_1'].str.contains(codigo_pdf.split('-')[0], na=False)) |
            (distribuicao_df['CODIGO_2'].str.contains(codigo_pdf.split('-')[0], na=False))
        ]
        
        if not distribuicao_row.empty:
            fardo = float(distribuicao_row['FARDO'].values[0])
            
            # Calcular Caixas e Unidades
            caixas = int(quantidade_pdf // fardo)
            unidades = quantidade_pdf - (caixas * fardo)

            # Adicionar resultado à lista
            resultados.append([codigo_pdf, descricao_pdf, quantidade_pdf, fardo, caixas, unidades])

    # Converter a lista de resultados para DataFrame
    colunas_resultado = ['Codigo', 'Descricao', 'Quantidade', 'FARDO', 'Caixa', 'Unidade']
    resultado_df = pd.DataFrame(resultados, columns=colunas_resultado)

    return resultado_df

# Função principal para processar todos os PDFs e salvar os resultados
def processar_compras_total(arquivo_pdf, arquivo_distribuicao):
    try:
        # Verificar se os arquivos existem
        if not os.path.exists(arquivo_pdf):
            raise FileNotFoundError(f"Arquivo PDF não encontrado: {arquivo_pdf}")
        if not os.path.exists(arquivo_distribuicao):
            raise FileNotFoundError(f"Arquivo de distribuição não encontrado: {arquivo_distribuicao}")

        linhas_relevantes = ler_texto_pdf_por_fornecedor(arquivo_pdf)
        prefixo_arquivo = 'resultado_compras_total'
        salvar_em_excel_por_fornecedor(linhas_relevantes, prefixo_arquivo)

        distribuicao_excel = arquivo_distribuicao
        with ExcelWriter(f'DISTRIBUICAO/resultado_final_completo.xlsx', engine='xlsxwriter') as writer:
            for fornecedor in set(fornecedor for fornecedor, _ in linhas_relevantes):
                arquivo_pdf_excel = f"DISTRIBUICAO/{prefixo_arquivo}_{fornecedor.replace(' ', '_')}.xlsx"
                resultado_df = processar_dados(arquivo_pdf_excel, distribuicao_excel)
                sheet_name = fornecedor.replace(' ', '_')[:31]
                resultado_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("Todos os resultados salvos no arquivo Excel: DISTRIBUICAO/resultado_final_completo.xlsx")
    
    except Exception as e:
        print(f"Erro ao processar os arquivos: {str(e)}")
        raise

# Chamada da função principal para processar o arquivo compras_total.pdf
if __name__ == "__main__":
    processar_compras_total('compras_total.pdf', 'DISTRIBUICAO.xlsx')