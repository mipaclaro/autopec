import os
import PyPDF2
import pandas as pd

# Função para ler e exibir o texto do PDF, filtrando apenas as linhas relevantes
def ler_texto_pdf(nome_arquivo):
    print(f"Lendo o arquivo PDF: {nome_arquivo}")
    with open(nome_arquivo, 'rb') as file:
        leitor = PyPDF2.PdfReader(file)
        linhas_relevantes = []
        for pagina_numero in range(len(leitor.pages)):
            texto_pagina = leitor.pages[pagina_numero].extract_text()
            linhas_pagina = texto_pagina.split('\n')
            # Filtrar apenas as linhas com códigos de produtos
            for linha in linhas_pagina:
                if linha.strip() and any(char.isdigit() for char in linha.split()[0]):
                    linhas_relevantes.append(linha.strip())
            print(f"--- Página {pagina_numero + 1} ---")
            print('\n'.join(linhas_relevantes))  # Imprime as linhas relevantes extraídas de cada página
        return linhas_relevantes

# Função para salvar as linhas relevantes em um arquivo Excel
def salvar_em_excel(linhas, nome_arquivo):
    colunas_esperadas = ['Codigo', 'Descricao', 'Quantidade', 'Valor Unitario', 'Valor Total']
    dados = []
    for linha in linhas:
        partes = linha.split()
        if len(partes) >= 5:
            codigo = partes[0]
            quantidade = partes[-3]
            valor_unitario = partes[-2]
            valor_total = partes[-1]
            descricao = ' '.join(partes[1:-3])
            dados.append([codigo, descricao, quantidade, valor_unitario, valor_total])
        else:
            print(f"Linha ignorada (dados insuficientes): {linha}")
    # Converter para DataFrame
    df = pd.DataFrame(dados, columns=colunas_esperadas)
    # Salvar em um arquivo Excel
    df.to_excel(nome_arquivo, index=False)
    print(f"Dados salvos no arquivo Excel: {nome_arquivo}")

# Função para ler os arquivos Excel e fazer a correspondência e cálculo
def processar_dados(arquivo_pdf_excel, arquivo_distribuicao_excel, arquivo_saida):
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

    # Salvar o resultado em um novo arquivo Excel
    resultado_df.to_excel(arquivo_saida, index=False)
    print(f"Resultados salvos no arquivo Excel: {arquivo_saida}")

# Função principal para processar todos os PDFs e salvar os resultados
def processar_raios(lista_pdfs, arquivo_distribuicao):
    pasta_destino = 'DISTRIBUICAO'
    
    # Verifica se o diretório de destino existe, se não, cria
    if not os.path.exists(pasta_destino):
        print(f"Criando o diretório: {pasta_destino}")
        os.makedirs(pasta_destino)

    for raio in lista_pdfs:
        linhas_relevantes = ler_texto_pdf(f"{raio}.pdf")
        salvar_em_excel(linhas_relevantes, f"{pasta_destino}/resultado_{raio}.xlsx")
        processar_dados(f"{pasta_destino}/resultado_{raio}.xlsx", arquivo_distribuicao, f"{pasta_destino}/resultado_final_{raio}.xlsx")

# Lista de arquivos PDF para os raios
lista_pdfs = ['raio1', 'raio2', 'raio3', 'raio4']

# Chamada da função principal para processar todos os raios e salvar os resultados
if __name__ == "__main__":
    processar_raios(lista_pdfs, 'DISTRIBUICAO.xlsx')