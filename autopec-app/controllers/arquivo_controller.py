import os
import PyPDF2
import pandas as pd
from pandas import ExcelWriter

# Função para ler texto de um PDF e filtrar linhas relevantes
def ler_texto_pdf(nome_arquivo):
    print(f"Lendo o arquivo PDF: {nome_arquivo}")
    with open(nome_arquivo, 'rb') as file:
        leitor = PyPDF2.PdfReader(file)
        linhas_relevantes = []
        for pagina_numero in range(len(leitor.pages)):
            texto_pagina = leitor.pages[pagina_numero].extract_text()
            linhas_pagina = texto_pagina.split('\n')
            for linha in linhas_pagina:
                if linha.strip() and any(char.isdigit() for char in linha.split()[0]):
                    linhas_relevantes.append(linha.strip())
            print(f"--- Página {pagina_numero + 1} ---")
            print('\n'.join(linhas_relevantes))
        return linhas_relevantes

# Função para ler texto de um PDF por fornecedor
def ler_texto_pdf_por_fornecedor(nome_arquivo):
    print(f"Lendo o arquivo PDF: {nome_arquivo}")
    with open(nome_arquivo, 'rb') as file:
        leitor = PyPDF2.PdfReader(file)
        linhas_relevantes = []
        fornecedor_atual = None
        for pagina_numero in range(len(leitor.pages)):
            texto_pagina = leitor.pages[pagina_numero].extract_text()
            linhas_pagina = texto_pagina.split('\n')
            for linha in linhas_pagina:
                if linha.strip().startswith("NOME:"):
                    fornecedor_atual = linha.strip().replace("NOME:", "").strip()
                    continue
                if fornecedor_atual and linha.strip() and any(char.isdigit() for char in linha.split()[0]):
                    linhas_relevantes.append((fornecedor_atual, linha.strip()))
            print(f"--- Página {pagina_numero + 1} ---")
            print('\n'.join([f"{fornecedor}: {linha}" for fornecedor, linha in linhas_relevantes]))
        return linhas_relevantes

# Função para salvar dados em um arquivo Excel
def salvar_em_excel(dados, nome_arquivo, colunas):
    df = pd.DataFrame(dados, columns=colunas)
    df.to_excel(nome_arquivo, index=False)
    print(f"Dados salvos no arquivo Excel: {nome_arquivo}")

# Função para salvar dados em arquivos Excel separados por fornecedor
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

    pasta_destino = 'DISTRIBUICAO'
    if not os.path.exists(pasta_destino):
        print(f"Criando o diretório: {pasta_destino}")
        os.makedirs(pasta_destino)

    for fornecedor, dados in dados_por_fornecedor.items():
        nome_arquivo = f"{pasta_destino}/{prefixo_arquivo}_{fornecedor.replace(' ', '_')}.xlsx"
        salvar_em_excel(dados, nome_arquivo, colunas_esperadas)

# Função para processar dados de arquivos Excel
def processar_dados(arquivo_pdf_excel, arquivo_distribuicao_excel):
    pdf_df = pd.read_excel(arquivo_pdf_excel)
    distribuicao_df = pd.read_excel(arquivo_distribuicao_excel)
    distribuicao_df['CODIGO_1'] = distribuicao_df['CODIGO_1'].astype(str).fillna('')
    distribuicao_df['CODIGO_2'] = distribuicao_df['CODIGO_2'].astype(str).fillna('')
    resultados = []
    for index, row in pdf_df.iterrows():
        codigo_pdf = str(row['Codigo'])
        descricao_pdf = row['Descricao']
        quantidade_pdf = float(row['Quantidade'])
        distribuicao_row = distribuicao_df[
            (distribuicao_df['CODIGO_1'].str.contains(codigo_pdf, na=False)) |
            (distribuicao_df['CODIGO_2'].str.contains(codigo_pdf, na=False))
        ]
        if not distribuicao_row.empty:
            fardo = float(distribuicao_row['FARDO'].values[0])
            caixas = int(quantidade_pdf // fardo)
            unidades = quantidade_pdf - (caixas * fardo)
            resultados.append([codigo_pdf, descricao_pdf, quantidade_pdf, fardo, caixas, unidades])
    colunas_resultado = ['Codigo', 'Descricao', 'Quantidade', 'FARDO', 'Caixa', 'Unidade']
    return pd.DataFrame(resultados, columns=colunas_resultado)

# Função para processar compras totais
def processar_compras_total(arquivo_pdf, arquivo_distribuicao):
    linhas_relevantes = ler_texto_pdf(arquivo_pdf)
    salvar_em_excel(linhas_relevantes, 'resultado_compras_total.xlsx', ['Codigo', 'Descricao', 'Quantidade', 'Valor Unitario', 'Valor Total'])
    resultado_df = processar_dados('resultado_compras_total.xlsx', arquivo_distribuicao)
    resultado_df.to_excel('resultado_final_compras_total.xlsx', index=False)
    print("Processamento de compras total concluído.")

# Função para processar raios
def processar_raios(lista_pdfs, arquivo_distribuicao):
    for raio in lista_pdfs:
        linhas_relevantes = ler_texto_pdf(f"{raio}.pdf")
        salvar_em_excel(linhas_relevantes, f"resultado_{raio}.xlsx", ['Codigo', 'Descricao', 'Quantidade', 'Valor Unitario', 'Valor Total'])
        resultado_df = processar_dados(f"resultado_{raio}.xlsx", arquivo_distribuicao)
        resultado_df.to_excel(f"resultado_final_{raio}.xlsx", index=False)
        print(f"Processamento do raio {raio} concluído.")

# Execução principal
if __name__ == "__main__":
    print("Escolha uma funcionalidade:")
    print("1. Processar Compras Total")
    print("2. Processar Raios")
    escolha = input("Digite o número da funcionalidade desejada: ")

    if escolha == "1":
        processar_compras_total('compras_total.pdf', 'DISTRIBUICAO.xlsx')
    elif escolha == "2":
        lista_pdfs = ['raio1', 'raio2', 'raio3', 'raio4']
        processar_raios(lista_pdfs, 'DISTRIBUICAO.xlsx')
    else:
        print("Opção inválida!")