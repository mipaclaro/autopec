import pandas as pd
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from time import sleep
from utils.selenium_utils import carregar_coordenadas_excel, inicializar_driver

# Função para acessar a área de lançamento de preso
def lancamento_preso(driver, coordenadas):
    print("Acessando a área de Lançamento Preso...")
    sleep(2)
    driver.find_element(By.XPATH, coordenadas['Lançamento Preso']['xPATH']).click()
    sleep(2)

# Função para preencher a matrícula
def preencher_matricula(driver, matr, coordenadas):
    print(f"Preenchendo matrícula: {matr}")
    driver.find_element(By.XPATH, coordenadas['Matrícula']['xPATH']).click()
    sleep(1)
    matricula_field = driver.find_element(By.XPATH, coordenadas['Matrícula']['xPATH'])
    matricula_field.send_keys(str(matr))
    sleep(1)
    driver.find_element(By.XPATH, coordenadas['Botão Consultar']['xPATH']).click()
    sleep(3)

# Função para verificar e abrir saldo inicial
def verificar_saldo_inicial(driver, matr, coordenadas):
    print(f"Verificando saldo inicial para matrícula: {matr}")
    sleep(1)
    driver.find_element(By.XPATH, coordenadas['Matrícula']['xPATH']).click()
    sleep(1)
    matricula_field = driver.find_element(By.XPATH, coordenadas['Matrícula']['xPATH'])
    matricula_field.send_keys(str(matr))
    sleep(1)
    driver.find_element(By.XPATH, coordenadas['Botão Consultar']['xPATH']).click()
    sleep(5)
    try:
        gravar_button = driver.find_element(By.XPATH, coordenadas['Gravar']['xPATH'])
        gravar_button.click()
        sleep(4)
    except NoSuchElementException:
        print("Já possui conta aberta na Unidade!")
        lancamento_preso(driver, coordenadas)

# Função para processar lançamentos genéricos
def processar_lancamentos(tabela, coordenadas, driver, cod, observacao_coluna, valor_coluna):
    for index, row in tabela.iterrows():
        matr = row['MATR']
        valor = row[valor_coluna]
        observacao = row[observacao_coluna]
        print(f"Processando matrícula: {matr}, Valor: {valor}, Observação: {observacao}")
        preencher_matricula(driver, matr, coordenadas)
        driver.find_element(By.XPATH, coordenadas['Código de lançamento']['xPATH']).send_keys(str(cod))
        sleep(3)
        driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).click()
        driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).send_keys(observacao)
        sleep(3)
        driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).click()
        driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).send_keys(f'{valor:.2f}')
        sleep(3)
        driver.find_element(By.XPATH, coordenadas['Lançar']['xPATH']).click()
        sleep(5)

# Função para processar passagens
def passagem():
    print("Iniciando processamento de passagens...")
    tabela = pd.read_excel("passagens.xlsx")
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)
    processar_lancamentos(tabela, coordenadas, driver, cod=2003, observacao_coluna='PRODUTO', valor_coluna='VALOR')
    driver.quit()

# Função para processar retiradas
def processar_retirada():
    print("Iniciando processamento de retiradas...")
    tabela = pd.read_excel("retiradas.xlsx")
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)
    processar_lancamentos(tabela, coordenadas, driver, cod=2016, observacao_coluna='PRODUTO', valor_coluna='VALOR')
    driver.quit()

# Função para processar compras
def processar_compras():
    print("Iniciando processamento de compras...")
    tabela = pd.read_excel("compras.xlsx")
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)
    processar_lancamentos(tabela, coordenadas, driver, cod=3000, observacao_coluna='DESCRICAO', valor_coluna='VALOR')
    driver.quit()

# Função para liquidar contas
def liquidar_contas():
    print("Iniciando liquidação de contas...")
    tabela = pd.read_excel("liquidacao_contas.xlsx")
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)
    processar_lancamentos(tabela, coordenadas, driver, cod=4000, observacao_coluna='OBSERVACAO', valor_coluna='VALOR')
    driver.quit()

# Execução principal
if __name__ == "__main__":
    # Escolha qual funcionalidade executar
    print("Escolha uma funcionalidade:")
    print("1. Passagem")
    print("2. Retirada")
    print("3. Compras")
    print("4. Liquidação de Contas")
    escolha = input("Digite o número da funcionalidade desejada: ")

    if escolha == "1":
        passagem()
    elif escolha == "2":
        processar_retirada()
    elif escolha == "3":
        processar_compras()
    elif escolha == "4":
        liquidar_contas()
    else:
        print("Opção inválida!")