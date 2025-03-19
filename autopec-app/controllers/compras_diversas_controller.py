import pandas as pd
from selenium.webdriver.common.by import By
from time import sleep
from datetime import date
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

# Função para imprimir e voltar à aba anterior
def imprimir(driver):
    print("Imprimindo...")
    sleep(3)
    driver.switch_to.window(driver.window_handles[-1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    sleep(3)

# Função principal para processar compras
def processar_compras():
    print("Iniciando processamento de compras...")
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    if coordenadas is None:
        print("Erro: coordenadas não carregadas.")
        return

    # Carregar a tabela de dados
    tabela = pd.read_excel("COMPRAS2008.xlsx", sheet_name=0, dtype={
        'MATR': 'int64',
        'CODIGO': 'Int64',
        'VALOR': float
    })

    # Inicializar o driver
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)

    # Processar cada entrada na tabela
    for cod, matr, prod, valor in zip(tabela['CODIGO'], tabela['MATR'], tabela['PRODUTO'], tabela['VALOR']):
        print(f"Processando {matr}, {cod}, {prod}, {valor}")
        preencher_matricula(driver, matr, coordenadas)
        driver.find_element(By.XPATH, coordenadas['Código de lançamento']['xPATH']).send_keys(str(cod))
        sleep(3)
        driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).click()
        driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).send_keys(prod)
        sleep(1)
        driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).click()
        driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).send_keys(f'{valor:.2f}')
        sleep(3)
        driver.find_element(By.XPATH, coordenadas['Lançar']['xPATH']).click()
        print(f"Lançado com sucesso! Matrícula: {matr}")
        sleep(5)

    driver.quit()
    print("Processamento de compras concluído.")

# Execução principal
if __name__ == "__main__":
    processar_compras()