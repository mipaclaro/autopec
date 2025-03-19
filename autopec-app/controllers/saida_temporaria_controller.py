import pandas as pd
import re
from selenium.webdriver.common.by import By
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

# Função para imprimir e voltar à aba anterior
def imprimir(driver):
    print("Imprimindo...")
    sleep(3)
    driver.execute_script("window.print();")
    sleep(3)
    driver.switch_to.window(driver.window_handles[0])

# Função para processar passagens
def passagem():
    print("Iniciando processamento de passagens...")
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    if coordenadas is None:
        print("Erro: coordenadas não carregadas.")
        return

    # Carregar a tabela de dados
    tabela = pd.read_excel("passagens.xlsx", sheet_name=0)
    cod = 2003

    # Inicializar o driver
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)

    for nome, matr, prod, valor in zip(tabela['NOME'], tabela['MATR'], tabela['PRODUTO'], tabela['VALOR']):
        print(f"Processando {matr}, {nome}, {prod}, {valor}")
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
        print(f"Compra de Passagem lançada com sucesso! Matrícula: {matr}")
        sleep(5)

    driver.quit()
    print("Processamento de passagens concluído.")

# Função para processar retiradas
def processar_retirada():
    print("Iniciando processamento de retiradas...")
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    if coordenadas is None:
        print("Erro: coordenadas não carregadas.")
        return

    # Carregar a tabela de dados
    df_retirada = pd.read_excel("passagens.xlsx", sheet_name=1)
    cod = 2016

    # Inicializar o driver
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)

    for matr, obs, nome in zip(df_retirada['MATR'], df_retirada['PRODUTO'], df_retirada['NOME']):
        try:
            print(f"Processando matrícula {matr}, Nome: {nome}")
            preencher_matricula(driver, matr, coordenadas)
            disponivel_text = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div[3]/table/tbody/tr/td[2]').text
            disponivel_text = re.sub(r'[^\d,]', '', disponivel_text).replace(',', '.')
            valor_numerico = float(disponivel_text)
            driver.find_element(By.XPATH, coordenadas['Código de lançamento']['xPATH']).send_keys(str(cod))
            sleep(3)
            driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).click()
            driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).send_keys(obs)
            sleep(1)
            driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).click()
            driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).send_keys(f'{valor_numerico:.2f}')
            sleep(3)
            driver.find_element(By.XPATH, coordenadas['Lançar']['xPATH']).click()
            print(f"Retirada lançada com sucesso! Matrícula: {matr}, com o valor de R${valor_numerico:.2f}")
            sleep(5)
        except Exception as e:
            print(f"Erro ao processar a matrícula {matr}: {e}")
            continue

    driver.quit()
    print("Processamento de retiradas concluído.")

# Execução principal
if __name__ == "__main__":
    passagem()
    processar_retirada()