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

# Função principal para processar lançamentos
def processar_lancamentos():
    print("Iniciando processamento de lançamentos...")
    # Carregar a tabela de dados
    tabela = pd.read_excel("OUTRAS.xlsx")

    # Carregar as coordenadas calibradas do arquivo
    coordenadas = carregar_coordenadas_excel(arquivo='coordenadas_mouse.xlsx')
    if coordenadas is None:
        print("Erro: coordenadas não carregadas.")
        return

    # Inicializar o driver
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)

    # Processar cada entrada na tabela
    for matr, disp, fr, unid in zip(
        tabela['MATR'], tabela['DISP'], tabela['FR'], tabela['UNID']
    ):
        print(f"Processando matrícula: {matr}, Disponível: {disp}, Fundo: {fr}, Unidade: {unid}")
        verificar_saldo_inicial(driver, matr, coordenadas)

        if disp > 0:
            cod = 1008
            valor = disp
            preencher_matricula(driver, matr, coordenadas)
            driver.find_element(By.XPATH, coordenadas['Código de lançamento']['xPATH']).send_keys(str(cod))
            sleep(3)
            driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).click()
            driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).send_keys(unid)
            sleep(3)
            driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).click()
            driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).send_keys(f'{valor:.2f}')
            sleep(3)
            driver.find_element(By.XPATH, coordenadas['Lançar']['xPATH']).click()
            sleep(5)

        if fr > 0:
            cod = 1014
            valor = fr
            preencher_matricula(driver, matr, coordenadas)
            driver.find_element(By.XPATH, coordenadas['Código de lançamento']['xPATH']).send_keys(str(cod))
            sleep(2)
            driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).click()
            driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).send_keys(unid)
            sleep(3)
            driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).click()
            driver.find_element(By.XPATH, coordenadas['Valor']['xPATH']).send_keys(f'{valor:.2f}')
            sleep(3)
            driver.find_element(By.XPATH, coordenadas['Lançar']['xPATH']).click()
            sleep(5)

        print("==========")

    # Finalizar o driver
    driver.quit()
    print("Processamento de lançamentos concluído.")

# Execução principal
if __name__ == "__main__":
    processar_lancamentos()