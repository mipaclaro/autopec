import pandas as pd
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
    sleep(7)
    driver.switch_to.window(driver.window_handles[-1])
    driver.switch_to.window(driver.window_handles[0])
    sleep(3)

# Função principal para liquidar contas
def liquidar_contas():
    print("Iniciando liquidação de contas...")
    # Carregar a tabela de liquidação de contas
    tabela = pd.read_excel("LIQUIDACAO_CONTAS.xlsx")
    cod = 3000

    # Carregar as coordenadas calibradas do arquivo
    coordenadas = carregar_coordenadas_excel(arquivo='coordenadas_mouse.xlsx')
    if coordenadas is None:
        print("Erro: coordenadas não carregadas.")
        return

    # Inicializar o driver
    driver = inicializar_driver()
    lancamento_preso(driver, coordenadas)

    # Processar cada entrada na tabela
    for matr, unid, disp, fr, nome in zip(
        tabela['Matr.'], tabela['Unidade'], tabela['Saldo'], tabela['Fundo'], tabela['Nome']
    ):
        try:
            print(f"Processando matrícula {matr}, Nome: {nome}")
            preencher_matricula(driver, matr, coordenadas)
            unidade_text = driver.find_element(By.CLASS_NAME, 'fichaSimplesUnidade').text
            driver.find_element(By.XPATH, coordenadas['Código de lançamento']['xPATH']).send_keys(str(cod))
            driver.find_element(By.XPATH, coordenadas['Observação']['xPATH']).send_keys(f'TRANSF. P/ {unidade_text}')
            sleep(5)
            driver.find_element(By.XPATH, coordenadas['Lançar']['xPATH']).click()
            sleep(10)
            driver.find_element(By.XPATH, coordenadas['Botão Sim Encerrar']['xPATH']).click()
            sleep(5)
            imprimir(driver)
            print(f"Conta de {nome} encerrada com sucesso.")
        except Exception as e:
            print(f"Erro ao processar a matrícula {matr}: {e}")
            continue

    # Finalizar o driver
    driver.quit()
    print("Liquidação de contas concluída.")
    print("Verificar se tem pertences para enviar junto!")

# Execução principal
if __name__ == "__main__":
    liquidar_contas()