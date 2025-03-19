import pandas as pd
from selenium.webdriver.common.by import By
from time import sleep
from utils.selenium_utils import carregar_coordenadas_excel
from config_model import config  # Para acessar as credenciais
from utils.selenium_utils import inicializar_driver

# Função para acessar a área PIX
def acessar_area_pix(driver, coordenadas):
    print("Acessando a área PIX...")
    driver.find_element(By.XPATH, coordenadas['AREA_PIX']['xPATH']).click()
    sleep(3)

# Função para verificar o saldo inicial
def verificar_saldo_inicial(driver, matricula, coordenadas):
    print(f"Verificando saldo inicial para matrícula: {matricula}")
    driver.find_element(By.XPATH, coordenadas['SALDO_INICIAL']['xPATH']).click()
    driver.find_element(By.XPATH, coordenadas['SALDO_INICIAL']['xPATH']).send_keys(matricula)
    sleep(2)

# Função para preencher a matrícula
def preencher_matricula(driver, matricula, coordenadas):
    print(f"Preenchendo matrícula: {matricula}")
    driver.find_element(By.XPATH, coordenadas['MATRICULA']['xPATH']).click()
    driver.find_element(By.XPATH, coordenadas['MATRICULA']['xPATH']).send_keys(matricula)
    sleep(2)

# Função para verificar se o lançamento foi bem-sucedido
def verificar_lancamento_sucesso(driver, coordenadas):
    try:
        sucesso_element = driver.find_element(By.XPATH, coordenadas['SUCESSO']['xPATH'])
        if sucesso_element.is_displayed():
            return True
    except Exception:
        return False
    return False

# Função principal para processar PIX
def processar_pix():
    # Carregar coordenadas do Excel
    coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
    if coordenadas is None:
        print("Erro: coordenadas não carregadas.")
        return

    # Carregar a tabela de dados PIX
    tabela = pd.read_excel("PIX.xlsm")
    matriculas = tabela['Matr']
    transacoes = tabela['ID']
    valores = tabela['valor']
    observacoes = tabela['historico']

    # Inicializar o driver
    driver = inicializar_driver()

    # Acessar a área PIX
    acessar_area_pix(driver, coordenadas)

    # Processar cada entrada na tabela
    for matricula, transacao, valor, observacao in zip(matriculas, transacoes, valores, observacoes):
        print(f"Processando matrícula: {matricula}, Transação: {transacao}, Valor: {valor}")

        # Verificar saldo inicial
        verificar_saldo_inicial(driver, matricula, coordenadas)

        # Preencher matrícula
        preencher_matricula(driver, matricula, coordenadas)

        # Preencher os campos de transação
        driver.find_element(By.XPATH, coordenadas['ID']['xPATH']).click()
        driver.find_element(By.XPATH, coordenadas['ID']['xPATH']).send_keys(str(transacao))
        sleep(2)

        driver.find_element(By.XPATH, coordenadas['PIX']['xPATH']).click()
        driver.find_element(By.XPATH, coordenadas['PIX']['xPATH']).send_keys(f'{valor:.2f}')
        sleep(2)

        driver.find_element(By.XPATH, coordenadas['OBS_PIX']['xPATH']).click()
        driver.find_element(By.XPATH, coordenadas['OBS_PIX']['xPATH']).send_keys(observacao)
        sleep(2)

        # Lançar o PIX
        driver.find_element(By.XPATH, coordenadas['LANCAR_PIX']['xPATH']).click()
        sleep(3)

        # Verificar se o lançamento foi bem-sucedido
        if verificar_lancamento_sucesso(driver, coordenadas):
            print(f"PIX lançado com sucesso para matrícula {matricula}, valor R$ {valor}")
        else:
            print(f"Erro ao lançar PIX para matrícula {matricula}, valor R$ {valor}")
            continue

    # Finalizar o driver
    print("Processamento de PIX concluído.")
    driver.quit()

# Execução principal
if __name__ == "__main__":
    processar_pix()