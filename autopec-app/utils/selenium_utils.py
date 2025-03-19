import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from datetime import date
from selenium.common.exceptions import NoSuchElementException
from config_model import config  # Importa a classe Config para acessar as credenciais

# Inicializa o WebDriver
def inicializar_driver():
    print("INICIANDO O CHROMEDRIVER")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    return driver

# Realiza login no sistema
def realizar_login(driver):
    print("REALIZANDO LOGIN NO SISTEMA")
    driver.get("http://10.14.5.121/gpu/")
    sleep(5)

    # Localiza os campos de usuário e senha e o botão de acessar
    usuario_field = driver.find_element(By.NAME, 'usuario')
    senha_field = driver.find_element(By.NAME, 'senha')
    acessar_button = driver.find_element(By.ID, 'botao')

    # Insere usuário e senha principais
    usuario_field.send_keys(config.main_username)
    senha_field.send_keys(config.main_password)
    acessar_button.click()
    sleep(5)

    # Localiza e clica no elemento "Administrativo"
    admin_button = driver.find_element(By.XPATH, '//*[@id="tituloQuadrado"]')
    admin_button.click()
    sleep(5)

    # Preenche os campos de login adicionais
    login_field = driver.find_element(By.XPATH, '//*[@id="loginSiafem"]')
    senha_field = driver.find_element(By.XPATH, '//*[@id="senhaSiafem"]')
    enviar_button = driver.find_element(By.XPATH, '//*[@id="quadrado"]/form/fieldset/input[7]')

    login_field.send_keys(config.additional_username)
    senha_field.send_keys(config.additional_password)
    enviar_button.click()
    sleep(5)

# Carrega coordenadas de um arquivo Excel
def carregar_coordenadas_excel(arquivo='coordenadas_mouse.xlsx'):
    try:
        df = pd.read_excel(arquivo, index_col='Descrição')
        print(f"Coordenadas carregadas de {arquivo}.")
        return df.to_dict(orient='index')
    except FileNotFoundError:
        print(f"Arquivo {arquivo} não encontrado. Execute a calibração primeiro.")
        return None

# Exemplo de uso da data atual
def obter_data_atual():
    dia = date.today()
    return dia.strftime("%d%m%Y")