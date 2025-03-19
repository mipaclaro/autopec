import pandas as pd
from selenium.webdriver.common.by import By
from time import sleep
from utils.selenium_utils import carregar_coordenadas_excel, inicializar_driver

# Função para filtrar e excluir códigos específicos (tipo 5003)
def filtrar_codigos(tabela):
    print("Filtrando códigos...")
    tabela_filtrada = tabela[tabela['CODIGO'] != 5003]
    return tabela_filtrada

# Função para executar a automação do preenchimento
def executar_lancamentos(driver, tabela, coordenadas):
    codigos = tabela['CODIGO']
    historicos = tabela['HISTÓRICO']
    valores = tabela['VALOR']
    t = 3  # Tempo de espera para operações

    print("Iniciando os Lançamentos na Unidade...")
    sleep(3)
    
    # Clique inicial para focar na área correta
    driver.find_element(By.XPATH, coordenadas['Lançamento Unidade']['xPATH']).click()
    sleep(1)
    
    for cod, hist, valor in zip(codigos, historicos, valores):
        print(f"Processando Código: {cod}, Histórico: {hist}, Valor: {valor:.2f}")
        
        # Clicar no campo do código
        driver.find_element(By.XPATH, coordenadas['Código Unidade']['xPATH']).click()
        sleep(2)
        
        # Escrever o código
        codigo_field = driver.find_element(By.XPATH, coordenadas['Código Unidade']['xPATH'])
        codigo_field.send_keys(str(cod))
        sleep(3)
        
        # Clicar no campo de observação
        driver.find_element(By.XPATH, coordenadas['Observação Unidade']['xPATH']).click()
        
        # Escrever o histórico
        observacao_field = driver.find_element(By.XPATH, coordenadas['Observação Unidade']['xPATH'])
        observacao_field.send_keys(hist)
        sleep(3)
        
        # Clicar no campo de valor
        driver.find_element(By.XPATH, coordenadas['Valor Unidade']['xPATH']).click()
        
        # Escrever o valor
        valor_field = driver.find_element(By.XPATH, coordenadas['Valor Unidade']['xPATH'])
        valor_field.send_keys(f'{valor:.2f}')
        sleep(3)
        
        # Clicar em 'Lançar'
        driver.find_element(By.XPATH, coordenadas['Lançar Unidade']['xPATH']).click()
        sleep(t)
        
        # Refresh para preparar o próximo lançamento
        driver.find_element(By.XPATH, coordenadas['Lançamento Unidade']['xPATH']).click()
        sleep(4)
    
        print("==========")

# Função principal
def main():
    # Carregar a tabela de lançamentos
    print("Carregando a tabela de lançamentos...")
    tabela = pd.read_excel("lancamentos.xlsx")
    
    # Carregar as coordenadas calibradas do arquivo
    print("Carregando as coordenadas...")
    coordenadas = carregar_coordenadas_excel(arquivo='coordenadas_mouse.xlsx')
    if coordenadas is None:
        print("Erro: coordenadas não carregadas.")
        return
    
    # Exibir a tabela original
    print("Tabela Original:")
    print(tabela)
    
    # Filtrar os códigos (removendo o código 5003)
    tabela_filtrada = filtrar_codigos(tabela)
    
    # Exibir a tabela filtrada
    print("Tabela Filtrada (Sem Código 5003):")
    print(tabela_filtrada)
    
    # Inicializar o driver
    driver = inicializar_driver()
    
    # Executar os lançamentos com a tabela filtrada
    executar_lancamentos(driver, tabela_filtrada, coordenadas)
    
    # Finalizar o driver
    driver.quit()
    print("LANÇAMENTOS REALIZADOS COM SUCESSO!")

# Executar o script
if __name__ == "__main__":
    main()