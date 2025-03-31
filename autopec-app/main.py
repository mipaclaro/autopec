from utils.selenium_utils import inicializar_driver, realizar_login, carregar_coordenadas_excel
from controllers.lancamento_preso_controller import passagem, processar_retirada, processar_compras, liquidar_contas
from controllers.transf_saldo_outras_unidades_controller import processar_lancamentos
from controllers.arquivo_controller import processar_compras_total, processar_raios

def exibir_menu():
    """Exibe o menu de opções para o usuário."""
    print("\n=== Sistema de Automação ===")
    print("1. Processar Passagem")
    print("2. Processar Retirada")
    print("3. Processar Compras")
    print("4. Liquidar Contas")
    print("5. Transferir Saldo para Outras Unidades")
    print("6. Processar Compras Total (PDF)")
    print("7. Processar Raios (PDF)")
    print("0. Sair")
    return input("Escolha uma opção: ")

def main():
    """Fluxo principal do aplicativo."""
    try:
        # Inicializar o driver do Selenium
        driver = inicializar_driver()

        # Realizar login no sistema
        realizar_login(driver)

        # Carregar coordenadas do Excel
        coordenadas = carregar_coordenadas_excel('coordenadas_mouse.xlsx')
        if not coordenadas:
            print("Erro: Falha ao carregar coordenadas. Verifique o arquivo 'coordenadas_mouse.xlsx'.")
            driver.quit()
            return

        print("Coordenadas carregadas com sucesso.")

        # Loop do menu principal
        while True:
            escolha = exibir_menu()
            try:
                if escolha == "1":
                    print("\nIniciando processamento de passagem...")
                    passagem()
                elif escolha == "2":
                    print("\nIniciando processamento de retirada...")
                    processar_retirada()
                elif escolha == "3":
                    print("\nIniciando processamento de compras...")
                    processar_compras()
                elif escolha == "4":
                    print("\nIniciando liquidação de contas...")
                    liquidar_contas()
                elif escolha == "5":
                    print("\nIniciando transferência de saldo para outras unidades...")
                    processar_lancamentos()
                elif escolha == "6":
                    print("\nIniciando processamento de compras total (PDF)...")
                    processar_compras_total('compras_total.pdf', 'DISTRIBUICAO.xlsx')
                elif escolha == "7":
                    print("\nIniciando processamento de raios (PDF)...")
                    lista_pdfs = ['raio1', 'raio2', 'raio3', 'raio4']
                    processar_raios(lista_pdfs, 'DISTRIBUICAO.xlsx')
                elif escolha == "0":
                    print("\nSaindo do sistema...")
                    break
                else:
                    print("\nOpção inválida! Tente novamente.")
            except Exception as e:
                print(f"Erro ao executar a funcionalidade escolhida: {e}")

    except Exception as e:
        print(f"Erro crítico: {e}")
    finally:
        # Finalizar o driver
        try:
            driver.quit()
        except:
            pass
        print("Sistema finalizado.")

if __name__ == "__main__":
    main()