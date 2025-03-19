# Autopec

Automação de lançamento no sistema GPU da SAP.

O **Autopec** é uma aplicação desenvolvida para automatizar processos repetitivos no sistema GPU da SAP, como lançamentos de PIX, compras, transferências de saldo e liquidação de contas. Ele utiliza Selenium para interagir com a interface do sistema e pandas para manipulação de dados em arquivos Excel.

---

## 📂 Estrutura do Projeto

```plaintext
autopec-app/
├── main.py                    # Arquivo principal com a interface de linha de comando
├── controllers/               # Controladores para lógica de negócios
│   ├── pix_controller.py      # Lógica do PIX
│   ├── compras_controller.py  # Lógica das compras
│   ├── transfers_controller.py # Lógica das transferências
│   └── lancamento_preso_controller.py # Lógica de lançamentos para presos
├── models/                    # Modelos para gerenciamento de dados
│   └── config_model.py        # Gerenciamento de configurações
├── utils/                     # Funções auxiliares e utilitários
│   ├── selenium_helper.py     # Funções auxiliares do Selenium
│   └── excel_handler.py       # Funções para manipulação de Excel
├── notebooks/                 # Notebooks para experimentação e análise
│   └── exemplo_notebook.ipynb # Exemplo de uso dos controladores
├── resources/                 # Recursos adicionais
│   └── templates/             # Arquivos de configuração e templates
└── requirements.txt           # Dependências do projeto


✨ Funcionalidades
Lançamentos de PIX: Automatiza o processo de lançamento de PIX no sistema.
Processamento de Compras: Lê arquivos Excel e realiza lançamentos de compras.
Transferências de Saldo: Automatiza transferências de saldo entre unidades.
Liquidação de Contas: Automatiza o encerramento de contas no sistema.
Manipulação de Arquivos: Lê e processa arquivos PDF e Excel para integração com o sistema.
🛠️ Pré-requisitos
Certifique-se de ter os seguintes itens instalados no seu ambiente:

Python 3.8 ou superior
Google Chrome
ChromeDriver compatível com a versão do Chrome instalada
🚀 Instalação
1.Clone este repositório:
git clone https://github.com/seu-usuario/autopec.git
cd autopec

2.Crie um ambiente virtual e ative-o:
python3 -m venv venv
source venv/bin/activate  # No Windows: venv\Scripts\activate

3.Instale as dependências:
pip install -r requirements.txt

4.Configure o arquivo config.ini com as credenciais e configurações necessárias.

🖥️ Uso
Execute o arquivo principal main.py para acessar o menu de funcionalidades:
python main.py

Funcionalidades Disponíveis
Processar Passagem
Processar Retirada
Processar Compras
Liquidar Contas
Transferir Saldo para Outras Unidades
Processar Compras Total (PDF)
Processar Raios (PDF)
Escolha a opção desejada no menu interativo.

⚙️ Estrutura de Configuração
O arquivo config.ini deve conter as credenciais e configurações necessárias para acessar o sistema GPU. Exemplo:

[credentials]
main_username = seu_usuario
main_password = sua_senha
additional_username = usuario_adicional
additional_password = senha_adicional

🤝 Contribuição
Contribuições são bem-vindas! Siga os passos abaixo para contribuir:

Faça um fork do repositório.
Crie uma branch para sua feature ou correção:
git checkout -b minha-feature
Faça commit das suas alterações:
git commit -m "Minha nova feature"
Envie para o repositório remoto:
git push origin minha-feature
Abra um Pull Request.
📜 Licença
Este projeto está licenciado sob a MIT License.

📧 Contato
Para dúvidas ou sugestões, entre em contato pelo e-mail: rclaro@live.com
