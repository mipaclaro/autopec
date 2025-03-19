# autopec
Automação de lançamento no sistema GPU da SAP

autopec-app/
├── src/
│   ├── main.py                    # Arquivo principal com a interface
│   ├── controllers/
│   │   ├── __init__.py
│   │   ├── pix_controller.py      # Lógica do PIX
│   │   ├── compras_controller.py  # Lógica das compras
│   │   └── transfers_controller.py # Lógica das transferências
│   ├── models/
│   │   ├── __init__.py
│   │   └── config_model.py        # Gerenciamento de configurações
│   ├── views/
│   │   ├── __init__.py
│   │   ├── main_window.py         # Interface principal
│   │   └── components/
│   │       └── custom_widgets.py   # Componentes reutilizáveis
│   └── utils/
│       ├── __init__.py
│       ├── selenium_helper.py      # Funções auxiliares do Selenium
│       └── excel_handler.py        # Funções para manipulação de Excel
├── resources/
│   └── templates/                  # Arquivos de configuração e templates
└── requirements.txt               # Dependências do projeto
