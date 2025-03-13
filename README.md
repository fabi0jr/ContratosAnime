# Gerador de Contratos Automáticos

Este projeto tem como objetivo facilitar a geração de contratos automáticos a partir de um modelo pré-definido. O usuário insere as informações necessárias, e o sistema gera um arquivo DOCX com os dados preenchidos, além de salvar essas informações em uma planilha do Google.

## Estrutura do Projeto

```
APPANIMECONTRAT
├── src
│   ├── assets
│   │   └── main_img.png          # Imagem de fundo
│   ├── credentials
│   │   └── service-account-file.json # Credenciais para a API do Google Sheets
│   ├── templates
│   │   └── contract_template.docx # Modelo do contrato em formato DOCX
│   ├── services
│   │   └── google_sheets_service.py # Interação com a API do Google Sheets
│   ├── utils
│   │   └── docx_generator.py     # Geração do arquivo DOCX a partir do modelo
│   └── main.py                  # Ponto de entrada da aplicação
├── requirements.txt              # Dependências do projeto
└── README.md                     # Documentação do projeto
```

## Instalação

1. Clone o repositório:
   ```sh
   git clone <URL do repositório>
   cd APPANIMECONTRAT
   ```

2. Crie e ative um ambiente virtual:
   ```sh
   python -m venv venv
   venv\Scripts\activate  # No Windows
   source venv/bin/activate  # No macOS/Linux
   ```

3. Instale as dependências:
   ```sh
   pip install -r requirements.txt
   ```

## Uso

1. Execute o arquivo principal:
   ```sh
   python src/main.py
   ```

2. Siga as instruções na tela para inserir as informações necessárias.

3. O contrato será gerado e salvo no formato DOCX na pasta de documentos do usuário, e os dados serão armazenados em uma planilha do Google.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests. Para contribuir, siga estas etapas:

1. Fork o repositório.
2. Crie uma nova branch (`git checkout -b feature/nome-da-sua-feature`).
3. Faça suas alterações e commit (`git commit -m 'Adicionando nova feature'`).
4. Envie para o repositório remoto (`git push origin feature/nome-da-sua-feature`).
5. Abra um Pull Request.

## Licença

Este projeto está licenciado sob a MIT License - veja o arquivo [LICENSE](LICENSE) para mais detalhes.