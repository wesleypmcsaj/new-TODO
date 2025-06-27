README - Sistema de Automação de Documentos Oficiais
Descrição
Este projeto é um sistema de automação para geração, gerenciamento e impressão de documentos oficiais (Comunicações Internas e Ofícios) utilizado pela Secretaria de Assuntos Jurídicos. O sistema integra várias funcionalidades em um fluxo automatizado, desde a coleta de dados até a impressão final dos documentos.

Funcionalidades Principais
Integração com Google Sheets: Coleta dados de uma planilha online para processamento

Geração de documentos: Cria automaticamente documentos Word (CI e Ofícios) com base em templates

Geração de QR Codes: Inclui QR codes nos documentos gerados, onde tais QR codes servem para atualizar a planilha de registros desses documentos

Gerenciamento de numeração: Controla automaticamente a numeração sequencial realizada nos documentos

Integração com GIAP: Sistema de gestão de processos para registro dos mesmos

Impressão automática: Envia os documentos gerados para impressão

Armazenamento em nuvem: Salva cópias dos documentos gerados no Google Drive

Atualização de planilhas: Mantém registros atualizados no Google Sheets

Pré-requisitos
Python 3.7 ou superior

Conta de serviço do Google Cloud Platform com acesso às APIs:

Google Drive API

Google Sheets API

Google Chrome instalado

Microsoft Word instalado

Impressora configurada no sistema

Instalação
Clone este repositório:

bash
git clone [URL_DO_REPOSITORIO]
cd [NOME_DO_DIRETORIO]
Instale as dependências:

bash
pip install -r requirements.txt
Configure o arquivo service_account.json com as credenciais da conta de serviço do Google Cloud

Configure o arquivo infos_giap.json com as credenciais do sistema GIAP

Estrutura de Arquivos
text
.
├── 1Informações.xlsx                 # Planilha principal de dados
├── modelo_de_ci.docx                 # Template para Comunicações Internas
├── modelo_de_oficio.docx             # Template para Ofícios
├── modelo de assinaturas.docx        # Template para registro de assinaturas
├── service_account.json              # Credenciais do Google Service Account
├── infos_giap.json                   # Configurações do sistema GIAP
├── numeros_de_ci.txt                 # Registro de números de CI utilizados
├── numeros_de_oficio.txt             # Registro de números de Ofício utilizados
├── CIS e oficios/                    # Pasta de documentos gerados
└── QRcodesGerados/                   # Pasta temporária para QR codes
Como Usar
Preencha a planilha "CONTROLE DE OFÍCIOS" no Google Drive com os dados necessários

Execute o script principal:

bash
python main.py
Siga as instruções exibidas no console:

Forneça os intervalos de numeração quando solicitado

Confirme as operações de impressão

Fluxo de Trabalho
Coleta de dados da planilha Google

Atribuição automática de números identificadores de documento

Busca de informações complementares (nomes de secretários, cópia de demais informações nescessárias de plataformas web)

Geração de GIAPs quando aplicável

Criação dos documentos Word com QR codes com as informações obtidas através do web-scrapping

Impressão dos documentos

Upload para o Google Drive

Atualização da planilha de controle dos documentos

Limpeza dos arquivos gerados pós impressão

Personalização
Templates: Modifique os arquivos modelo_de_ci.docx e modelo_de_oficio.docx para alterar o layout dos documentos

Configurações: Ajuste os arquivos JSON para modificar credenciais e parâmetros do sistema

Intervalos de numeração de documentação: O sistema solicitará novos intervalos quando os atuais se esgotarem

Observações Importantes
O sistema foi desenvolvido para uso específico em um dos setores da Secretaria de Assuntos Jurídicos

Requer conexão com a internet para acessar serviços do Google e o sistema GIAP

Mantenha os arquivos de configuração em local seguro (não os compartilhe)

Para ver a sua execução e ter mais detalhes de como funciona tal automação, acesse o seguinte link:





Official Documents Automation System

Description
This project is an automation system for generating, managing, and printing official documents (Internal Communications and Official Letters) used by the Legal Affairs Department. The system integrates multiple functionalities into an automated workflow, from data collection to the final printing of documents.

Key Features
Google Sheets Integration: Collects data from an online spreadsheet for processing

Document Generation: Automatically creates Word documents (Internal Communications and Official Letters) based on templates

QR Code Generation: Includes QR codes in generated documents, which serve to update the document registry spreadsheet

Numbering Management: Automatically controls sequential document numbering

GIAP Integration: Process management system for document registration

Automatic Printing: Sends generated documents to the printer

Cloud Storage: Saves copies of generated documents on Google Drive

Spreadsheet Updates: Keeps records updated in Google Sheets

Requirements
Python 3.7 or higher

Google Cloud Platform service account with access to:

Google Drive API

Google Sheets API

Google Chrome installed

Microsoft Word installed

Printer configured on the system

Installation
Clone this repository:

bash
git clone [REPOSITORY_URL]  
cd [DIRECTORY_NAME]  
Install dependencies:

bash
pip install -r requirements.txt  
Configure the service_account.json file with Google Cloud service account credentials

Configure the infos_giap.json file with GIAP system credentials

File Structure
text
.  
├── 1Informations.xlsx                 # Main data spreadsheet  
├── internal_comm_template.docx        # Internal Communications template  
├── official_letter_template.docx      # Official Letters template  
├── signatures_template.docx           # Signatures registry template  
├── service_account.json               # Google Service Account credentials  
├── infos_giap.json                    # GIAP system settings  
├── internal_comm_numbers.txt          # Registry of used Internal Communication numbers  
├── official_numbers.txt               # Registry of used Official Letter numbers  
├── Generated_Documents/               # Folder for generated documents  
└── Generated_QRcodes/                 # Temporary folder for QR codes  
How to Use
Fill out the "DOCUMENT CONTROL" spreadsheet on Google Drive with the required data

Run the main script:

bash
python main.py  
Follow the instructions displayed in the console:

Provide numbering ranges when prompted

Confirm printing operations

Workflow
Data collection from Google Sheets

Automatic assignment of document identification numbers

Retrieval of complementary information (department heads' names, copying other required information from web platforms)

GIAP generation when applicable

Creation of Word documents with QR codes containing information obtained through web scraping

Document printing

Upload to Google Drive

Update of the document control spreadsheet

Cleanup of generated files after printing

Customization
Templates: Modify the internal_comm_template.docx and official_letter_template.docx files to change document layout

Settings: Adjust the JSON files to modify credentials and system parameters

Document numbering ranges: The system will request new ranges when current ones are exhausted

Important Notes
This system was developed for specific use in one of the Legal Affairs Department's sectors

Requires an internet connection to access Google services and the GIAP system

Keep configuration files secure (do not share them)

To see its execution and learn more about how this automation works, visit the following link: [INSERT_LINK_HERE]
