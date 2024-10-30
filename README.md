![logo](https://i.ibb.co/YthtbLh/Giifff-mid.gif)
***
# Coleta de Dados do GMAIL
Neste guia, você criará uma coleta dados de e-mails do GMAIL. Sendo possível armazenar as informações do corpo do e-mail e os arquivos anexados.
***
## O Desafio
A gestão de e-mails pode ser uma tarefa árdua, especialmente quando precisamos extrair informações específicas e anexos de forma manual. Esse processo não só consome tempo, mas também está sujeito a erros. O objetivo do projeto foi criar uma solução automatizada que pudesse coletar dados relevantes e baixar arquivos CSV e XLSX anexados aos e-mails, facilitando a análise e o armazenamento dessas informações.
***
## A Solução
Utilizando a biblioteca imap_tools, foi possível desenvolver um script em Python que se conecta à caixa de entrada do e-mail, filtra as mensagens não lidas e verifica se elas contêm um assunto específico. A partir daí, o script extrai as informações necessárias do corpo do e-mail e baixa os anexos relevantes, salvando-os em uma pasta designada no Google Drive.
***
## Principais Funcionalidades:
- **Conexão Segura**: Utilização de imap_tools para acessar a caixa de entrada de forma segura.
- **Filtragem de E-mails**: Busca por e-mails não lidos com assuntos específicos.
- **Extração de Dados**: Coleta de informações do corpo do e-mail utilizando expressões regulares.
- **Download de Anexos**: Salvamento automático de arquivos CSV e XLSX em uma pasta específica no Google Drive.
- **Organização dos Dados**: Armazenamento das informações extraídas em um DataFrame do Pandas para fácil manipulação e análise.
***
## Resultados e Benefícios
A implementação deste projeto trouxe diversos benefícios, incluindo:

- **Redução de Tempo**: A automação do processo de extração de dados e download de anexos economizou horas de trabalho manual.
- **Precisão**: A utilização de expressões regulares garantiu a extração precisa das informações necessárias.
- **Organização**: Os dados extraídos foram organizados em um DataFrame, facilitando a análise e a manipulação.

Com projetos desse tipo, é possível transformar diversas ideias em realidade. No meu caso, utilizei os dados coletados para criar um painel de monitoramento que se atualiza automaticamente com as informações recebidas por e-mail.
***
