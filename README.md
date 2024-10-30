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
## O Código
Aqui está um código genérico que pode ser utilizado para a extração de dados e download de anexos:

### Bibliotecas

```python
import os
import pandas as pd
from imap_tools import MailBox, AND
Funções
```

A função **baixarDados** extrai informações específicas do corpo do e-mail, como data, remetente e outras informações definidas (Informação 01, Informação 02, etc.), e as armazena no dicionário dic_infos.

```python
def baixarDados(email, dic_infos):
    dic_infos['Data'].append(email.date)
    dic_infos['Remetente'].append(email.from_)
    corpo = email.text
    for paragraph in corpo.split('\r\n'):
        if 'Informação 01' in paragraph:
            info1 = paragraph.split(':')[1].strip()
            dic_infos['Informação 01'].append(info1)
        if 'Informação 02' in paragraph:
            info2 = paragraph.split(':')[1].strip()
            dic_infos['Informação 02'].append(info2)
        if 'Informação 03' in paragraph:
            info3 = paragraph.split(':')[1].strip()
            dic_infos['Informação 03'].append(info3)
        if 'Informação 04' in paragraph:
            info4 = paragraph.split(':')[1].strip()
            dic_infos['Informação 04'].append(info4)
```

Essa função, **baixarAnexos**, percorre os anexos de um e-mail e, se o anexo for um arquivo CSV ou XLSX, salva-o na pasta especificada (pasta_destino).

```python
def baixarAnexos(email, pasta_destino):
    for attachment in email.attachments:
        if attachment.filename.endswith(('.csv', '.xlsx')):
            filepath = os.path.join(pasta_destino, attachment.filename)
            with open(filepath, 'wb') as f:
                f.write(attachment.payload)
```

### Informações Básicas

Esse código define as configurações de conexão ao servidor de e-mail (host, port, user, password) e inicializa um dicionário (dic_infos) para armazenar informações extraídas dos e-mails, como remetente, data e outras informações específicas.

```python
host     = 'pop.googlemail.com'
port     = '995'
user     = 'SEU E-MAIL'
password = 'SENHA DE APP DO SEU E-MAIL'

dic_infos = {'Remetente': [],
                    'Data': [],
                    'Informação 01': [],
                    'Informação 02': [],
                    'Informação 03': [],
                    'Informação 04': [],
                   }
```

### Acesso ao E-mail

Esse código se conecta à caixa de entrada do e-mail usando as credenciais fornecidas (user e password), e busca todos os e-mails não lidos (seen = False) sem marcá-los como lidos (mark_seen = False). Sinta-se à vontade para modificar esta etapa conforme necessário.

```python
bot_email = MailBox('imap.gmail.com').login(user, password)
emails = bot_email.fetch(AND(seen = False), mark_seen = False)
```

### Coleta de Dados

Esse código percorre os e-mails não lidos, verifica o assunto de cada e-mail e, se o assunto corresponder a “ASSUNTO PADRÃO”, extrai dados específicos do e-mail e os armazena em um dicionário (dic_infos). Se o assunto corresponder a “E-MAIL COM ANEXO”, baixa os anexos para uma pasta especificada. Em seguida, os dados coletados são convertidos em um DataFrame do Pandas e o DataFrame é salvo como um arquivo Excel (Dados dos E-mails.xlsx)

```python
for email in emails:
    assunto = email.subject
    if 'ASSUNTO PADRÃO' in assunto.upper():
        baixarDados(email, dic_infos)
    if 'E-MAIL COM ANEXO' in assunto.upper():
        baixarAnexos(email, 'PASTA PARA SALVAR OS ANEXOS')

df = pd.DataFrame(dic_infos)
df.to_excel('Dados dos E-mails.xlsx')
```
