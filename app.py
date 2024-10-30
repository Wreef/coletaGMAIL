import os
import pandas as pd
from imap_tools import MailBox, AND


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

def baixarAnexos(email, pasta_destino):
    for attachment in email.attachments:
        if attachment.filename.endswith(('.csv', '.xlsx')):
            filepath = os.path.join(pasta_destino, attachment.filename)
            with open(filepath, 'wb') as f:
                f.write(attachment.payload)


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

bot_email = MailBox('imap.gmail.com').login(user, password)
emails = bot_email.fetch(AND(seen = False), mark_seen = False)

for email in emails:
    assunto = email.subject
    if 'ASSUNTO PADRÃO' in assunto.upper():
        baixarDados(email, dic_infos)
    if 'E-MAIL COM ANEXO' in assunto.upper():
        baixarAnexos(email, 'PASTA PARA SALVAR OS ANEXOS')

df = pd.DataFrame(dic_infos)
df.to_excel('Dados dos E-mails.xlsx')
