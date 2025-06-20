import smtplib
import email.message
from senha_email import senha_app
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import pandas as pd
import unicodedata

def normalizar_nome(nome):
    nome = nome.lower()
    nome = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('utf-8')
    return nome

link_logo = 'https://upload.wikimedia.org/wikipedia/commons/thumb/c/c2/Btg-logo-blue.svg/800px-Btg-logo-blue.svg.png'

def enviar_email(indice, nome, email):

    msg = MIMEMultipart()
    msg['Subject'] = f'Relatório BTG Maio - {nome}'
    msg['From'] = 'andre.melo@btgpactual.com'
    msg['To'] = email
    msg['Cc'] = 'matheus.polo@btgpactual.com'

    corpo_email = f'''<p>Olá {nome}, tudo bem?</p>
    <p>Segue em anexo o ralatório de investimentos do mês de março.</p>
    <p>Att, César</p>
    <img src="{link_logo}" style="width: 200px; height: auto;">''' #texto formatado em html

    msg.attach(MIMEText(corpo_email, 'html'))

    #anexar arquivos
    caminho_anexo = r'C:\Users\cesar\Desktop\Codigos\Integracao_e-mail\anexos'
    anexo = f'dashboard_{normalizar_nome(nome)}.pdf'
    caminho_arquivo = os.path.join(caminho_anexo, anexo)
    with open(caminho_arquivo, 'rb') as arquivo:
        parte_anexo = MIMEApplication(arquivo.read(), Name=f'Relatório_{nome}')
        parte_anexo.add_header('Content-Disposition', 'attachment', filename=f'Relatório_{nome}.pdf')
        msg.attach(parte_anexo)

    servidor = smtplib.SMTP('smtp.gmail.com', 587) #configuração para conectar com o provedor, no caso o gmail
    servidor.starttls() #criptografia
    servidor.login(msg['From'], senha_app)
    servidor.send_message(msg)
    servidor.quit()
    print(f'({indice})Email enviado')

df_email_clientes = pd.read_excel(r'C:\Users\cesar\Desktop\Codigos\Integracao_e-mail\email_clientes.xlsx')

for _, linha in df_email_clientes.iterrows():
    indice = linha['ID']
    nome_cliente = linha['Cliente']
    email_cliente = linha['Email']
    enviar_email(nome_cliente, email_cliente)