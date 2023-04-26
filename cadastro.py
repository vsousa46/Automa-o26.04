from Classs import Cadastro
import sys
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import smtplib

sys.path.append(r'C:\Users\vitor.varelo\Desktop\Criação de Acessos\Classs.py')

tabela = pd.read_excel('Automacaoxl.xlsx')
#tamanho_coluna = tabela.shape[1]

for i, carteira in enumerate(tabela['CARTEIRA']):

    classcadastro = Cadastro(tabela.loc[i,'PRIMEIRO NOME'],tabela.loc[i,'SOBRENOME'],tabela.loc[i,'PERFIL MÓVEL'],tabela.loc[i,'NOME COMPLETO'],tabela.loc[i,'OLOS'], tabela.loc[i,'E-MAIL'], tabela.loc[i,'CPF'])
    try:
        if carteira == 'ATIVOS G1':
       
            classcadastro.zimbra()
            classcadastro.total()
            classcadastro.intra()
            classcadastro.oolos()

        elif carteira == 'ATIVOS G2':
            classcadastro.zimbra()
            classcadastro.total()
            classcadastro.intra()
            classcadastro.oolos()


        elif carteira == 'EMGEA':
            classcadastro.zimbra()
            classcadastro.total()
            classcadastro.intra()
            classcadastro.oolos()


        elif carteira == "JURÍDICO":
            classcadastro.zimbra()
            classcadastro.total()
        
        
        else:
       
            classcadastro.zimbra()
            classcadastro.cob()
            classcadastro.intra()
            classcadastro.oolos()
            classcadastro.total()
    
    except:
            email = 'vitor.varelo@ferreiraechagas.com.br'
            senha = 'Vtr021109!'
            cam_arquivo = 'C:/Users/vitor.varelo/Desktop/Automação TCD/COD/Termo de Aceite.pdf'
            host = "smtp.gmail.com"
            port = "587"

            server = smtplib.SMTP(host,port)

            server.ehlo()
            server.starttls()
            server.login(email,senha)

            corpo_email = f"""
            <p>Prezados,</p>
            <p></p>
            <p>Ocorreu um erro no cadastro de usuarios.</p>
            <p>Favor verificar.</p>
            <p></p>
            <p>Att,</p>
            """

            msg = MIMEMultipart()
            msg['Subject'] = '### ERRO NO CADASTRO DE USUARIO ###'
            msg['From'] = email
            msg['To'] = 'vitor.varelo@ferreiraechagas.com.br'
            msg.attach(MIMEText(corpo_email, 'html'))

            #enviar anexo
            attchment = open(cam_arquivo, 'rb')
            att = MIMEBase('application', 'octet-stream')
            att.set_payload(attchment.read())
            encoders.encode_base64(att)

            # att.add_header('Content-Disposition', f'attchment; filename=Termo de Aceite.pdf')
            # attchment.close()
            # msg.attach(att)


            #Enviar email
            server.sendmail(msg['From'],msg['To'], msg.as_string())
            server.quit()
