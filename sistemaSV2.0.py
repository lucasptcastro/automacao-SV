from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from time import sleep
from colorama import Fore, Back, Style, init
from datetime import date
from modulos.moeda import moedabr
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import pandas as pd
import modulos.uteis as uteis

init(autoreset=True)
mes = date.today().month-1

#FILTRO DE VARIÁVEIS
if mes < 10:
    endereco = fr'C:\Users\lucas\Downloads\mes0{mes}_22.xlsx'
    data = f'0{mes}/{date.today().year}' 
    arquivo = f'mes0{mes}_22.xlsx'
else:
    endereco = fr'C:\Users\lucas\Downloads\mes{mes}_22.xlsx'
    data = f'{mes}/{date.today().year}' 
    arquivo = f'mes{mes}_22.xlsx'


#ABRIR O ARQUIVO
tabela = pd.read_excel(endereco) 
faturamento = tabela['ValorFinal'].sum()
quantidade = tabela['Quantidade'].sum()


#MOSTRAR O FATURAMENTO NO CMD
uteis.cab("SUPLEMENTOS VIRTUAIS")
print(f'''O faturamento referente ao mês {Style.BRIGHT+data} foi de {Fore.GREEN+Style.BRIGHT+moedabr(faturamento)+Style.RESET_ALL}!
Foram vendidos {int(quantidade)} produtos!
''')
sleep(2)
print('Enviando o E-mail...')
uteis.espace()


#ENVIAR O E-MAIL
def sendEmail():
    
    #MONTANDO O CORPO DO E-MAIL
    corpo_email = f"""
    Equipe Suplementos Virtuais, tudo bem?<br><br>
    
    Segue abaixo o relatório de vendas referente ao mês <b>{data}</b><br>
    O faturamento fechou em: <b>{moedabr(faturamento)}</b><br>
    Foram vendidos <b>{int(quantidade)}</b> produtos!<br><br>
    
    Atenciosamente,<br><br>
    
    Kowalski, contador.
    """
    msg = MIMEMultipart()
    msg['Subject'] = f'Relatório de vendas {data}' #TÍTULO DO EMAIL
    msg['From'] = f'botsuplementosvirtuais@gmail.com' #DE
    msg['To'] = 'suplementosvirtuais@gmail.com' #PARA
    password = 'suplementosvirtuais2022' #SENHA DO E-MAIL 


    #ANEXO
    msg.attach(MIMEText(corpo_email, 'html')) #FAZER COM QUE O TEXTO SEJA LIDO EM HTML
    att = MIMEBase('application', 'octet-stream')
    attchment = open(endereco, 'rb') #ABRIR O ARQUIVO E LER EM BINÁRIO
    att.set_payload(attchment.read())
    encoders.encode_base64(att) #CONVERTER O ARQUIVO PARA BASE64
    att.add_header('Content-Disposition', f'attachment; filename= {arquivo}')
    msg.attach(att)
    attchment.close()


    #ENVIO DO E-MAIL
    s = smtplib.SMTP('smtp.gmail.com: 587') 
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email Enviado!')


sendEmail()
sleep(2)

