import pyautogui
import pyperclip
from time import sleep
import pandas as pd
from modulos.moeda import moedabr
from datetime import date

pyautogui.PAUSE = 2

#variáveis
data = f'{date.today().month}/{date.today().year}' 
endereco = fr'C:\Users\Lucas\Downloads\mes{date.today().month}_22.xlsx' 

#abrir o google
pyautogui.press('win')
pyautogui.write('chr')
pyautogui.press('enter')

#ir até o drive
pyperclip.copy('https://drive.google.com/drive/u/1/folders/1TWTsFWbSp4G0RqwG6TEjF2tXwx0jRbtE')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#baixar o arquivo
sleep(1)
pyautogui.click(x=333, y=284, button='right') 
pyautogui.click(x=452, y=703) 

#abrir o arquivo
sleep(3)
tabela = pd.read_excel(endereco)         
faturamento = tabela['ValorFinal'].sum()
quantidade = tabela['Quantidade'].sum()
print(f'''O faturamento referente ao mês \033[1m{data}\033[m foi de \033[1m{moedabr(faturamento)}\033[m!
Foram vendidos \033[1m{int(quantidade)}\033[m produtos!
''')

#abrir o google
sleep(1)
pyautogui.press('win')
pyautogui.write('chr')
pyautogui.press('enter')

#ir até o gmail
pyperclip.copy('https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#mandar o e-mail
sleep(2)
pyautogui.click(x=86, y=209)
#endereçamento
pyautogui.write('suplementosvirtuais@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
#título
pyperclip.copy(f'Relatório de vendas {data}')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
#corpo do texto
pyperclip.copy(f'''Equipe Suplementos Virtuais, tudo bem?

Segue abaixo o relatório de vendas referente ao mês {data}
O faturamento fechou em: {moedabr(faturamento)}
Foram vendidos {int(quantidade)} produtos!

Atenciosamente, 
Lucas Peixoto.

''')
pyautogui.hotkey('ctrl','v')

#anexar a planilha
pyautogui.click(x=1115, y=1008)
pyperclip.copy(endereco)          
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')

#enviar o e-mail
sleep(2)
pyautogui.hotkey('ctrl', 'enter')