import pyautogui
import pyperclip
import time
import pandas as pd
import os
import datetime
import timedelta
pyautogui.PAUSE = 0.5
# abrir o navegador
pyautogui.press('win')
pyautogui.write('Edge')
pyautogui.press('Enter')

# abrir o drive
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('Ctrl', 'V')
pyautogui.press('Enter')

# abrir o exportar de drive
time.sleep(3)
pyautogui.click(x=425, y=276, clicks=2)

# selecionar com o direito a pasta
time.sleep(3)
pyautogui.click(x=392, y=366, button='Right')

# Selecionar a opção download
time.sleep(3)
pyautogui.click(x=498, y=640)

# Abrir a planilha
time.sleep(7)
tabela = pd.read_excel(r'C:\Users\gusta\Downloads\Vendas - Dez.xlsx')
# print(tabela)

quantidade = tabela['Quantidade'].sum()
faturamento = tabela['Valor Final'].sum()
print(f'Vendemos um total de {quantidade} produtos e um faturamento de R${float(faturamento):.2f}')

# Excluindo arquivo
arquivo = 'Vendas - Dez.xlsx'
localizacao = r"\Users\gusta\Downloads"
path = os.path.join(localizacao, arquivo)
os.remove(path)
print('Arquivo Excluído')

# Entrar no gmail
pyautogui.press('win')
pyautogui.write('Edge')
pyautogui.press('Enter')
pyperclip.copy('https://mail.google.com/mail/u/0/?tab=km#inbox')
pyautogui.hotkey('Ctrl', 'v')
pyautogui.press('Enter')

# Enviar o gmail
# clicando no botão +
time.sleep(5)
pyautogui.click(x=85, y=166)
destinatario = 'email para quem deseja enviar'
assunto = f'Relatorio de Vendas do dia'
corpo_do_email = f'''Prezada Bom dia
O faturamento de ontem foi de R$ {faturamento:,.2f}
A quantidade de produtos foi de {quantidade}

Abs
Gustavo'''
pyautogui.write(destinatario)
pyautogui.press('Enter')
pyautogui.click(x=866, y=346)
pyautogui.write(assunto)
pyautogui.click(x=880, y=456)
pyautogui.write(corpo_do_email)
pyautogui.hotkey('Ctrl', 'Enter')