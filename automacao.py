import pyautogui
import time
import pyperclip
from IPython.display import display
import pandas as pd

pyautogui.PAUSE = 1  # definir o tempo de espera entre cada execução

# abrir o chrome
# pyautogui.press('winleft')
# pyautogui.write('chrome')
# pyautogui.press('enter')

#iniciar o projeto
pyautogui.alert('Vai começar, aperte OK e não toque em mais nada')
pyautogui.hotkey('ctrl', 't')

# abrir o drive
link = 'https://drive.google.com/drive/folders/1wRTFw0sUVBjRr4hW5U9LF7DjLixRyxym'
pyperclip.copy(link)  # melhor abordagem que copiar a string dentro do método. digitar a string aos poucos pode induzir a erros por autocorretor, por exemplo
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(7)  # instrui o sistema a aguardar até a operação completar

#baixar a base de dados
# time.sleep(5)
# pyautogui.position() # capturar posição do mouse
pyautogui.rightClick(1345, 495)
pyautogui.click(1395, 923)
time.sleep(10)


# acessar o banco de dados pelo pandas
df = pd.read_excel(r'C:\Users\gusta\Downloads\Vendas - Dez.xlsx')
display(df)

# calculando faturamento e quantidade de produtos
faturamento = df['Valor Final'].sum()
qtde = df['Quantidade'].sum()

# acessar o gmail
pyautogui.hotkey('ctrl', 't')
pyautogui.write('mail.google.com')
pyautogui.press('enter')
time.sleep(5)

# clicar em novo e-mail
pyautogui.click(52, 259)
pyautogui.write('gustavocambraia@outlook.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = 'Relatório de Vendas de Ontem'
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')

# escrever o corpo do e-mail
pyautogui.press('tab')
texto = f"""
Bom dia.

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos vendidos foi de: {qtde:,}

Obrigado.
Gustavo
"""
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')