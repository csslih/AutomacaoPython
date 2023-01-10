# AutomacaoPython
Este projeto trabalha automatização utilizando controle de teclado e mouse, onde o escopo é:
- acessar um diretório na nuvem
- exportar o arquivo excel contido no diretório
- realizar um cálculo que mostra o total das colunas de faturamento, e total de itens vendidos
- acessar a caixa de e-mail do usuário, gerar um novo e-mail contendo os dados de soma do item anterior
- encaminhar o e-mail para o endereço pré selecionado. 

# Passo 1: Importar bibliotecas 
import pyautogui
import pyperclip
import time
import pyinstaller

# Passo 2: Abrir o navegador utilizando os atalhos de teclado do computador
pyautogui.PAUSE = 1
pyautogui.hotkey("win")
time.sleep(1)
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(1)

# Passo 3: Inserir o atalho para seu diretório na nuvem
pyperclip.copy("https://drive.google.com/drive/folders....")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(2)

# Passo 4: Navegar até o local do relatório (Esses valores variam de acordo com o local das pastas e a resolução do monitor do usuário)
pyautogui.click(x=523, y=474, clicks=2)
time.sleep(2)

# Passo 5: Exportar o relatório

pyautogui.click(x=568, y=488) # clicar no arquivo
pyautogui.click(x=1614, y=357) # clicar nos 3 pontinhos
time.sleep(1)
pyautogui.click(x=1325, y=897) # clicar no fazer download
time.sleep(5) # esperar o download acabar

# Passo 6: Calcular os indicadores

import pandas as pd

tabela = pd.read_excel(r"C:\Users\carol\Downloads\Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)
# Passo 7: Enviar um e-mail com os resultados

# abrir aba
pyautogui.hotkey("ctrl", "t")


# entrar no link do email -
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)


# clicar no botão escrever
pyautogui.click(x=60, y=315)

# preencher as informações do e-mail
pyautogui.write("csslih@gmail.com")
pyautogui.press("tab") # selecionar o email

pyautogui.press("tab") # pular para o campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

time.sleep(2)

pyautogui.press("tab") # pular para o campo de corpo do email

texto = f"""
Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Caroline S.
"""

#Copiar o texto acima e colar no corpo do e-mail

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")
