#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[23]:

import pyautogui
import time
import pyperclip

#pausa padrao
pyautogui.PAUSE=1

#abrir navegador
#pyautogui.press("winleft")
#pyautogui.write("chrome")
#pyautogui.press("enter")

#Alerta sobre inicio do processo
pyautogui.alert("Começarei o processo, aperte OK e não mexa em mais nada!")

#Abrir aba
pyautogui.hotkey('ctrl', 't')

#Abrir drive
#Evitar pyautogui.write porque o navegador pode modificar o endereço com autocompletar
#Criar uma constante com o endereço e copiá-lo diretamente
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
#Esperar para o próximo comando, evitando que a açao seja realizada durante algum delay
time.sleep(10)
#Parte problemática do código
#Encontrar a posição do ponteiro na tela pyautogui.position()
#Usar o ponteiro do mouse para clicar em uma região da tela, caso mude resolução da tela/janela/posição do objeto vai dar errado
pyautogui.click(227, 194, clicks=2)
time.sleep(6)
pyautogui.click(227, 194)
pyautogui.click(1465, 125)
pyautogui.click(1362, 404)
time.sleep(6)

# In[26]:
#  Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

import pandas as pd

df = pd.read_excel(r'C:\Users\soare\Downloads\Vendas - Dez.xlsx')
display(df)
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()
display(faturamento, qtde_produtos)

# In[34]:
# ### Vamos agora enviar um e-mail pelo gmail

pyautogui.hotkey('ctrl', 't')
pyautogui.write("mail.google.com")
pyautogui.press("enter")
time.sleep(6)

#Criar novo email
pyautogui.click(90, 190)
pyautogui.write("seugmail@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
assunto = "Relatório Imersão"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
time.sleep(6)

#Criar corpo do email
texto = f"""
Saudações.

O faturamento de ontem foi de : R${faturamento:,.2f}
A quantidade de produtos vendidos foi de: {qtde_produtos:,}

Att.:

Paulo Soares"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
