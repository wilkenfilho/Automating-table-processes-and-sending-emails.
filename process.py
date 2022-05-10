!pip install pyautogui       
!pip install pyperclip

import pyautogui
import pyperclip
import time

#pyautogui.hotkey -> conjunto de teclas
#pyautogui.write -> escrever um texto
#pyautogui.press -> apertar uma tecla

pyautogui.PAUSE = 1                             #o PAUSE tem que estar maiusculo mesmo, o pause serve para pausar 1 segundo entre todos os comandos.

# Passo 1: Entrar no sistema da empresa (no nosso caso vai ser o link do drive).
pyautogui.hotkey("Command", "t")                #no windows é ctrl t
                                                #pyautogui.copy serve para copiar e o pyautogui.hotkey serve para fazer o atalho de colar.
pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.press("return")

# Passo 2: Navegar no sistema e encontrar a base de dados (Entrar na pasta Exportar).
                                                #o time serve para controlar o tempo
time.sleep(5)                                   #o time vai esperar 5 segundos nesse MOMENTO.
                                                #O time é essencial pq as vezes a internet pode demorar para carregar os arquivos.
pyautogui.click(x=366, y=340, clicks=2)         #o clicks serve para clicar duas vezes
                                                #O X e o Y significam a posição do mouse.
pyautogui.press("return")

# Passo 3: Fazer o download da base de Dados.
time.sleep(3)                                   #tempo para abrir a pasta
pyautogui.click(x=351, y=449)                   #clicar no arquivo
pyautogui.click(x=1242, y=215)                  #clicar nos tres pontinhos
pyautogui.click(x=998, y=658)                   #clicar no download
time.sleep(5)                                   #tempo para fazer o download



# Passo 4: importar a base de dados para o Python.
import pandas as pd                             #pandas é o pacote de python que permite trabalhar com dados de forma eficiente
tabela = pd.read_excel(r"/Users/soraialima/Downloads/Vendas - Dez.xlsx") 
                                                #se tiver abas na planilha, vc pode colocar uma vírgula após o caminho e colocar sheets = 1 (referente a primeira aba)
                                                #pd.read_excel serve para ler arquivos em excel
                                                #o r no inicio do caminho serve para dizer pro python que isso é um caminho.
display(tabela)                                 #mostrar a tabela

# Passo 5: Calcular os indicadores.
#faturamento - soma da coluna valor final
faturamento = tabela["Valor Final"].sum()

#quantidade de produtos - soma da quantidade de produtos
quantidade = tabela["Quantidade"].sum()

print(faturamento)
print(quantidade)

# Passo 6: Enviar um email para a diretoria com o relatório.
pyautogui.hotkey("Command", "t")   #abrir nova aba ou ctrl t se for windows.
pyautogui.write("dhttps://gmail.com")#abrir o e-mail
pyautogui.hotkey("Return")         #ou enter se for windows
time.sleep(3)                      #tempo para carregar o site
pyautogui.click(x=88, y=223)       #escrever
pyautogui.click(x=966, y=372)      #para
pyautogui.write("email_do_destinatário")
pyautogui.hotkey("Return")         #ou enter se for windows
pyautogui.click(x=960, y=407)      #assunto
pyautogui.write("Automatizando")
pyautogui.click(x=994, y=436)      #texto
texto = f"""
Oi, amor. 
Estou enviando este e-mail como teste.

Segue a planilha, os valores de faturamento e as quantidades:

Faturamento: {faturamento}
Quantidade: {quantidade}
Espero que goste.

Com amor,

Wilken

"""
#para o python entender várias linhas de texto, precisa-se colocar três aspas e para o python reconhecer as chaves, é preciso acrescentar um f antes das primeiras aspas.
pyautogui.write(texto)

pyautogui.click(x=1049, y=750)     #arquivo
pyautogui.click(x=427, y=176)      #selecionar arquivo
pyautogui.click(x=567, y=206)
pyautogui.hotkey("Return")
time.sleep(10)                     #tempo para carregar o arquivo
pyautogui.click(x=943, y=741)
#enviar
