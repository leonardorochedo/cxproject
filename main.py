import pandas as pd
import pyautogui as pa
import pyperclip as pc
import time
from selenium import webdriver

pa.PAUSE = 5

### Excel

# Importando a base de dados
tabela = pd.read_excel("Vendas.xlsx") # Baixando a tabela

# Calculando indicadores
faturamento = tabela[r"LÍQUIDO"].sum() # Calculando quanto que vendeu no mês

### Enviar pro EMAIL

# Abrindo o Google.exe
navegador = webdriver.Chrome("chromedriver.exe")

# Entrando no email
navegador.get("https://outlook.live.com/mail/0/inbox")
time.sleep(5)

navegador.find_element_by_xpath("/html/body/header/div/aside/div/nav/ul/li[2]/a").click() # Clicando em Entrar
time.sleep(5)

pc.copy(r"learnir@outlook.com.br") # Digitando o email
pa.hotkey("ctrl", "v")
pa.hotkey("enter")
time.sleep(5)

pa.write("Leoluccadavi321") # A senha
pa.hotkey("enter")
time.sleep(5)

pa.hotkey("enter")
time.sleep(5)

# Digitando o e-mail
time.sleep(5)
pa.click(x=173, y=240, clicks=1) # Clique em Nova Mensagem
time.sleep(3.5)
pc.copy(r"wrufino+py@carbonxinsumos.com") # Colocando para quem vai ser enviado
time.sleep(4)
pa.hotkey("ctrl", "v")
pa.hotkey("tab")
pc.copy("Relatório de Vendas")
pa.hotkey("ctrl", "v")
pa.hotkey("tab")

texto = f"""
Prezados, bom dia

O faturamento do mês foi de: R${faturamento:,.2f}

Abs
Tu Rufino"""

pc.copy(texto)
pa.hotkey("ctrl", "v")
pa.hotkey("ctrl", "enter")

time.sleep(10)
pa.hotkey("alt", "f4")

##### FIM