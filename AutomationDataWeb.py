import os

import config

import datetime

import pandas as pd

import sys

from selenium import webdriver

from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.support import expected_conditions as EC

import time

import autoit as autoit

import openpyxl

import locale

from msedge.selenium_tools import Edge, EdgeOptions

from selenium.webdriver.edge.options import Options

import pyautogui



#___________________________________________________________________________________________________________________________

#Inicia o modulo webdriver com o Edge entrando no Sistema WEB da empresa, utilizando a biblioteca selenium



# Inicializando o WebDriver do Edge com as opções configuradas

time.sleep(10)

edge_driver_path = 'caminho do msedgedriver.exe'

options = webdriver.EdgeOptions()

service = webdriver.EdgeService(edge_driver_path)

driver = webdriver.Edge (service=service,options=options)





driver.get('Site do sistema web da empresa')

driver.set_window_position(-2000, 0)



print("Iniciando o Navegador")

driver.minimize_window()

driver.set_window_position(-2000, 0)

#___________________________________________________________________________________________________________________________

#iniciando o assistente autoit para logar na tela de "Segurança do Windows"

print("Logando nas permissões do Windows")

autoit.win_wait("Segurança do Windows")



autoit.win_activate("Segurança do Windows")

usuario = "Login"

autoit.send(usuario) #Login de Usuário





autoit.send("{TAB}")

senha = "Senha"

autoit.send(senha)  # Senha de Usuário



autoit.send("{ENTER}")

#Finaliza o log na tela "Segurança do Windows"

#___________________________________________________________________________________________________________________________

pyautogui.hotkey('alt', 'tab')

time.sleep(1)

pyautogui.hotkey('win', 'down')

pyautogui.hotkey('win', 'down')

#Iniciando o login no Sistema WEB da empresa

#Selecionando a Opção do menu suspenso

print("Logando no Sistema WEB da empresa")

wait = WebDriverWait(driver, 15)# atributo para esperar a página carregar

dropdown_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[xpath para selecionar opção desejada]")))

# Clicando no elemento para abrir o dropdown

dropdown_element.click()

time.sleep(2)

wait = WebDriverWait(driver, 5)

print("Selecionando Opção do menu suspenso")

# Aguardando até que a opção "Opção do menu suspenso" esteja visível e clicável

option_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[xpath para selecionar opção desejada]")))

# Clicando na opção "Opção do menu suspenso" para selecioná-la

option_element.click()

print("Preenchendo os campos de Login no Sistema WEB da empresa")

#iniciando preenchimento dos campos do login Sistema WEB da empresa



wait = WebDriverWait(driver, 5)# atributo para esperar a página carregar



# Localizando os campos "User name" e "Password"

username_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "xpath para selecionar opção desejada'username'] input")))

password_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "xpath para selecionar opção desejada'password'] input")))



# Preenchendo os campos com os dados desejados

username_element.send_keys("Login")

password_element.send_keys("Senha")

#Localizando o botão "Login"

login_button = wait.until(EC.visibility_of_element_located((By.XPATH, "//xpath para selecionar opção desejada")))



# Clicando no botão "Login"

login_button.click()

time.sleep(4)

#Finalizado Login Sistema WEB da empresa

#___________________________________________________________________________________________________________________________

print("Selecionando menu")

# Selecionando Menu

wait = WebDriverWait(driver, 10)# atributo para esperar a página carregar



menu_button = wait.until(EC.visibility_of_element_located((By.XPATH, "xpath para selecionar opção desejada")))

time.sleep(1)

# Clicando no ícone "Menu Button"

menu_button.click()

time.sleep(2)

#___________________________________________________________________________________________________________________________

#Clicando na opção do menu na pagina de consulta

print("Selecionando opção do menu ")

time.sleep(1)

wait = WebDriverWait(driver, 5)

elemento_opção_do_menu  = wait.until(EC.element_to_be_clickable((By.XPATH, '/xpath para selecionar opção desejada')))

driver.maximize_window()

driver.set_window_position(-2000, 0)

time.sleep(0.5)

elemento_opção_do_menu.click()

driver.minimize_window()

driver.set_window_position(-2000, 0)

#___________________________________________________________________________________________________________________________

#Clicando em Consultar

time.sleep(2)

print("Selecionando 'Consultar'")

wait = WebDriverWait(driver, 5)

elemento_a_span = wait.until(EC.element_to_be_clickable((By.XPATH, 'xpath para selecionar opção desejada')))



# Clicando em Consultar

elemento_a_span.click()



time.sleep(420)



wait = WebDriverWait(driver, 10)

elemento_a_span = wait.until(EC.element_to_be_clickable((By.XPATH, 'xpath para selecionar opção desejada')))

print("Selecionando 'Exportar'")

# Clicando em Exportar

elemento_a_span.click()

pyautogui.hotkey('alt', 'tab')

driver.set_window_position(-2000, 0)

pyautogui.hotkey('win', 'down')

time.sleep(470)







#___________________________________________________________________________________________________________________________



print("Iniciando a importação do arquivo exportado")



# Caminho da pasta desejada

pasta_alvo = "pasta de downloads"



# Obtendo a lista de arquivos na pasta

lista_arquivos = os.listdir(pasta_alvo)

print("Procurando arquivo mais recente")

# Filtrando apenas os arquivos (excluindo pastas)

lista_arquivos = [arquivo for arquivo in lista_arquivos if os.path.isfile(os.path.join(pasta_alvo, arquivo))]



# Se não houver arquivos na pasta, imprime uma mensagem e termina o script

if not lista_arquivos:

    print("Não há arquivos na pasta especificada.")

    exit()

print("Obtendo arquivo mais recente")

# Obtendo o arquivo mais recente com base na data de modificação

arquivo_mais_recente = max([os.path.join(pasta_alvo, arquivo) for arquivo in lista_arquivos], key=os.path.getmtime)



# Importando o arquivo como dataframe, pulando a primeira linha

df = pd.read_excel(arquivo_mais_recente, skiprows=1)



# Gerando o nome do arquivo com base na data e hora do arquivo

data_hora_arquivo = datetime.datetime.fromtimestamp(os.path.getmtime(arquivo_mais_recente)).strftime("%d_%m %Hhrs")

nome_arquivo = f"Arquivo_Exportado_{data_hora_arquivo}.xlsx"



#Definindo tipo de Data/Hora do Brasil

locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

# Obtendo a data atual

data_atual = datetime.datetime.now()

# Extraindo o nome do mês em português e formatando a primeira letra como maiúscula

nome_mes = data_atual.strftime('%B').capitalize()





# Caminho da pasta onde o arquivo será salvo

pasta_salvar = os.path.join("\\\Caminho da pasta onde o arquivo será salvo", nome_mes)



# Criando a pasta do mês (caso ainda não exista)

if not os.path.exists(pasta_salvar):

    os.makedirs(pasta_salvar)



# Obtendo o dia atual e criando a pasta do dia

dia_atual = datetime.datetime.now().day

pasta_dia_atual = os.path.join(pasta_salvar, str(dia_atual))



print("Procurando ou criando pasta destino")

# Criando a pasta do dia (caso ainda não exista)

if not os.path.exists(pasta_dia_atual):

    os.makedirs(pasta_dia_atual)

print("Salvando o arquivo novo")

# Salvando o dataframe como um arquivo xlsx na pasta do dia

caminho_arquivo_salvar = os.path.join(pasta_dia_atual, nome_arquivo)

df.to_excel(caminho_arquivo_salvar, index=False)



print("Arquivo salvo com sucesso:", caminho_arquivo_salvar)
