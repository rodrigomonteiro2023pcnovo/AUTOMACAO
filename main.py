import pandas as pd
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


# 1 ler o arquivo excel
escala = 'planilha.xlsx'
df = pd.read_excel(escala)
site = "https://www.expresso.pe.gov.br/login.php"
usuario = "rodrigo.monteiro1" 
senha = "Rm@29210306"

# Abrindo o navegador
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
driver.get(site)

# localizando o campo de usuario e senha para fazer o login
campo_usuario = driver.find_element(By.XPATH, '//*[@id="user"]')
sleep(1)
campo_senha = driver.find_element(By.XPATH, '//*[@id="passwd"]')
sleep(1)
campo_usuario.send_keys(usuario)
sleep(1)
campo_senha.send_keys(senha)
botao = driver.find_element(By.XPATH, '//*[@id="conteudo_login"]/div[3]/input').click()
botao2 = driver.find_element(By.XPATH, '//*[@id="expressoMail12id"]').click()
sleep(1)

# Loop Enviando os emails para cada linha do arquivo excel
for index, row in df.iterrows():

    #clicar em nova mensagem
    nova_mensagem = driver.find_element(By.XPATH, '//*[@id="folders_tbl"]/tbody/tr[1]/td/table/tbody/tr[2]/td/div/span').click()
    sleep(1)

    #digitar o destinatario do email
    destinatario = driver.find_element(By.XPATH, '//*[@id="to_1"]').send_keys(row['Email'])
    sleep(1)

    #digitar o nome no assunto
    assunto = driver.find_element(By.XPATH, '//*[@id="subject_1"]').send_keys(row['Nome'])
    sleep(1)

    #clicar em enviar
    enviar = driver.find_element(By.XPATH, '//*[@id="send_button_1"]').click()
    sleep(1)
   


