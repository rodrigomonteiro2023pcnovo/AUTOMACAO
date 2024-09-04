import pandas as pd
from time import sleep
from playwright.sync_api import sync_playwright
import customtkinter as ctk

# 1 ler o arquivo excel
escala = 'planilha.xlsx'
df = pd.read_excel(escala)
site = "https://www.expresso.pe.gov.br/login.php"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto(site)
    sleep(1)
    page.fill("//*[@id='user']", "rodrigo.monteiro1")
    sleep(.5)
    page.fill("//*[@id='passwd']", "Rm@29210306")
    sleep(.5)
    page.click("//*[@id='conteudo_login']/div[3]/input")
    sleep(1)
    page.click("//*[@id='tableDivAppbox']/tbody/tr/td/table/tbody/tr[1]/td[2]/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td/ul/li[1]/div/a")
    sleep(1)
    page.click("//*[@id='folders_tbl']/tbody/tr[1]/td/table/tbody/tr[2]/td/div/img")
    
 # Loop Enviando os emails para cada linha do arquivo excel
for index, row in df.iterrows():

    #clicar em nova mensagem
    page.click("//*[@id='folders_tbl']/tbody/tr[1]/td/table/tbody/tr[2]/td/div/img")
    sleep(1)
'''
    #digitar o destinatario do email
    page.fill("//*[@id='to_1']", row["Email"])
    sleep(1)

    #digitar o nome no assunto
    page.fill("//*[@id='subject_1']", row["Nome"])
    sleep(1)

    #digitar a matricula no corpo do texto
    page.fill("/html/body", row["Matricula"])
    sleep(1)

    #clicar em enviar
    page.click("//*[@id='send_button_1']")
    sleep(1)'''
 