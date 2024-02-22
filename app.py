import datetime
import os
import smtplib
import time
from email import encoders
from email.mime.base import MIMEBase
from selenium import webdriver
import subprocess
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from docx import Document
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

#acessar site da oab e preencher o formulário de login
oab = 93876
with open ('psswrd.txt') as file:
     senha = file.readlines()
     senha_oab = senha[1]

chrome = webdriver.Chrome()
chrome.get('https://www2.oabsp.org.br/asp/dotnet/LoginSite/LoginMain.aspx?ReturnUrl=https:%2f%2fwww2.oabsp.org.br%2fasp%2fdotnet%2fLoginSite%2fAcessoRestrito%2fDefaultMain.aspx')

time.sleep(5)

oab_input = chrome.find_elements(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ctrLogin_UserName"]')[0]
oab_input.send_keys(oab)

senha_input = chrome.find_elements(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ctrLogin_Password"]')[0]
senha_input.send_keys(senha_oab)

time.sleep(10)

botao = chrome.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ctrLogin_Login"]')
chrome.execute_script("arguments[0].click();", botao)

time.sleep(5)

print('Fez login')

#entrar na aba de intimações
intimacoes = chrome.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder2_pnIcones"]/span/table/tbody/tr[3]/td[1]/a/img')
chrome.execute_script("arguments[0].click();", intimacoes)

print('Entrando em intimações')

time.sleep(10)

if (chrome.find_element(By.XPATH, '//*[@id="ctl00_btnPopupAvisoFechar"]')):
    chrome.find_element(By.XPATH, '//*[@id="ctl00_btnPopupAvisoFechar"]').click()

time.sleep(5)

#pegar e preencher as datas
dia_de_hoje = datetime.date.today()
desconto_dias = dia_de_hoje.day - 5
dias_atras = dia_de_hoje.replace(day=desconto_dias).strftime("%d/%m/%Y")
dia_de_hoje = dia_de_hoje.strftime("%d/%m/%Y")

data_inicio = chrome.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtSelDataInicio"]')
data_inicio.clear()
data_inicio.send_keys(dias_atras)

data_fim = chrome.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtSelDataFim"]')
data_fim.clear()
data_fim.send_keys(dia_de_hoje)

#pesquisar intimações
pesquisar = chrome.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnConsultar"]')
chrome.execute_script("arguments[0].click();", pesquisar)

print(f'Colocou e pesquisou as datas {dias_atras} e {dia_de_hoje}')

time.sleep(10)

#associações todas as intimações
ver_todos = chrome.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_Button2"]')
chrome.execute_script("arguments[0].click();", ver_todos)

print('Clicou em ver todos')

time.sleep(10)

#exibir todas as intimações
exibir_todas_paginas = chrome.find_element(By.XPATH, '//*[@title="Atenção: Exibir todas as publicações em uma única tela pode deixar o carregamento da página mais lento"]')
chrome.execute_script("arguments[0].click();", exibir_todas_paginas)

print('Clicou em exibir todas as páginas')

time.sleep(30)

###################################
def acha_btn():
    try:
        print('Tentando clicar nos três pontos')
        wait = WebDriverWait(chrome, 10)
        tres_pts = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@title="Exibir conteúdo completo"]')))
        chrome.execute_script("arguments[0].click();", tres_pts)
        print('Clicou nos três pontos')
        time.sleep(15)
        acha_btn()
    except StaleElementReferenceException:
        print('Não achou os três pontos')
        pass
    except NoSuchElementException:
        print("Elemento não encontrado")
        pass
        
def filtro():
    retorno = []
    todas_publicacoes = chrome.find_elements(By.TAG_NAME, 'tr')
    for elemento in todas_publicacoes:
        if str(oab) in elemento.text:
            retorno.append(elemento.text)

    return retorno

def converter_docx_para_pdf(arquivo_docx, arquivo_pdf):
    # Comando para converter usando o LibreOffice
    comando = ['libreoffice', '--convert-to', 'pdf', '--outdir', '.', arquivo_docx]

    # Executar o comando
    subprocess.run(comando)

acha_btn()
publicações_filtradas = filtro()

if os.path.exists('intimações_novas.docx'):
    os.remove('intimações_novas.docx')
    print('Arquivo removido')
    
doc = Document()
for publicação in publicações_filtradas:
    doc.add_paragraph(publicação)

doc.save('intimações_novas.docx')
print('Arquivo criado')

converter_docx_para_pdf('intimações_novas.docx', 'intimações_novas.pdf')

pdf = 'intimações_novas.pdf'
#################################################

server = 'smtp.gmail.com'
port = 587

email_envio = 'emailautomatico.oab@gmail.com'
with open ('psswrd.txt') as file:
     senha = file.readlines()
     senha_envio = senha[0]
           
destinatario = 'gorcuchomarcelo@gmail.com'
arquivo_anexo = 'intimações_novas.docx'

corpo = f'Segue em anexo as intimações do dia {dias_atras} até {dia_de_hoje}'
msg = MIMEMultipart()
msg['From'] = email_envio
msg['To'] = destinatario
msg['Subject'] = 'Intimações OAB'
msg.attach(MIMEText(corpo, 'plain'))

with open(pdf, 'rb') as anexo:
    part = MIMEBase('application', 'pdf')
    part.set_payload(anexo.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{pdf}"')
    msg.attach(part)

# Iniciando conexão SMTP
with smtplib.SMTP(server, port) as smtp:
    smtp.starttls()
    smtp.login(email_envio, senha_envio)
    smtp.send_message(msg)

print(f"Email enviado para {destinatario}")
os.remove('intimações_novas.pdf')
