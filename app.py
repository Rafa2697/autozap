import openpyxl #para ler planilha
from urllib.parse import quote #para codificar caracteres especiais
import webbrowser #para abrir links
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
import random

webbrowser.open('https://web.whatsapp.com/')
sleep(5)
#ler planilha e gaurdar informações sobre o nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('./clientes.xlsx')
pagina_clientes = workbook['Sheet1']
driver = webdriver.Chrome()



for linha in pagina_clientes.iter_rows(min_row=2):
    sleep_time = random.uniform(1, 30) #tempo de espera entre cada mensagem
    sleep(sleep_time)
    #nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    mensagem = f'Olá {nome}, seu vencimento é em {vencimento.strftime("%d/%m/%Y")}'
    #criar link do whatsapp
    link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem)
    sleep(5)
    #Preciso teclar o enter para enviar a mensagem digitada no whatsapp
    pyautogui.press('enter')  # Aqui você escolhe a tecla
    sleep(2)  # Tempo para trocar de aba/janela
    pyautogui.hotkey('ctrl', 'w')  # Atalho para fechar a aba

#Criar links personalizados do whatsapp e eviar mensagens para cada cliente
#com base nos dados da planilha
