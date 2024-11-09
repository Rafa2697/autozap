import openpyxl #para ler planilha
from time import sleep
import pyautogui
import random
import pywhatkit

def fim(cont):
    meuNum = "+5513981317461"
    msg = f"Fim da transmissão, {cont} pessoas foram atingidas"
    pywhatkit.sendwhatmsg_instantly(meuNum, msg)
    
# webbrowser.open('https://web.whatsapp.com/')
#sleep(5)
#ler planilha e gaurdar informações sobre o nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('./alunos.xlsx')
pagina_clientes = workbook['Planilha1']

cont = 0

for linha in pagina_clientes.iter_rows(min_row=2):
    sleep_time = random.uniform(1, 3) #tempo de espera entre cada mensagem
    
    cont += 1 #contador de linhas
    if cont == 10:
        sleep(10)
    print(f'{cont} - {linha[0].value} - {linha[2].value} - {linha[3].value}')
    nome = linha[0].value
    primeiro_nome = nome.split()[0]
    telefone = str(linha[2].value)
    curso = linha[3].value
    midia = './img/outuvembro.jpeg'
    
    
    mensagem = f'''
Olá {primeiro_nome}, Gostaríamos de informar que o curso de *{curso}* está com as inscrições abertas. 
E tem mais: estamos com uma campanha especial de OUTUVEMBRO, onde você pode obter até *50% de desconto* nas mensalidades! Aproveite essa oportunidade e transforme sua carreira. 

Para mais informações, entre em contato conosco. 

*Atenciosamente, Faculdade Peruíbe*

Digite 1 para mais informações

Digite 2 não tem interesse
    '''

    #Preciso inserir a imagem aqui para ir junto com a mensagem
    pywhatkit.sendwhats_image(telefone,midia, mensagem)
    
    sleep(5)
    pyautogui.hotkey('ctrl', 'w')  # Atalho para fechar a aba
    sleep(sleep_time)

fim(cont)
print(f'Fim da transmissão, {cont} mensagens enviadas')