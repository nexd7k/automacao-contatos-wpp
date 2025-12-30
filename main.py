import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

chrome_path = r'"C:/Program Files/Google/Chrome/Application/chrome.exe"'
webbrowser.get(f'{chrome_path} %s').open('https://web.whatsapp.com/')
sleep(15)

# Ler planilha e guardar informações sobre nome, telefone e mensagem 
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, mensagem
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = linha[2].value
    
    mensagem_def = f'Olá {nome}. Aqui o usuário pode colocar alguma mensagem fixa que sempre precisará enviar, alterar o que precisar na planilha no campo -> \n {mensagem}'

    # Criando links personalizados do whatsapp e enviando mensagens para cada cliente
    # com base nos dados da planilha
    try:
        if mensagem is None:
            print(f'Não havia mensagem para enviar para {nome}')
            continue
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem_def)}'
        webbrowser.get(f'{chrome_path} %s').open(link_mensagem_whatsapp)
        sleep(10)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
    

