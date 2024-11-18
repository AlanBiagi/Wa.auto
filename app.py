"""
FEITO PARA ENVIAR AS MESAGENS DE UMA EXTENSA PLANILHA
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(15)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('nome-da-planilha.xlsx')
pagina_clientes = workbook['Nome-da-pagina']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, codico da promoção
    nome = linha[0].value #pega o valor da primeira coluna 
    telefone = linha[1].value #pega o valor da segunda coluna 
    cod_promo = linha [2].value #pega o valor da terceira coluna 
    
    mensagem = f'Insira a mensagem aqui '

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(4)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(6)
    except: #Caso o numero seja invalido, cria um arquivo em CSV e passa o nome e telefone em que não foi possivel entrar em contatt
        print(f'Não foi possível enviar mensagem para {nome} {telefone}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
    
