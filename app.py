""""""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com')
sleep(5)

#Descrever os passos manuais e depois transformar em código 
#Ler planilha e guardar informações

workbook= openpyxl.load_workbook('Protótipo.app.xlsx')
pagina_clientes= workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Oi {nome} seu boleto vence no dia {vencimento}. Favor pagar'
#Criar links personalizado com base na planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey('ctrl','w')
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
