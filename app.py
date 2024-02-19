from urllib.parse import quote
import openpyxl
import webbrowser
import pyautogui
from time import sleep

planilha = openpyxl.load_workbook('teste.xlsx')
desculpa = planilha['Planilha1']

for linha in desculpa.iter_rows(min_row=2):
    msg = linha[0].value
    telefone = linha[1].value

    mensagem = f'{msg}'

    link_mensagem_zap = f'https://web.whatsapp.com/send?phone={
        telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_zap)
    sleep(9)
    try:
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(4)
        pyautogui.click(seta[0], seta[1])
        sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
    except:
        print(f'n deu')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'Erro \n')
