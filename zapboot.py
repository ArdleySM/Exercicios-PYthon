import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Abre o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(15)  # Tempo para você escanear o QR Code e carregar o WhatsApp Web

# Carrega a planilha do Excel
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['lista']

# Percorre as linhas da planilha a partir da segunda linha
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Cria a mensagem personalizada
    mensagem = (f"Olá {nome}, seu mão de vaca, pague o que me deve no dia "
                f"{vencimento.strftime('%d/%m/%Y')}. Meu PIX é esse 292.054.618-06. "
                "Consegui fazer minha primeira automação com Python!!!")

    # Gera o link para a mensagem no WhatsApp
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    # Abre o link no navegador
    webbrowser.open(link_mensagem_whatsapp)
    sleep(15)  # Espera o link carregar

    # Localiza a seta de envio e clica nela
    pyautogui.click(1852, 957, button='left')
    sleep(5)
    # Fecha a aba do WhatsApp Web
    pyautogui.hotkey('ctrl', 'w')
    sleep(5)