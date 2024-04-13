import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('chamados.xlsx')
sheet_produtos = workbook['chamados']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Inserir chamado novamente
    pyautogui.click(2377,816,duration=2)

    # Data
    data = linha[0].value
    pyperclip.copy(data)
    pyautogui.click(1938,361,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Chamadomesa despacho
    numerochamado = linha[1].value
    pyperclip.copy(numerochamado)
    pyautogui.click(1935,453,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Atendente
    atendente = linha[2].value
    pyperclip.copy(atendente)
    pyautogui.click(1940,546,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Solicitante
    solicitante = linha[3].value
    pyperclip.copy(solicitante)
    pyautogui.click(1937,638,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Local
    local = linha[4].value
    pyperclip.copy(local)
    pyautogui.click(1936,729,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Motivo
    motivo = linha[5].value
    pyperclip.copy(motivo)
    pyautogui.click(1935,823,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Botão próximo
    pyautogui.click(1962,999,duration=1)
    sleep(3)

    # Inserir chamado novamente
    pyautogui.click(2377,816,duration=2)

