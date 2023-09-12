import openpyxl
import pyautogui
import keyboard
import pyperclip

def type_with_accent_support(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']


try:
    for linha in vendas_sheet.iter_rows(min_row=2):
        pyautogui.click(950, 453, duration=0.01)
        type_with_accent_support(linha[0].value)
        pyautogui.click(942, 474, duration=0.01)
        type_with_accent_support(linha[1].value)
        pyautogui.click(933, 500, duration=0.01)
        type_with_accent_support(str(linha[2].value))
        pyautogui.click(1014, 526, duration=0.01)
        type_with_accent_support(linha[3].value)
        pyautogui.click(866, 552, duration=0.01)
        pyautogui.click(957, 582, duration=0.01)
        if keyboard.is_pressed('esc'):
            print("Programa interrompido pelo usuário.")
            break
except KeyboardInterrupt:
    print("Programa interrompido pelo usuário.")
