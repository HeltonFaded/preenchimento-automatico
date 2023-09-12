import openpyxl
import pyautogui
import keyboard
import pyperclip

# Configurações
coordenadas_clicar = [(950, 453), (942, 474), (933, 500), (1014, 526), (866, 552), (957, 582)] #alterar conforme a posição do mouse
tecla_interrupcao = 'esc'

# Função para inserir texto com suporte a acentos
def type_with_accent_support(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')

try:
    # Abrir a planilha Excel
    workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
    vendas_sheet = workbook['vendas']

    for linha in vendas_sheet.iter_rows(min_row=2):
        for coord in coordenadas_clicar:
            pyautogui.click(coord[0], coord[1], duration=0.01)
            type_with_accent_support(str(linha[coordenadas_clicar.index(coord)].value))
        if keyboard.is_pressed(tecla_interrupcao):
            print("Programa interrompido pelo usuário.")
            break

except FileNotFoundError:
    print("Arquivo 'vendas_de_produtos.xlsx' não encontrado.")
except Exception as e:
    print(f"Ocorreu um erro: {str(e)}")
