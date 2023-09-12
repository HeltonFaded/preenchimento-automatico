# preenchimento-automatico
aplicação python que usa as informações do mouse para preencher automaticamente um cadastro com base em uma planilha excel.

Este é um script Python que automatiza a entrada de dados em uma planilha Excel utilizando as bibliotecas openpyxl, pyautogui, keyboard e pyperclip. O script lê dados de uma planilha Excel e os insere em campos de um aplicativo, simulando a entrada manual de dados.

## Pré-requisitos

Certifique-se de ter as seguintes bibliotecas Python instaladas:

- openpyxl: Para manipulação de planilhas Excel.
- pyautogui: Para simular a digitação e cliques do mouse.
- keyboard: Para detectar se o usuário pressionou a tecla "esc" para interromper o programa.
- pyperclip: Para copiar e colar o texto na área de transferência.

Você pode instalá-las usando o seguinte comando:

```bash
pip install openpyxl pyautogui keyboard pyperclip
