import openpyxl
import pyautogui

workbook=openpyxl.load_workbook("D:/Programação/Python/Automacao/vendas.xlsx")

paginas_vendas = workbook['Planilha1']

for linha in paginas_vendas.iter_rows(min_row=2):
    pyautogui.click(145,102,duration=0.1)
    pyautogui.write(linha[0].value)

    pyautogui.click(147,123,duration=0.01)
    pyautogui.write(linha[1].value)
    
    pyautogui.click(144,145,duration=0.01)
    pyautogui.write(str(linha[2].value))

    pyautogui.click(148,166,duration=0.01)
    pyautogui.write(linha[3].value)

    pyautogui.click(197,201,duration=0.01)

    