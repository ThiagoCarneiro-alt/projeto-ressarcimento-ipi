from BuscarDados import BuscarDados
import time
import pyautogui

WORKBOOK = BuscarDados.ler_planilha("/Desktop/SELECTO/NOTAS.xlsx")

Excel = WORKBOOK['1101']

max_row = Excel.max_row
max_column = Excel.max_column
linha_atual = 2

time.sleep(2)


def chamar_funcao(Escrever):
    for i in range(2, max_row + 1):
        CNPJ = str(Escrever.cell(i, 1).value)
        NUM_DOC = str(Escrever.cell(i, 2).value)
        SERIE = str(Escrever.cell(i, 3).value)
        DATA_DOC = str(Escrever.cell(i, 4).value)
        DATA_ENT = str(Escrever.cell(i, 5).value)
        VLR_BASE_IPI = str(Escrever.cell(i, 6).value)
        VLR_IPI = str(Escrever.cell(i, 7).value)

        pyautogui.write(CNPJ);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(NUM_DOC);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(SERIE);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(DATA_DOC);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(DATA_ENT);pyautogui.press('tab');time.sleep(1)
        pyautogui.write('1');time.sleep(2)
        pyautogui.press('tab');time.sleep(2)
        pyautogui.write(VLR_BASE_IPI);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(VLR_IPI);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(VLR_IPI);pyautogui.press('tab');time.sleep(1)
        pyautogui.press('enter');time.sleep(2)
        pyautogui.click(758, 117);time.sleep(2)

chamar_funcao(Excel)





