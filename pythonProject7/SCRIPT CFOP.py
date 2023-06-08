from BuscarDados import BuscarDados
import time
import pyautogui

WORKBOOK = BuscarDados.ler_planilha("\\Desktop\\Ressarcimento\\NOTAS.xlsx")

Excel = WORKBOOK['2101']

max_row = Excel.max_row
max_column = Excel.max_column
linha_atual = 2

time.sleep(2)


def chamar_funcao(Escrever):
    for i in range(2, max_row + 1):
        CNPJ = str(Escrever.cell(i, 1).value)
        Num_Doc = str(Escrever.cell(i, 2).value)
        Serie = str(Escrever.cell(i, 3).value)
        Dat_Doc = str(Escrever.cell(i, 4).value)
        Dat_Ent = str(Escrever.cell(i, 5).value)
        Vlr_Bs_Calc_IPI = str(Escrever.cell(i, 6).value)
        Vlr_IPI = str(Escrever.cell(i, 7).value)

        pyautogui.write(CNPJ);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(Num_Doc);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(Serie);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(Dat_Doc);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(Dat_Ent);pyautogui.press('tab');time.sleep(1)
        pyautogui.press('Down', presses=132);time.sleep(2)
        pyautogui.press('tab');time.sleep(1)
        pyautogui.write(Vlr_Bs_Calc_IPI);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(Vlr_IPI);pyautogui.press('tab');time.sleep(1)
        pyautogui.write(Vlr_IPI);pyautogui.press('tab');time.sleep(1)
        pyautogui.press('enter');time.sleep(2)
        pyautogui.press('tab');time.sleep(1)
        pyautogui.press('tab');time.sleep(1)
        pyautogui.press('enter');time.sleep(1)


chamar_funcao(Excel)