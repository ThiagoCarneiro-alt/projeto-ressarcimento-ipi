import pyautogui
import time

#ABERTURA DA DECLARAÇÃO
time.sleep(5)
pyautogui.doubleClick(622, 254);time.sleep(5)
pyautogui.doubleClick(19, 56);time.sleep(2)
pyautogui.write('02082021');time.sleep(2)
pyautogui.press('tab', presses=2);time.sleep(2)
pyautogui.write('05559838000133');time.sleep(2)
pyautogui.press('tab');pyautogui.write('Outra Qualificação');time.sleep(2)
pyautogui.press('tab');pyautogui.press('DOWN');time.sleep(2);time.sleep(2)

#HABILITAÇÃO DO PREENCHIMENTO DA DECLARAÇÃO
pyautogui.press('tab', presses=2);time.sleep(2)
pyautogui.write('RE');pyautogui.press('tab');pyautogui.write('Não');time.sleep(2);time.sleep(2)
pyautogui.press('tab', presses=4);time.sleep(2)
pyautogui.write('000133');pyautogui.press('tab');time.sleep(2)
pyautogui.write('1');time.sleep(2)
pyautogui.press('tab', interval=0.25);pyautogui.press('DOWN', presses=4);time.sleep(2)
pyautogui.press('tab', interval=0.25);pyautogui.press('space');time.sleep(2)
pyautogui.press('tab', interval=0.25);pyautogui.press('space');time.sleep(2)
pyautogui.press('tab', interval=0.25);pyautogui.press('space');time.sleep(2)
pyautogui.press('tab', interval=0.25);pyautogui.press('space');time.sleep(2)
pyautogui.doubleClick(702, 409);time.sleep(1)
pyautogui.doubleClick(591, 554);time.sleep(1)
pyautogui.doubleClick(771, 449);time.sleep(1)

#Ficha de dados iniciais
pyautogui.write('Ruplast Industria Comercio LTDA');time.sleep(2)
pyautogui.press('tab', presses=3);time.sleep(2)
pyautogui.press('DOWN', presses=2);pyautogui.press('tab');time.sleep(2)
pyautogui.write('033');time.sleep(1)
pyautogui.press('tab');time.sleep(2)
pyautogui.press('enter');time.sleep(2)
pyautogui.write("1212");time.sleep(2)
pyautogui.press('tab');time.sleep(1)
pyautogui.write('12121212');time.sleep(1)
pyautogui.press('tab');time.sleep(1)
pyautogui.write('7');time.sleep(1)

#FICHA DADOS DO RESPONSÁVEL DA PESSOA JURÍDICA PERANTE A RFB
pyautogui.click(71, 119);time.sleep(1)
pyautogui.write('MARCELLO JAROSLAVSKY RUSHANSKY');pyautogui.press('tab');time.sleep(2)
pyautogui.write('00789914476');pyautogui.press('tab');time.sleep(2)
pyautogui.press('enter');time.sleep(1)
pyautogui.press('tab', presses=7);time.sleep(2)
pyautogui.write('Pedro Henrique Martins Barros');pyautogui.press('tab');time.sleep(2)
pyautogui.write('07337822480');pyautogui.press('tab');time.sleep(2)
pyautogui.write('025938/O');time.sleep(1)
pyautogui.press('tab')
pyautogui.write('PE');time.sleep(2)
pyautogui.press('tab', presses=6);time.sleep(2)
pyautogui.write('Pedro.Barros@BWA.GLOBAL');time.sleep(1)
pyautogui.doubleClickclick(32, 719);time.sleep(1)

#FICHA DE RESSARCIMENTO DE IPI

#JANEIRO - ENTRADA




