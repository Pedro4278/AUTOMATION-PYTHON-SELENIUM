import pyautogui
import time
#774,738 - 774,738 - 574,390 - 472,396

#abrir a internet
pyautogui.moveTo(775,753,duration=2)
pyautogui.click()
time.sleep(5)

#clicar dinheirinho
pyautogui.moveTo(23,334,duration=2)
pyautogui.click()
time.sleep(5)

#repetir os comandos
for c in range(5):

#abaixar a pagina
 if c == 0:
    for c in range (10):
        pyautogui.moveTo(1304,708,duration=2)
        pyautogui.click()
        time.sleep(0.5)
 else:
    for c in range (1):
        pyautogui.moveTo(1304,708,duration=2)
        pyautogui.click()
        time.sleep(0.5)

    #clicar no negocio
    pyautogui.moveTo(555,202,duration=2)
    pyautogui.click()
    time.sleep(5)

    #botão vermelho
    pyautogui.moveTo(1217,147,duration=2)
    pyautogui.click()
    time.sleep(5)

    # abrir Nota
    pyautogui.moveTo(323,217,duration=2)
    pyautogui.click()
    time.sleep(5)

    #adiou a compra
    pyautogui.moveTo(397,229,duration=2)
    pyautogui.click()
    time.sleep(5)

    # finalizar
    pyautogui.moveTo(1074,158,duration=2)
    pyautogui.click()
    time.sleep(5)

    #voltar
    pyautogui.moveTo(85,95,duration=2)
    pyautogui.click()
    time.sleep(5)




