import time

import pyautogui

'''
Aulas 1 - (-114,593) 
      2 - (-1151, 790)
      
professores 
      1 - (-106,481) 
      2 - (-1164, 821)
x = 1809 y = 599
x = 1821 y = 483
'''


#crtl+f2

time.sleep(5)
for i in range(0,9999999999999999):
    pyautogui.click(1821,483)
    time.sleep(5)
    pyautogui.click(1004, 680)
    time.sleep(2)


'''x, y = pyautogui.position()
print("Posicao atual do mouse:")
print("x = "+str(x)+" y = "+str(y))'''