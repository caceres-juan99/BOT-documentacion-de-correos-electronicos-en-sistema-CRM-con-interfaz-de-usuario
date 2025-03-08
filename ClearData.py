import os
import time
import pyautogui
from StorageFunctions import (OpenExcel, SaveCloseExcel)

# ===================================================================================================================================================================
    
OpenExcel()

# ===================================================================================================================================================================
    
pyautogui.hotkey("alt", "o")
time.sleep(1)
pyautogui.hotkey("z", "e")
time.sleep(1)
pyautogui.hotkey("h", "u")
time.sleep(1)
pyautogui.hotkey("i")
time.sleep(1)
pyautogui.write("Datos")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
pyautogui.click(662, 506)
time.sleep(1)
pyautogui.hotkey("ctrl", "home")
time.sleep(1)
pyautogui.press("down")
time.sleep(1)
pyautogui.hotkey("ctrl", "shift", "down")
time.sleep(1)
pyautogui.press("delete")
time.sleep(1)
pyautogui.press("right", 3)
time.sleep(1)
pyautogui.hotkey("ctrl", "shift", "end")
time.sleep(1)
pyautogui.press("delete")
time.sleep(1)
print("Hoja de datos actualizada.")
pyautogui.hotkey("ctrl", "home")
time.sleep(1)

# ===================================================================================================================================================================

# Cerramos la ventana de Excel
SaveCloseExcel()
