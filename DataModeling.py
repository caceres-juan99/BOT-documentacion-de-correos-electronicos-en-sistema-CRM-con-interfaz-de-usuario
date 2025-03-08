import os
import time
import openpyxl
import pyautogui
import pandas as pd
from datetime import datetime
#from dotenv import load_dotenv
from StorageFunctions import (OpenExcel, SaveCloseExcel)

file_excel = "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Query.xlsm"

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
pyautogui.write("Mail")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)

# ===================================================================================================================================================================  

pyautogui.hotkey("ctrl", "h")
print("Actualizando consultas y conexiones...")
time.sleep(1)

macro1 = None

while macro1 is None:
    macro1 = pyautogui.locateOnScreen(
        "C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_MACROActualizarConexionesejecutada.png",
        grayscale=False,
        confidence=0.25)
    
    if macro1 is not None:
        time.sleep(5)
        pyautogui.press("enter")  
        print("¡Actualización de consultas y conexiones exitosa!")
        time.sleep(5)
        break 
    else:
        print("Esperando...")
        time.sleep(5)

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

# ===================================================================================================================================================================

pyautogui.hotkey("ctrl", "n")
print("Buscando OTPs...")
time.sleep(1)

macro2 = None

while macro2 is None:
    macro2 = pyautogui.locateOnScreen(
        "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_MACROejecutadauniversal.png",
        grayscale=True,
        confidence=0.4)
    if macro2 is not None:
        time.sleep(5)
        pyautogui.press("enter")
        print("¡Busqueda y extracción de OTPs exitosa!")
        time.sleep(5)  
        break
    else:
        print("Esperando...")
        time.sleep(5)

# ===================================================================================================================================================================  
    
pyautogui.hotkey("ctrl", "f")
print("Buscando información relacionada con OTPs...")
time.sleep(1)

macro3 = None

while macro3 is None:
    macro3 = pyautogui.locateOnScreen(
        "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_MACROejecutadauniversal.png",
        grayscale=True,
        confidence=0.4)
    if macro3 is not None:
        time.sleep(5)
        pyautogui.press("enter")
        print("¡Busqueda y extracción de información relacionada a OTPs exitosa!")
        time.sleep(5) 
        # Primera secuencia de teclas
        pyautogui.hotkey("ctrl", "home")
        time.sleep(2)
        pyautogui.hotkey("ctrl", "shift", "space")
        time.sleep(2)
        pyautogui.press("alt")
        time.sleep(2)
        pyautogui.press("o")
        time.sleep(2)
        pyautogui.hotkey("u", "j")
        time.sleep(2)
        
        for _ in range(1):
            pyautogui.press("alt")
            time.sleep(2)
            pyautogui.press("o")
            time.sleep(2)
            pyautogui.hotkey("u", "j")
            time.sleep(2)
        break
    else:
        print("Esperando...")
        time.sleep(5)

# ===================================================================================================================================================================  


    
# ===================================================================================================================================================================  

# Cerramos la ventana de Excel
SaveCloseExcel()
time.sleep(5)

# ===================================================================================================================================================================  

# Lectura del excel
df = pd.read_excel(file_excel, sheet_name='Datos', engine='openpyxl')

# Imprimir los nombres de las columnas para verificar
print("Nombres de las columnas en el DataFrame:")
print(df.columns)

# Filtrar el DataFrame para excluir filas donde 'VAL_OTP' es 'NO'
df = df[df['VAL_OTP'] != 'NO']

# Nombre de la columna que contiene las fechas y horas
nombre_columna_fecha_hora = 'DateTimeSent'

# Convertir la columna de fechas y horas del DataFrame a un formato datetime
df[nombre_columna_fecha_hora] = pd.to_datetime(df[nombre_columna_fecha_hora], dayfirst=True)

# Filtrar el DataFrame por una fecha y hora específica
definir_fecha_hora = datetime(2023, 1, 1, 0, 0, 0)  # (1 de enero de 2023, a las 00:00:00) Cambiar esta fecha y hora según sea necesario 
df_filtrado = df[df[nombre_columna_fecha_hora] >= definir_fecha_hora]

# Seleccionar solo las columnas requeridas
columnas_requeridas = ['VAL_OTP','VAL_OTH','Subject', 'DisplayTo', 'DateTimeSent', 'Preview','PreviewFinal']
df_filtrado = df_filtrado.filter(items=columnas_requeridas)

# Insertar la columna 'Status' entre las columnas dos y tres
df_filtrado.insert(2, 'Status', '')  # Insertar en la posición 2

# Guardar el resultado en un nuevo archivo Excel
ruta_destino = "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Insumos/Input BOT 001 V1.xlsx"
df_filtrado.to_excel(ruta_destino, index=False)

print(f"Archivo guardado en {ruta_destino}")

time.sleep(1)