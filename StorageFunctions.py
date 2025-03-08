import os
import time
import win32gui
import winsound
import win32com.client
import win32con
import pyautogui
import pyperclip
from datetime import datetime
#from dotenv import load_dotenv

file_excel = "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Query.xlsm"

# ===================================================================================================================================================================

# Abrir archivo de Excel 
def OpenExcel():
    
    
    time.sleep(2)
    os.startfile(file_excel)
    print("Abriendo Excel...")
    
    open_excel = None
        
    while open_excel is None:
        ventana_excel = pyautogui.getWindowsWithTitle("Query - Excel")
        if ventana_excel:  
            open_excel = ventana_excel[0]  # Asigna la ventana encontrada
            open_excel.maximize()  # Maximiza la ventana
            print("Excel abierto y maximizado.")
            time.sleep(2)
        else:
            print("Esperando...")
            time.sleep(5)
    
    home_excel = None
    
    while home_excel is None:
        home_excel = pyautogui.locateOnScreen(
            'C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_excelcargado.png',
            grayscale=False,
            confidence=0.9
        )
        if home_excel is not None:
            print("Excel cargado y listo para usar.")
            time.sleep(2)
            break
        else:
            print("Esperando...")
            time.sleep(5)
            
# ===================================================================================================================================================================

# Guardar y cerrar archivo de Excel      
def SaveCloseExcel():
    
    save_excel = None
    
    pyautogui.hotkey("ctrl", "g")
    time.sleep(5)
    
    while save_excel is None:
        save_excel = pyautogui.locateOnScreen(
            "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_excelguardado.png",
            grayscale=False,
            confidence=0.9
        )
        if save_excel is not None:
            print("Excel guardado.")
            time.sleep(1)
            break
        else:
            print("Esperando que Excel guarde el archivo...")
            time.sleep(5)
    
    pyautogui.getWindowsWithTitle("Query - Excel")[0].close()
            
# ===================================================================================================================================================================

# Abrir CRM
def OpenCRM():
    
    time.sleep(1)
    pyautogui.press('win')
    time.sleep(1)
    pyautogui.write('Nuevo CRM')    
    time.sleep(1)
    pyautogui.press('enter')
    MakeWindowVisible('Control de Acceso Unificado CRM')

# ===================================================================================================================================================================

# Ingreso a CRM
def ConnectToCRM():
    
    #load_dotenv()
    
    #Ingreso de credenciales de ingreso de CRM
    pyautogui.press('tab', 2)
    time.sleep(1)
    pyautogui.write("")
    time.sleep(1)
    pyautogui.press('tab') 
    pyperclip.copy("")
    time.sleep(1)
    
    pyautogui.hotkey('ctrl','v')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(2)
    pyautogui.press('enter')
    
    popup_window = None
    
    while popup_window is None:
        print("Buscando ventana emergente...")
        popup_window = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_ventanaemergenteIngresoCRM.png', 
            grayscale = True,
            confidence=0.6
            )
    print("¡La ventana emergente de ingreso de CRM está presente!")
    
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('tab', 4)
    time.sleep(1)
    pyautogui.press('enter')
    
    homepage_crm = None
    
    while homepage_crm is None:
        homepage_crm = pyautogui.locateOnScreen(
            "C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_homepageCRM.png", 
            grayscale=True,
            confidence=0.05
            )
        if homepage_crm is not None:
            time.sleep(5)
            print("¡Ya está disponible la GUI de CRM!")
            break
        else:
            print("Esperando...")
            time.sleep(5)

# ===================================================================================================================================================================

# Busqueda activa de ventana CRM
def MakeWindowVisible(TargetWindow):
    
    time.sleep(2)
    count = 0
    encontrado = None
    print("Cantidad de ventanas activas: ",len(pyautogui.getAllWindows()))
    
    while encontrado is None:
        window_title = pyautogui.getActiveWindowTitle()
        if TargetWindow not in window_title:
            count += 1
            with pyautogui.hold('alt'):
                pyautogui.press('tab')
                for _ in range(0, count):                    
                    pyautogui.press('left')
        else:
            print("¡La ventana " + TargetWindow + " se encuentra visible y activa!")
            encontrado = window_title
        time.sleep(1)

# ===================================================================================================================================================================

# Modelado Datos de la Hora
def getCurrentDateAndTime():
    return datetime.now().strftime("%H:%M:%S")

# ===================================================================================================================================================================

# Modelado Datos de la Fecha
def formatDate(date):      
    return date.strftime("%d/%m/%Y")

# ===================================================================================================================================================================

# Modelado Datos de usuario
def get_username_os():
   return os.getenv("USERNAME")

# ===================================================================================================================================================================

# Presión de teclas
def pressingKey(key,times = 1):
    for i in range(times):
        # Presiona la tecla key varias veces con un intervalo de 0.6 segundos.
        time.sleep(0.6)
        pyautogui.press(key)
        
# ===================================================================================================================================================================

# Función de Minimizar todas las ventanas para mostrar el escritorio
def showDesktop():
    # Muestra el escritorio presionando la tecla de Windows y 'd' simultáneamente.
    pyautogui.hotkey('win', 'd')
    pyautogui.hotkey('win', 'd')
    
# ===================================================================================================================================================================

# Función de subrayar texto al editar en casilla
def selectToEnd():
        pyautogui.hotkey('end')
        time.sleep(1)
        pyautogui.doubleClick()
        time.sleep(1)

# ===================================================================================================================================================================

# Función que reproduce un tono para indicar inicio y finalizacion del BOOT        
def make_noise():
    duration = 1500  # milliseconds
    freq = 1140  # Hz
    winsound.Beep(freq, duration)
    winsound.Beep(freq, duration)
    winsound.Beep(freq, duration)

# ===================================================================================================================================================================

# Función para buscar ventana objetivo
def find_and_activate_window(target_window_title):
    window_list = []
    win32gui.EnumWindows(enum_window_callback, window_list)

    for hwnd, title in window_list:
        if target_window_title in title:
            print(f"Encontrada ventana: {title}, activándola...")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # Restaurar si está minimizada
            win32gui.SetForegroundWindow(hwnd)  # Activar la ventana
            return True

    print("No se encontró la ventana objetivo.")
    return False

# ===================================================================================================================================================================

# Función para recopilar titulos de ventanas activas
def enum_window_callback(hwnd, window_list):
    window_list.append((hwnd, win32gui.GetWindowText(hwnd)))

# ===================================================================================================================================================================

# Listar todas las ventanas visibles
def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
            print ( hex( hwnd ), win32gui.GetWindowText( hwnd ) )
    win32gui.EnumWindows( winEnumHandler, None )
    return 0

# ===================================================================================================================================================================

# Listar todas las ventanas activas
def list_all_active_windows():
    for x in pyautogui.getAllWindows():  
        print(x.title)
    while True:
        time.sleep(1)
        print(pyautogui.getActiveWindowTitle())

# ===================================================================================================================================================================

# Envio de notificacion lider de desarrollo
def terminateProcess(ProcessName):
    return os.system('taskkill /IM "' + ProcessName + '" /F')

# ===================================================================================================================================================================

# Función para enviar un correo electrónico a través de Outlook con asunto, destinatarios, cuerpo y archivos adjuntos.
def sendEmail(subject,to,cc,body,attachedFile,attachedFile2):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= subject
    newmail.To=to
    newmail.CC=cc 
    newmail.Body= body 
    if attachedFile != 'null':        
        newmail.Attachments.Add(attachedFile)     
    if attachedFile2 != 'null':
        newmail.Attachments.Add(attachedFile2)
    
    newmail.Send()

# ===================================================================================================================================================================

def CloseCRM():
    
    if(pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.5]")):
        pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.5]")[0].close()
        pyautogui.press('s')
        print("¡CRM cerrado exitosamente!")
    else:
        print("¡CRM no se encuentra activo!")
        return