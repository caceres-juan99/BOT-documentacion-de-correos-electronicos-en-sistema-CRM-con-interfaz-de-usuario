import time
import pyautogui
import pyperclip
import unidecode
import re
from StorageFunctions import (pressingKey, find_and_activate_window)

def searchUpdateOTPControl(incidentId, OTH, Subject, DisplayTo, DateTimeSent, Preview):
    
    crm_dashboard = crm_warning_message = crm_ot_blocked_message = mod_consulta_popup  = crm_ot_blocked_message = crm_edit_incident = crm_assign_user = crm_save_incident = crm_otp_saved_sucessfully = None 
    crmAttempts = 0 

    # Bucle de busqueda
    target_window_title = 'Sistema Avanzado'
    if find_and_activate_window(target_window_title):
        print(f"¡La ventana '{target_window_title}' es ahora visible y activa!")
        time.sleep(2)
    else:
        print(f"No se pudo encontrar la ventana '{target_window_title}'")

    # ============================================= CRM HOME PAGE =============================================   

    while crm_dashboard is None and crmAttempts < 5: 
        crm_dashboard = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_homepageCRM.png',
            grayscale=True,
            confidence=0.1
        )
        if crm_dashboard is not None:  # Si encuentra la imagen
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes")[0].maximize()
            print("¡Ya está disponible la GUI de CRM!")
        else:
            time.sleep(2)  # Esperar 1 segundo antes de reintentar
            crmAttempts += 1  # Incrementar el contador    
    
    pyautogui.click(pyautogui.center(crm_dashboard))
    time.sleep(2)
    pressingKey('f2')
    
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_ventanaconsultasCRM.png', 
            grayscale = True,
            confidence=0.3)            
        pressingKey('f2')
    print("¡La ventana de consultas de CRM está presente la GUI!")
    time.sleep(2)
    pyautogui.write(incidentId)
    time.sleep(2)
    pressingKey('enter')
    pressingKey('enter')
    
    #=========================================== FASE DE INGRESO ===========================================
    
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_mensajeAdvertenciaCRM.png', 
            grayscale = True,
            confidence=0.4)   
        crm_ot_blocked_message = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_OTblockedmmessageCRM.png', 
            grayscale = True,
            confidence=0.4)   
        time.sleep(2)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos para esperar alguna ventana emergente inesperada de OT bloqueada en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)            

    # Close Edit Incident View
    time.sleep(6)
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()      
    time.sleep(2)
    
    while crm_edit_incident is None:
        crm_edit_incident = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_botonEditarIncidenteCRM.png', 
            grayscale = True,
            confidence=0.9)   
        if crm_edit_incident is not None:
            crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
            pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)
            break
    

    while crm_assign_user is None and crmAttempts < 5:
        crm_assign_user = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_ventanaAsignarIncidenteUsuarioCRM.png', 
            grayscale = True,
            confidence=0.6)   
        time.sleep(2)
        print("Buscando ventana de asignación de usuario...")
        crmAttempts +=1            
        if crm_assign_user is not None:
            print("¡La ventana de asignación de incidente a usuario en CRM esta presente!")
            time.sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            time.sleep(2)
            break
    
    # =================================== FASE DE TRANSCRIPCIÓN DE CORREO ELECTRÓNICO ===================================
    
    pyautogui.press('tab')
    time.sleep(1)
    
    # Función para eliminar tildes
    def eliminar_tildes(texto):
        return unidecode.unidecode(texto)
    
    # Función para dar formato al texto (capitalización adecuada)
    def format_text_properly(text):
        
        # Divide el texto en líneas
        lines = text.split("\n")
        
        # Procesa cada línea de forma individual
        formatted_lines = [line.strip().title() if line.strip() else "" for line in lines]
        
        # Vuelve a unir las líneas con saltos de línea
        return "\n".join(formatted_lines)
    
    
    # Copiar y pegar el asunto
    pyperclip.copy(Subject)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    
    # Copiar y pegar el destinatario
    pyperclip.copy(DisplayTo)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    
    # Copiar y pegar la fecha/hora enviada
    pyperclip.copy(DateTimeSent)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press("enter", 2)
    time.sleep(1)
    
    
    # Texto original
    texto_original = Preview
    
    # Reemplazar caracteres no deseados
    texto_limpio = texto_original.replace("_x000D_", "").strip()
    
    # Eliminar saltos de línea adicionales
    texto_limpio = re.sub(r'(__|\n\s*\n|\n)+', lambda m: '\n\n' if '\n\n' in m.group(0) else '\n', texto_limpio)
    
    # Eliminar tildes
    texto_sin_tildes = eliminar_tildes(texto_limpio)
    
    # Aplicar formato al texto
    texto_formateado = format_text_properly(texto_sin_tildes)
    
    # Imprimir el texto formateado
    print(texto_formateado)
    time.sleep(1)
    
    # Escritura de textos largos en partes
    partes = [texto_formateado[i:i+100] for i in range(0, len(texto_formateado), 100)]
    
    # Escribir cada parte
    for parte in partes:
        pyautogui.write(parte)
        time.sleep(0.5)
    
    print("¡La documentación solicitada fue ingresada con éxito!")
    time.sleep(1)
    
    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen(
            "C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_botonGuardarincidenteCRM.png", 
            grayscale=True,
            confidence=0.9
        )
        if crm_save_incident is not None:
            crm_save_incident_x, crm_save_incident_y = pyautogui.center(crm_save_incident)
            pyautogui.click(crm_save_incident_x, crm_save_incident_y)
            break

    time.sleep(1) 
    
    while crm_warning_message is None and crm_assign_user is None:  
        print("Buscando Ventana de mensaje de advertencia o de asignación de incidente a usuario...")      
        crm_warning_message = pyautogui.locateOnScreen(
            "C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_mensajeAdvertenciaCRM.png", 
            grayscale = True,
            confidence=0.4)   
        crm_assign_user = pyautogui.locateOnScreen(
            'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_ventanaAsignarIncidenteUsuarioCRM.png', 
            grayscale = True,
            confidence=0.6)   
        time.sleep(1)
        
        if crm_warning_message is None:
            print("¡Ventana de mensaje de advertencia no esta presente!")
            pressingKey('enter')
                
            crm_assign_user = None  # reset variable
            crmAttempts = 0
            
            while crm_assign_user is None and crmAttempts < 5:
                crm_assign_user = pyautogui.locateOnScreen(
                    'C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_ventanaAsignarIncidenteUsuarioCRM.png', 
                    grayscale = True,
                    confidence=0.6)   
                time.sleep(2)
                print("Buscando ventana de asignación de usuario...")
                crmAttempts +=1            
                if crm_assign_user is not None:
                    print("¡La ventana de asignación de incidente a usuario en CRM esta presente!")
                    time.sleep(1)
                    pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
                    time.sleep(2)
                    break
            
    # Validate save OTP successfully  pop up is visible and on focus   
    while crm_otp_saved_sucessfully is None:
        crm_otp_saved_sucessfully = pyautogui.locateOnScreen(
            "C:/Users/jcacd/OneDrive - Nae/BPO/BOT BOTONES DE PROYECTO/Imagenes/img_incidenteguardadoCRM.png", 
            grayscale = True,
            confidence=0.9)   
        print("¡OTP guardada exitosamente!")
        time.sleep(1)
        pressingKey('enter')
        time.sleep(2)
    
    
    
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close() 
    time.sleep(2)
    
    # Agregado para pruebas
    pyautogui.press("n")
    time.sleep(2)
    
    
    return 0