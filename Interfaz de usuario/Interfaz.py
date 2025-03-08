import tkinter as tk
from tkinter import PhotoImage
from PIL import Image, ImageTk
import subprocess
import threading

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Interfaz de usuario NECTIM")

# Establecer el icono de la ventana
ventana.iconbitmap("C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Interfaz de usuario/Iconos/Icono de ventana.ico")  # Cambia esta ruta a tu archivo .ico

# Ajustar el tamaño de la ventana
ventana.geometry("400x400")
ventana.resizable(False, False)  # Evitar que la ventana cambie de tamaño

# Funciones para los botones
def ClearData_action():
    def task():
        subprocess.run(["python", "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/ClearData.py"])
    threading.Thread(target=task).start()
    
def DataModeling_action():
    def task():
        subprocess.run(["python", "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/DataModeling.py"])
    threading.Thread(target=task).start()

def main_action():
    def task():
        subprocess.run(["python", "C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/main.py"])
    threading.Thread(target=task).start()

# Cargar la imagen de fondo
fondo_image = Image.open("C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Interfaz de usuario/Iconos/Fondo de interfaz.png")
fondo_image = fondo_image.resize((400, 400), Image.Resampling.LANCZOS)
fondo_photo = ImageTk.PhotoImage(fondo_image)

# Crear un widget de etiqueta para mostrar la imagen de fondo
label_fondo = tk.Label(ventana, image=fondo_photo)
label_fondo.place(x=0, y=0, relwidth=1, relheight=1)

# Crear el logo de la compañía con subtítulo
logo_frame = tk.Frame(ventana, bg="white")

# Crear etiquetas con fondo igual al del contenedor
logo_label = tk.Label(
    logo_frame, 
    text="NECTIM", 
    font=("Times", 40, "bold"), 
    fg='#1babff', 
    bg=logo_frame.cget("bg"),  # Igualar el fondo al del marco
    padx=20,  # Ancho extra horizontal
    pady=5   # Altura extra vertical
)

subtitulo_label = tk.Label(
    logo_frame, 
    text="DOCUMENTACIÓN ESTÁNDAR", 
    font=("Arial", 8, "bold"), 
    fg="black", 
    bg=logo_frame.cget("bg"),  # Igualar el fondo al del marco
    pady=10
)


logo_label.pack()
subtitulo_label.pack()
logo_frame.pack(pady=30)

# Crear un frame transparente para los botones
frame_botones = tk.Frame(ventana, bg='', relief='flat')
frame_botones.pack(pady=0)

# Crear botones con estilo
boton_ClearData = tk.Button(
    frame_botones,
    text="Clear Data",
    command=ClearData_action,
    bg='#dfdfdf', 
    fg='black',
    font=("Lato", 10, "bold"),
    width=14,
    height=2,
    relief="solid"
)
boton_DataModeling = tk.Button(
    frame_botones,
    text="Data Modeling",
    command=DataModeling_action,
    bg='#dfdfdf', 
    fg='black',
    font=("Lato", 10, "bold"),
    width=14,
    height=2,
    relief="solid"
)
boton_main = tk.Button(
    frame_botones,
    text="Main",
    command=main_action,
    bg='#dfdfdf', 
    fg='black',
    font=("Lato", 10, "bold"),
    width=14,
    height=2,
    relief="solid"
)

# Colocar botones en el frame

boton_ClearData.pack(pady=5)
boton_DataModeling.pack(pady=5)
boton_main.pack(pady=5)

# Iniciar el bucle principal de Tkinter
ventana.mainloop()