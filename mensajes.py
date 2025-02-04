import pandas as pd
import webbrowser as web
import pyautogui as pg
import time
import random
import subprocess

# URL con el formato correcto para descargar CSV desde Google Sheets
csv_url = "https://docs.google.com/spreadsheets/d/1fX9CZbrEdLE6e7OJC2wuctOLPoN1sc2XjV3dLjYfVuk/export?format=csv"
data = pd.read_csv(csv_url)

print(data.head())

# Ruta de Google Chrome (ajusta según tu sistema)
chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe'

for i in range(len(data)):
    celular = str(data.loc[i, 'Celular'])  # Convertir a string
    nombre = data.loc[i, 'Nombre']
    maquina = data.loc[i, 'Maquina']
    concesionario = data.loc[i, 'Concesionario']
    
    # Crear mensaje personalizado
    mensajeConcesionario = f"Hola concesionario {concesionario}, estoy probando un codigo para el envio masivo de mensajes a varias personas, no te voy a hackear ni nada. Hola {nombre}! Gracias por comprar una Sembradora {maquina} con nosotros 🙌, ojala no se te rompa🫣"
    mensajeUsuario = f"Hola {nombre}! Gracias por comprar una Sembradora {maquina} con nosotros 🙌, ojala no se te rompa🫣"
    
    # Abrir WhatsApp Web en Google Chrome
    url = f"https://web.whatsapp.com/send?phone={celular}&text={mensajeUsuario}"
    subprocess.Popen([chrome_path, url])
    
    # Esperar un tiempo aleatorio para cargar la página
    time.sleep(random.randint(10, 15))
    
    # Simular la interacción para enviar el mensaje
    pg.press('tab')  
    time.sleep(random.uniform(1.5, 3.0))
    pg.press('enter')  
    time.sleep(random.uniform(2.0, 4.0))  
    
    # Cerrar la pestaña
    pg.hotkey('ctrl', 'w')
    time.sleep(random.uniform(1.0, 2.0))
    

print("Mensajes enviados correctamente 🎉")
