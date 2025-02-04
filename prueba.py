import pandas as pd
import webbrowser as web
import pyautogui as pg
import time
import random
import subprocess

# URL con el formato correcto para descargar CSV desde Google Sheets
csv_url = "https://docs.google.com/spreadsheets/d/1bgbrblcVg2qz8WfOSH0UIwvVg7o-9-e9wKNnrU-KydQ/export?format=csv"
data = pd.read_csv(csv_url)

print(data.head())

# Ruta de Google Chrome (ajusta según tu sistema)
chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe'

for i in range(len(data)):
    concesionario = data.loc[i, 'Concesionarios']
    
    # Crear mensaje personalizado
    formula = f'=QUERY(IMPORTRANGE("1Y_FwGsWkmN_SJj7aRiCl2Gld4_XvlO6HMP7VGSgOd-A";"Respuestas de formulario 1!A:AA");"SELECT * WHERE Col4 = \'{concesionario}\' ";1)'
    
    # Abrir WhatsApp Web en Google Chrome
    url = f"https://docs.google.com/spreadsheets/u/0/?tgif=d"
    subprocess.Popen([chrome_path, url])
    
    # Esperar un tiempo aleatorio para cargar la página
    time.sleep(random.randint(10, 15))
    
    # Simular la interacción para enviar el mensaje
    pg.click(307, 316)  
    time.sleep(random.uniform(1.5, 3.0))
    pg.click(171, 142)
    time.sleep(random.uniform(0.5, 1.5))
    pg.press('delete')
    pg.write(concesionario, interval=0.05)
    time.sleep(random.uniform(0.5, 1.5))
    pg.press('enter')
    time.sleep(random.uniform(0.5, 1.5))
    pg.click(103, 298)  
    time.sleep(random.uniform(2.0, 4.0)) 
    pg.write(formula, interval=0.05)
    time.sleep(random.uniform(0.5, 1.5))
    pg.press('enter')
    time.sleep(random.uniform(2.0, 4.0))
    pg.click(102, 298)
    time.sleep(random.uniform(1.5, 3.0))
    pg.click(349, 446)
    time.sleep(random.uniform(10.0, 13.0))
    
    # Cerrar la pestaña
    pg.hotkey('ctrl', 'w')
    time.sleep(random.uniform(1.0, 2.0))
    
    
    # Intentar localizar el logo de WhatsApp en la pantalla (No es necesario esta parte del codigo, es solamente para ver si cargo correctamente el wapp pero internamente)
    # try:
    #     if pg.locateOnScreen('whatsapp_logo.png', confidence=0.8):  # Detectar si WhatsApp cargó
    #         print(f"WhatsApp cargó correctamente para {nombre}")
    #     else:
    #         print(f"Error: No se pudo encontrar el logo de WhatsApp para {nombre}")
    # except pg.ImageNotFoundException:
    #     print(f"Error: No se pudo localizar la imagen de WhatsApp para {nombre}")
    
print("Archivos creados correctamente 🎉")
