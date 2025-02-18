import pandas as pd
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
    url = str(data.loc[i, 'url'])  # Convertir a string
    concesionarioFure = data.loc[i, 'ConcesionarioFure']
    concesionarioVDG = data.loc[i, 'ConcesionarioVDG']
    
    # Formula para importar los datos de un concesionario específico
    formulaVdg = f'=QUERY(IMPORTRANGE("1ti5Dl8UGHXlNrtv0HsUStgTDVY41R1y-Tp9RiXBzPEU";"Respuestas de formulario 1!A:N");"SELECT * WHERE Col2 CONTAINS \'{concesionarioVDG} ";1)'
    formulaFure = f'=QUERY(IMPORTRANGE("11dAWSHafMhbI1mbZQjVyD_0idGlabzKp5K-WzKHvQC4";"Canal Rojo!A:V");"SELECT * WHERE Col6 CONTAINS \'{concesionarioFure} ";2)'
    NombreFure = f'FURE'
    NombreVdg = f'VDG'
    Aceptado = f'Aceptado'
    Rechazado = f'Rechazado'
    FormulaCondicionalFure  = f'Q1:Q1000'
    FormulaCondicionalVdg  = f'H1:H998'
    
# Abrir Google Sheets en Google Chrome
    subprocess.Popen([chrome_path, url])
    
# Esperar un tiempo aleatorio para cargar la página
    time.sleep(random.randint(5, 8))
    
# Posicionamiento en Hoja 2
    pg.click(502, 700)
    time.sleep(random.uniform(1.5, 2.5))
# Simular la interacción para la creacion de una nueva hoja (VDG)
    pg.click(62, 702)  
    time.sleep(random.uniform(1.5, 3.0))
# Escribir formula VDG
    pg.write(formulaVdg, interval=0.05)
    time.sleep(random.uniform(0.5, 1.5))
    pg.press('enter')
    time.sleep(random.uniform(2.0, 4.0))
# Crear nueva hoja (Fure)
    pg.click(62, 702)
    time.sleep(random.uniform(1.5, 3.0))
# Escribir formula Fure
    pg.write(formulaFure, interval=0.05)
    time.sleep(random.uniform(0.5, 1.5))
    pg.press('enter')
    time.sleep(random.uniform(2.0, 4.0))
# Dar permiso de acceso en la hoja de Fure
    pg.click(100, 298)
    time.sleep(random.uniform(1.5, 2.0))
    pg.click(348, 447)
    time.sleep(random.uniform(1.5, 2.0))
# Cambiar Nombre de las hojas
    pg.doubleClick(613, 703)
    time.sleep(random.uniform(1.0, 2.0))
    pg.write(NombreVdg, interval=0.05)
    pg.doubleClick(687, 704)
    time.sleep(random.uniform(1.0, 2.0))
    pg.write(NombreFure, interval=0.05)
    time.sleep(random.uniform(1.0, 2.0))
    pg.press('enter')
    time.sleep(random.uniform(1.0, 2.0))
# Agrandar la fila 1 del Fure
    pg.moveTo(15, 306)  
    pg.mouseDown()  
    pg.moveTo(27, 341, duration=0.5)  
    pg.mouseUp()  

# Pequeña pausa antes de la segunda acción
    time.sleep(1)
    
# Achicar la pantalla
    pg.click(230, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(211, 242)
    time.sleep(random.uniform(1.0, 1.5))
    

# Seleccionar el rango de celdas Del Fure
    pg.moveTo(47, 279)  
    pg.mouseDown()  
    pg.moveTo(360, 280, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))
    
# Pintar celdas de color Gris oscuro
    pg.click(799, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(850, 278)
    time.sleep(random.uniform(1.0, 1.5))
# Pintamos las letras de blanco
    pg.click(760, 211)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(963, 279)
    time.sleep(random.uniform(1.0, 1.5))
    
# Seleccionar el otro rango de celdas Del Fure
    pg.moveTo(411, 279)  
    pg.mouseDown()  
    pg.moveTo(1127, 280, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))

# Pintar celdas de color rojo mate
    pg.click(799, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(872, 478)
    time.sleep(random.uniform(1.0, 1.5))
# Pintamos las letras de blanco
    pg.click(760, 211)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(963, 279)
    time.sleep(random.uniform(1.0, 1.5))

# Centramos las celdas y encuadramos
    pg.click(15, 269)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(921, 207)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(954, 245)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(960, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(997, 245)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1005, 206)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1036, 246)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(833, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(845, 251)
    time.sleep(random.uniform(1.0, 1.5))

    

# Agregamos Formato condicional
    pg.click(314, 167)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(436, 540)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1144, 274)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1129, 361)
    pg.hotkey("ctrl", "a")  # Seleccionar todo (Windows/Linux)
    pg.press("delete")  # Borrar
    time.sleep(random.uniform(1.0, 1.5))
    pg.write(FormulaCondicionalFure, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1160, 480)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1169, 639)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1126, 527)
    pg.write(Aceptado, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1228, 644)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1116, 362)
    time.sleep(random.uniform(1.0, 1.5))
    pg.scroll(-300)
    time.sleep(random.uniform(0.5, 1.5))
    pg.click(1272, 586)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1130, 350)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1129, 361)
    pg.hotkey("ctrl", "a")  # Seleccionar todo (Windows/Linux)
    pg.press("delete")  # Borrar
    time.sleep(random.uniform(1.0, 1.5))
    pg.write(FormulaCondicionalFure, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1160, 480)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1169, 639)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1126, 527)
    pg.write(Rechazado, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1228, 644)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1051, 360)
    time.sleep(random.uniform(1.0, 1.5))
    pg.scroll(-300)
    time.sleep(random.uniform(0.5, 1.5))
    pg.click(1266, 644)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1322, 214)
    time.sleep(random.uniform(1.0, 1.5))
    

# #Estiramos las celdas del fure
    pg.moveTo(51, 269)  
    pg.mouseDown()  
    pg.moveTo(1126, 268, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))
    pg.moveTo(73, 270)  
    pg.mouseDown()  
    pg.moveTo(99, 268, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))

# Ocultamos columnas
    pg.click(1020, 546)
    pg.keyDown("ctrl")
    pg.press("right")
    time.sleep(random.uniform(0.5, 1.5))
    pg.keyUp("ctrl")
    time.sleep(random.uniform(0.5, 1.5))
    pg.moveTo(1173, 267)  
    pg.mouseDown()  
    pg.moveTo(1338, 268, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))
    pg.rightClick()
    pg.click(1136, 497)
    time.sleep(random.uniform(0.5, 1.5))
    pg.click(1020, 546)
    pg.keyDown("ctrl")
    pg.press("left")
    time.sleep(random.uniform(0.5, 1.5))
    pg.keyUp("ctrl")
    time.sleep(random.uniform(1.0, 2.0))
    
####################################################################################################################
    
# Pasamos a la hoja de VDG
    pg.click(608, 704)
    time.sleep(random.uniform(1.0, 2.0))
    
# Agrandar la fila 1 del Fure
    pg.moveTo(12, 284)  
    pg.mouseDown()  
    pg.moveTo(12, 304, duration=0.5)  
    pg.mouseUp()  

# Pequeña pausa antes de la segunda acción
    time.sleep(1)
    
# Seleccionar el rango de celdas Del VDG
    pg.moveTo(47, 279)  
    pg.mouseDown()  
    pg.moveTo(360, 280, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))

# Pintar celdas de color rojo mate
    pg.click(799, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(872, 478)
    time.sleep(random.uniform(1.0, 1.5))
# Pintamos las letras de blanco
    pg.click(760, 211)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(963, 279)
    time.sleep(random.uniform(1.0, 1.5))

# Seleccionar el otro rango de celdas Del Fure
    pg.moveTo(411, 279)  
    pg.mouseDown()  
    pg.moveTo(713, 281, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))
    
# Pintar celdas de color rojo mate
    pg.click(799, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(954, 342)
    time.sleep(random.uniform(1.0, 1.5))
# Pintamos las letras de negro fuerte
    pg.click(661, 207)
    time.sleep(random.uniform(1.0, 1.5))
    
# Centramos las celdas y encuadramos
    pg.click(15, 269)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(921, 207)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(954, 245)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(960, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(997, 245)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1005, 206)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1036, 246)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(833, 208)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(845, 251)
    time.sleep(random.uniform(1.0, 1.5))
    

# Agregamos Formato condicional
    pg.click(314, 167)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(436, 540)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1144, 274)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1129, 361)
    pg.hotkey("ctrl", "a")  # Seleccionar todo (Windows/Linux)
    pg.press("delete")  # Borrar
    time.sleep(random.uniform(1.0, 1.5))
    pg.write(FormulaCondicionalVdg, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1160, 480)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1169, 639)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1126, 527)
    pg.write(Aceptado, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1228, 644)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1116, 362)
    time.sleep(random.uniform(1.0, 1.5))
    pg.scroll(-300)
    time.sleep(random.uniform(0.5, 1.5))
    pg.click(1272, 586)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1130, 350)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1129, 361)
    pg.hotkey("ctrl", "a")  # Seleccionar todo (Windows/Linux)
    pg.press("delete")  # Borrar
    time.sleep(random.uniform(1.0, 1.5))
    pg.write(FormulaCondicionalVdg, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1160, 480)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1169, 639)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1126, 527)
    pg.write(Rechazado, interval=0.05)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1228, 644)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1051, 360)
    time.sleep(random.uniform(1.0, 1.5))
    pg.scroll(-300)
    time.sleep(random.uniform(0.5, 1.5))
    pg.click(1266, 584)
    time.sleep(random.uniform(1.0, 1.5))
    pg.click(1322, 214)
    time.sleep(random.uniform(1.0, 1.5))


# #Estiramos las celdas del fure
    pg.moveTo(51, 269)  
    pg.mouseDown()  
    pg.moveTo(714, 268, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))
    pg.moveTo(73, 270)  
    pg.mouseDown()  
    pg.moveTo(99, 268, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))
    
# Ocultamos columnas
    pg.click(1020, 546)
    pg.keyDown("ctrl")
    pg.press("right")
    time.sleep(random.uniform(0.5, 1.5))
    pg.keyUp("ctrl")
    time.sleep(random.uniform(0.5, 1.5))
    pg.moveTo(769, 267)  
    pg.mouseDown()  
    pg.moveTo(1332, 268, duration=0.5)  
    pg.mouseUp()
    time.sleep(random.uniform(0.5, 1.5))
    pg.rightClick()
    pg.click(1136, 497)
    time.sleep(random.uniform(0.5, 1.5))
    pg.click(1020, 546)
    pg.keyDown("ctrl")
    pg.press("left")
    time.sleep(random.uniform(0.5, 1.5))
    pg.keyUp("ctrl")
    time.sleep(random.uniform(1.0, 2.0))

# Cerrar la pestaña
    time.sleep(random.uniform(5.0, 6.5))
    pg.hotkey('ctrl', 'w')
    time.sleep(random.uniform(1.0, 2.0))
    
# print("Archivos creados correctamente 🎉")
