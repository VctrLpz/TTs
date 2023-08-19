# -*- coding: utf-8 -*-
"""
Spyder Editor
@author: Victor

"""
import pandas as pd
import pyttsx3
import os 

# Lee el archivo de Excel
df = pd.read_excel('Source/BLASTER TDC 030823.xlsx')
# Selecciona la columna que deseas
columna = df['B Texto']
#inicializamos el TTs
engine = pyttsx3.init()
# Cambiar la velocidad de la  voz
rate = engine.getProperty('rate')
#engine.setProperty('rate', rate-5)
# Cambiar el volumen de la voz
volume = engine.getProperty('volume')
#engine.setProperty('volume', volume-0.25)
#creamos la carpeta para la salida de los audios 
if not os.path.exists('Salida'):
        os.makedirs('Salida')
# Obtener linea por linea del archivo 
for i, line in enumerate(columna):
    #Guardar las pistas de audio 
    engine.save_to_file(line, f'Salida/test_{i}.mp3')
    engine.runAndWait()
    print("Procesando..")
print("Terminado")