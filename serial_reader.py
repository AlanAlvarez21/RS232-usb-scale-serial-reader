# -*- coding: utf-8 -*-

import serial
import csv
import serial.tools.list_ports

# Muestra los puertos seriales disponibles
puertos_disponibles = serial.tools.list_ports.comports()
for puerto in puertos_disponibles:
    print(puerto.device)

print('-------------------------------------------')
puerto = '/dev/cu.usbmodem11401' # Reemplaza '/dev/cu.usbmodem11401' con el nombre de tu puerto serial

baudios = 9600  # Ajusta la velocidad de baudios según tu configuración
archivo_csv = './peso.csv'  # Nombre del archivo CSV a crear/sobrescribir

try:
    ser = serial.Serial(puerto, baudios)
    print('Conexión establecida en el puerto', puerto)

    while True:
        if ser.in_waiting > 0:
            datos = ser.readline().decode().strip()
            print('Datos recibidos:', datos)

            with open(archivo_csv, 'w', newline='') as archivo:
                writer = csv.writer(archivo)
                writer.writerow([datos])

except serial.SerialException as e:
    print('Error de conexión:', str(e))