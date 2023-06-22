import serial
import csv
import serial.tools.list_ports

# Muestra los puertos seriales disponibles
puertos_disponibles = serial.tools.list_ports.comports()
for puerto in puertos_disponibles:
    print(puerto.device)

print('-------------------------------------------')
puerto = 'COM1' # Modificar el puerto serial a leer  
baudios = 9600  

archivo_csv = 'datos.csv'
delimiter = ','

ser = serial.Serial(puerto, baudios)

with open(archivo_csv, 'w', newline='') as archivo:
    writer = csv.writer(archivo, delimiter=delimiter)
    peso_actual = None

    while True:
        if ser.in_waiting > 0:
            datos = ser.readline().decode('utf-8').rstrip()
            peso = datos  # Supongamos que el peso está en la variable "datos"

            if peso != peso_actual:
                peso_actual = peso
                writer.writerow([peso])

ser.close()