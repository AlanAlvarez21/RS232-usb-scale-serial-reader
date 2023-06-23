import daemon
import time
import subprocess
import sys

def ejecutar_script(script):
    # Ejecuta el script utilizando subprocess
    subprocess.run(["python", script])

def main():
    # Ejecuta tus scripts
    script1 = "/Users/brito/excel_creator/serial_reader.py"
    script2 = "/Users/brito/excel_creator/script.py"
    
    while True:
        # Ejecuta el script 1
        print("Ejecutando script2...")
        ejecutar_script(script2)
        
        print("Ejecutando script1...")
        ejecutar_script(script1)
        
        # Espera 10 segundos
        print("Esperando 10 segundos...")
        time.sleep(10)

# Configura el contexto del demonio
with daemon.DaemonContext(stdout=sys.stdout, stderr=sys.stderr):
    try:
        main()
    except KeyboardInterrupt:
        print("Proceso detenido manualmente.")