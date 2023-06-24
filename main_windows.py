import time
import subprocess
import sys
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket

def ejecutar_script(script):
    # Ejecuta el archivo .exe utilizando subprocess
    subprocess.run([script])

class MyService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'MyService'
    _svc_display_name_ = 'Mi Servicio'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_running = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        # Ejecuta tus scripts
        script1 = r"C:\ruta\al\archivo.exe"
        script2 = r"C:\ruta\al\archivo2.exe"

        while self.is_running:
            # Ejecuta el archivo .exe 1
            print("Recibiendo datos de la bascula.exe...")
            ejecutar_script(script2)

            print("Ejecutando script.exe...")
            ejecutar_script(script1)

            # Espera 1 segundos
            print("Esperando 10 segundos...")
            time.sleep(1)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(MyService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(MyService)