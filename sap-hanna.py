import subprocess
import time
import win32com.client
import tkinter as tk
from tkinter import messagebox, Button
from datetime import datetime
import pyautogui
import cv2
import numpy as np

now = datetime.now()

# Formatear la fecha como DD-MM-YYYY
dt_string = now.strftime("%d-%m-%Y")

# Crear el nombre del archivo con el formato deseado
filename = f"REPORTE FLOTA {dt_string}.XLSX"

# Ruta del directorio donde se guardará el archivo
folderdir = "###########"

print(folderdir)
# Ruta completa del archivo
folderdira = folderdir + filename


class SapGui:
   
    def __init__(self):
        self.path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
        self.connection = None
        self.session = None

        # Inicia el proceso de SAP Logon
        subprocess.Popen(self.path)
        time.sleep(2)  # Espera a que SAP Logon se inicie

        # Conectar a la interfaz de SAP GUI
        self.connect_to_sap()

    def connect_to_sap(self):
        """Conectar a SAP GUI y abrir una sesión."""
        try:
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not isinstance(self.SapGuiAuto, win32com.client.CDispatch):
                raise Exception("No se pudo conectar a SAP GUI")

            application = self.SapGuiAuto.GetScriptingEngine
            if not application:
                raise Exception("GetScriptingEngine no está disponible. Verifica si el scripting está habilitado.")

            # Abre la conexión especificada
            self.connection = application.OpenConnection("########", True)
            time.sleep(2)  # Espera para que la conexión se complete

            # Inicia una sesión y maximiza la ventana
            self.session = self.connection.Children(0)
            self.session.findById("wnd[0]").maximize()
        except Exception as e:
            print(f"Error en la conexión a SAP: {e}")
            messagebox.showerror("Error de conexión", f"Ocurrió un error: {e}")

    def sapLogin(self):
        """Iniciar sesión en SAP."""
        try:
            # Accede a los campos de login en SAP y rellena la información de inicio de sesión
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "####"      # Cliente
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "####"    # Usuario
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "####"  # Contraseña
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "ES"        # Idioma
            self.session.findById("wnd[0]").sendVKey(0)                          # Envía la tecla Enter para iniciar sesión
            time.sleep(2)
            self.execute_transaction()

            messagebox.showinfo("Información", "Login realizado con éxito")
        except Exception as e:
            print(f"Error durante el inicio de sesión: {e}")
            messagebox.showerror("Error de inicio de sesión", f"Ocurrió un error: {e}")

    def execute_transaction(self):
        """Ejecutar la transacción ZPMRI0001 y realizar acciones necesarias."""
        try:
            session = self.session
            
            # Ejecutar la transacción
            session.findById("wnd[0]/tbar[0]/okcd").text = "######"
            session.findById("wnd[0]").sendVKey(0)

            # Enviar la tecla F4
            session.findById("wnd[0]").sendVKey(4)

            # Establecer el enfoque en una etiqueta y definir la posición del cursor
            session.findById("wnd[1]/usr/lbl[1,7]").setFocus()
            session.findById("wnd[1]/usr/lbl[1,7]").caretPosition = 3

            # Presionar el botón en la ventana modal
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # Seleccionar una casilla de verificación
            session.findById("wnd[0]/usr/chkP_CHECK").selected = True
            session.findById("wnd[0]/usr/chkP_CHECK").setFocus()

            # Presionar el botón "Ejecutar" (F8)
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            # Navegar por el menú
            session.findById("wnd[0]/mbar/menu[4]/menu[0]/menu[1]").select()

            # Seleccionar fila en la tabla
            table = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
            table.currentCellRow = 6
            table.selectedRows = "6"
            table.clickCurrentCell()

            # Interacción con la tabla en la ventana principal
            main_table = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            main_table.setCurrentCell(4, "SERGE")
            main_table.selectedRows = "4"
            main_table.contextMenu()

            # Seleccionar una opción del menú contextual
            main_table.selectContextMenuItem("&XXL")

            # Cerrar la ventana adicional
            session.findById("wnd[1]/tbar[0]/btn[20]").press()

            session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderdira
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            print("FINNNNNNNNNNNNN")
            def find_button(button_image_path):
                # Tomamos una captura de pantalla de la pantalla completa
                screenshot = pyautogui.screenshot()

                # Convertimos la captura de pantalla a un array de numpy
                screenshot_np = np.array(screenshot)

                # Convertimos la imagen de la captura de pantalla a formato RGB
                screenshot_rgb = cv2.cvtColor(screenshot_np, cv2.COLOR_BGR2RGB)

                # Cargamos la imagen del botón a buscar
                button_img = cv2.imread(button_image_path)

                # Usamos la función de template matching de OpenCV para encontrar la imagen en la pantalla
                result = cv2.matchTemplate(screenshot_rgb, button_img, cv2.TM_CCOEFF_NORMED)

                # Establecemos un umbral de similitud para el resultado (ajústalo según necesidad)
                threshold = 0.8
                locations = np.where(result >= threshold)

                # Si encontramos una coincidencia, devolvemos la posición
                if len(locations[0]) > 0:
                    # Obtener la primera coincidencia
                    loc = (locations[1][0], locations[0][0])
                    return loc
                else:
                    return None

            def click_button(button_image_path):
                # Buscar el botón en la pantalla
                button_location = find_button(button_image_path)

                if button_location:
                    pyautogui.click(button_location)
                    print(f"Botón encontrado y clickeado en {button_location}")
                else:
                    print("Botón no encontrado")
            button_image_path = "boton-save.png" 
            click_button(button_image_path)

        except Exception as e:
            print(f"Error durante la ejecución de la transacción: {e}")
            messagebox.showerror("Error en ejecución", f"Ocurrió un error: {e}")

if __name__ == "__main__":
    window = tk.Tk()
    window.geometry("300x200")

    # Botón de inicio de sesión
    login_btn = Button(window, text="Login SAP", command=lambda: SapGui().sapLogin())
    login_btn.pack(pady=10)

    window.mainloop()