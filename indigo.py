import psutil
import pyperclip
from pywinauto import Application
import logging
import os
import time
from pywinauto import Desktop
import pyautogui

class Indigo:
    
    def __init__(self):
        self.ruta = os.getenv('INDIGO_PATH')
        self.descargas = r"temp"
        
        
    def obtener_app(self):
        try:
            pid = self.verificar_proceso("Vie Cloud Platform.exe")
            if pid:
                print(f"✅ La aplicación está abierta.")
                return Application(backend="uia").connect(process=pid)
            else:
                print(f"😒 La aplicación no está abierta.")
                
                pid = self.abrir_app()
                return Application(backend="uia").connect(process=pid)
        except Exception as e:
            print(f"❌ Error al verificar la aplicación: {e}")
            return False
   

    def verificar_proceso(self, nombre_proceso):
        
        for proceso in psutil.process_iter(['pid', 'name']):
            try:
                if nombre_proceso.lower() in proceso.info['name'].lower():
                    return proceso.info['pid'] 
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        return None
        
    def abrir_app(self):
        try:            
            proceso = Application(backend="uia").start(self.ruta)
            return proceso.process
        except Exception as e:         
            print("❌  Error al abrir la aplicación:", str(e))
            logging.error(f'❌ Error al abrir la aplicación: {str(e)}')

    def validar_login(self, app, empresa):
        try:
            windows = app.windows()          
            # Buscar ventana de login
            login_window = next((w for w in windows if w.element_info.automation_id == "FrmLoginAzure"), None)
            if login_window:
                self.window = app.window(handle=login_window.handle)
                return True

            # Buscar ventana principal (ya logueado)
            main_window = next((w for w in windows if w.element_info.automation_id == "FormMdi"), None)
            if main_window:
                self.window = app.window(handle=main_window.handle)
                contenedor = self.window.child_window(auto_id="INDPceChangePerfil", control_type="Edit").wait('ready', timeout=30)
                btn_open = contenedor.children()[1]
                btn_open.click_input()
                self.select_workspace(app, empresa)
                return False

            print("❌ No se encontró ninguna ventana válida.")
            logging.error("❌ No se encontró ninguna ventana válida.")
            return False

        except Exception as e:
            print("❌ Error al validar login:", str(e))
            logging.error(f'❌ Error al validar login: {str(e)}')
            return False

    def login(self, app, user, password, empresa):
        
        if self.validar_login(app, empresa):
            try:
                btn_microsoft = self.window.child_window(title="Microsoft", auto_id="IndigoExchange", control_type="Button").wait('ready', timeout=30)
                btn_microsoft.click()
                
                correo = self.window.child_window(title="Escriba su correo electrónico, teléfono o Skype.", auto_id="i0116", control_type="Edit").wait('ready', timeout=30)
                correo.set_focus()
                correo.type_keys(user, with_spaces=True)

                btn_sgt = self.window.child_window(title="Siguiente", auto_id="idSIButton9", control_type="Button").wait('ready', timeout=30)
                btn_sgt.click()

                passw = self.window.child_window(title="Escriba la contraseña para juan.gallegor@zentria.com.co", auto_id="i0118", control_type="Edit").wait('ready', timeout=30)
                passw.set_focus()
                passw.type_keys(password, with_spaces=True)

                btn_signin = self.window.child_window(title="Iniciar sesión", auto_id="idSIButton9", control_type="Button").wait('ready', timeout=30)
                btn_signin.click()

                btn_si = self.window.child_window(title="Sí", auto_id="idSIButton9", control_type="Button").wait('ready', timeout=30)
                btn_si.click()
                
                print('✅ Inicio de sesión exitoso')
                
                self.select_workspace(app, empresa)
                return True
            
            except Exception as e:
                print("❌ Error al iniciar sesión:", str(e))
                logging.error(f'❌ Error al iniciar sesión: {str(e)}')
                return False
        print("✅ La aplicación ya está abierta y logueada.")
        return True
                
    def select_workspace(self, app, empresa):
        try:
            self.window.child_window(title="Cerrar", control_type="Button", auto_id="INDSmbNo").wait('ready', timeout=5).click()
        except Exception as e:
            print("✅ No se encontró la ventana para cerrar el error:", str(e))

        try:
            if empresa == "ODO":
                workspace = self.window.child_window(title="WorkspaceUser", auto_id="WorkspaceUser", control_type="Pane").wait('visible', timeout=30)
                workspace.set_focus()

                odo = self.window.child_window(title="Fila 2", control_type="Custom").wait('visible', timeout=30)

                odo.click_input()
                
                try:
                    self.window.child_window(title="11 PEREIRA", control_type="Text").wait('visible', timeout=30)
                except Exception as e:
                    self.window.child_window(title="13 MANIZALES", control_type="Text").wait('visible', timeout=30)

                acepta = self.window.child_window(title="Aceptar", auto_id="INDSmbAceptar", control_type="Button").wait('ready', timeout=30)
                acepta.click()
            
                self.window = app.window(auto_id="FormMdi")
                self.window.wait('visible', timeout=30)
                
                return True
            else:
                workspace = self.window.child_window(title="WorkspaceUser", auto_id="WorkspaceUser", control_type="Pane").wait('visible', timeout=30)
                workspace.set_focus()

                ccb = self.window.child_window(title="Fila 4", control_type="Custom").wait('visible', timeout=30)

                ccb.click_input()
                
                self.window.child_window(title="35 TUNJA", control_type="Text").wait('visible', timeout=30)

                acepta = self.window.child_window(title="Aceptar", auto_id="INDSmbAceptar", control_type="Button").wait('ready', timeout=30)
                acepta.click()
            
                self.window = app.window(auto_id="FormMdi")
                self.window.wait('visible', timeout=30)
                
                return True
        except Exception as e:
            print("Error al seleccionar el workspace:", str(e))
            return False


    def agregar_facturas(self, data):
            try:
                ventana_facturas = self.window.child_window(title="Facturas", control_type="Table").wait('ready', timeout=30)
                ventana_facturas.click_input()
                time.sleep(2)
                datos_factura = f"{data['factura']}	3	5	{data['valor_nota']}"
                pyperclip.copy(datos_factura)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(10)
                attempts = 3
                while attempts > 0:
                    try:
                        botones = [b for b in self.window.descendants(control_type="Button") if b.window_text() == "Página Derecha"]
                        botones[0].click_input()
                        break
                    except IndexError:
                        time.sleep(10)
                        attempts -= 1
                        if attempts == 0:
                            print("❌ No se pudo encontrar el botón 'Página Derecha' después de varios intentos.")
                            return False
                        continue
                
                self.agregar_conceptos(data)

            except Exception as e:
                print("❌ Error al agregar la factura:", e)
                return False

    def agregar_conceptos(self, data):
        try:
            conceptos_add = self.window.child_window(auto_id="INDpceConcept", title="Conceptos", control_type="Edit")
            abrir = [a for a in conceptos_add.children() if a.window_text() == "Abrir"]
            abrir[0].click_input()
            time.sleep(2)
            pyautogui.write('035', interval=0.1)
            time.sleep(5)
            pyautogui.press('down')
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.press('enter')
            pyautogui.write('110112100', interval=0.1)
            time.sleep(2)
            pyautogui.press('down')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(data['valor_nota'], interval=0.1)
            pyautogui.press('enter')
            pyautogui.press('enter')
            time.sleep(5)
            ventana_facturas = self.window.child_window(title="Facturas", control_type="Table").wait('ready', timeout=30)
            ventana_facturas.click_input()
        except Exception as e:
            print("Error al abrir conceptos:", e)
            return False

    def formulario_general(self, data):
        try:
            nuevo = self.window.child_window(title="Panel de transacción", control_type="ToolBar").wait('ready', timeout=30)
            nuevo.click_input()
            cliente = self.window.child_window(control_type="Edit", auto_id="INDsleClient").wait("ready", timeout=30)
            cliente.set_focus()
            time.sleep(2)
            cliente.type_keys(data['nit'], with_spaces=True, pause=0.3)
            time.sleep(4)
            pyautogui.press('down')
            pyautogui.press('enter')
            time.sleep(4)
            observacion = self.window.child_window(title="", control_type="Edit", auto_id="INDmemoComments").wait("ready", timeout=30)
            observacion.type_keys(data["observacion"], with_spaces=True)

            self.agregar_facturas(data)
            return True
        except Exception as e:
            print("Error al abrir el formulario general:", e)
            return False

"""
Name:	"Proveedor row0"
ControlType:	UIA_DataItemControlTypeId (0xC36D)
LocalizedControlType:	"elemento"

"""

if __name__ == "__main__":
    data = {
        "id_ticket": "123456",
        "nit": "800088702",
        "entidad": "Entidad de Prueba",
        "factura": "QC2863",
        "dpp": "DPP de Prueba",
        "dto_modelo_aplicado": "DTO Modelo Aplicado de Prueba",
        "valor_nota": "49831",
        "observacion": "NOTA CREDITO POR DESCUENTO DE CARTERA",
        "centro_costos": "Centro de Costos de Prueba",
        "numero_nota": "Número Nota de Prueba",
        "duracion_proceso": "Duración Proceso de Prueba",
        "fecha_proceso": "Fecha Proceso de Prueba"
    }
    indigo = Indigo()
    app = indigo.obtener_app()
    if app:
        ventana = app.top_window() # type: ignore
        ventana.set_focus()
        if indigo.login(app, 'juan.gallegor@zentria.com.co', 'Gzentria5657*', "ODO"):
             print("✅ Login exitoso")
        indigo.formulario_general(data)
    