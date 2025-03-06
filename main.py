import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from pandastable import Table, TableModel
import paramiko
from datetime import datetime
import os
import threading

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Respaldo Antenas Ubiquiti")
        self.root.geometry("900x900")
        
        # Botón para cargar archivo
        self.btn_cargar = tk.Button(root, text="Subir Archivo Excel", command=self.cargar_excel)
        self.btn_cargar.pack(pady=10)
        
        # Botón para limpiar datos
        self.btn_limpiar = tk.Button(root, text="Limpiar Datos", command=self.limpiar_datos)
        self.btn_limpiar.pack(pady=10)
        
        # Frame para la tabla
        self.frame_tabla = tk.Frame(root)
        self.frame_tabla.pack(fill="both", expand=True)
        
        # Botón para realizar respaldos
        self.btn_respaldo = tk.Button(root, text="Realizar Respaldo", command=self.iniciar_respaldo_thread)
        self.btn_respaldo.pack(pady=10)
        
        # Área de mensajes
        self.text_area = scrolledtext.ScrolledText(root, height=8, wrap=tk.WORD)
        self.text_area.pack(fill="both", expand=True, padx=10, pady=5)
        
        #Inicializar DataFrame con encabezados y una fila vacía
        columnas = ["IP", "Antena", "Usuario", "Contraseña", "Puerto"]
        self.df = pd.DataFrame(columns=columnas)
        # Añadir una fila vacía para mostrar los encabezados
        self.df = pd.concat([self.df, pd.DataFrame([{col: '' for col in columnas}])], ignore_index=True)
        self.pt = None
        self.mostrar_tabla()
    
    def cargar_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[["Excel Files", "*.xlsx;*.xls"]])
        if not file_path:
            return
        
        try:
            self.df = pd.read_excel(file_path, usecols="A:E")
            self.df.columns = ["IP", "Antena", "Usuario", "Contraseña", "Puerto"]
            self.actualizar_tabla()
            self.mostrar_mensaje("Archivo cargado exitosamente.")
        except Exception as e:
            self.mostrar_mensaje(f"Error al cargar el archivo: {str(e)}")
    
    def mostrar_tabla(self):
        if self.pt:
            self.pt.destroy()
        model = TableModel(self.df)
        self.pt = Table(self.frame_tabla, 
                      model=model,
                      showtoolbar=False,
                      showstatusbar=False,
                      editable=True,
                      showindex=False)
        self.pt.show()
        self.pt.autoResizeColumns()
    
    def actualizar_tabla(self):
        if self.pt:
            self.pt.model.df = self.df
            self.pt.showindex = False  # Ocultar índice
            self.pt.redraw()
            self.pt.autoResizeColumns()
        self.frame_tabla.update_idletasks()
    
    def obtener_datos_tabla(self):
        if self.pt:
            self.df = self.pt.model.df.copy()
    
    def iniciar_respaldo_thread(self):
        thread = threading.Thread(target=self.realizar_respaldo)
        thread.start()
    
    def realizar_respaldo(self):
        self.obtener_datos_tabla()
        
        if self.df.empty:
            self.mostrar_mensaje("Advertencia: No hay datos para respaldar.")
            return
        
        directorio = filedialog.askdirectory(title="Seleccionar carpeta para respaldos")
        if not directorio:
            self.mostrar_mensaje("No se seleccionó un directorio para los respaldos.")
            return
        
        self.mostrar_mensaje("Iniciando proceso de respaldo...")
        
        for _, row in self.df.iterrows():
            ip = str(row["IP"])
            usuario = str(row["Usuario"])
            password = str(row["Contraseña"])
            puerto = str(row["Puerto"])
            
            self.mostrar_mensaje(f"Conectando con IP: {ip}")
            self.respaldar_antena(ip, usuario, password, directorio, puerto)
        
        self.mostrar_mensaje("Proceso de respaldo finalizado.")
    
    def respaldar_antena(self, ip, usuario, contrasenia, directorio, puerto):
        cliente_ssh = None
        try:
            cliente_ssh = paramiko.SSHClient()
            cliente_ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            cliente_ssh.connect(ip, username=usuario, password=contrasenia, port=puerto, timeout=5)
            
            comando_respaldo = "cat /tmp/system.cfg"
            stdin, stdout, stderr = cliente_ssh.exec_command(comando_respaldo)
            configuracion = stdout.read().decode()
            
            fecha_actual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            nombre_archivo = f"respaldo_{ip}_{fecha_actual}.cfg"
            ruta_archivo = os.path.join(directorio, nombre_archivo)
            
            with open(ruta_archivo, "w") as archivo:
                archivo.write(configuracion)
            
            self.mostrar_mensaje(f"✅ Respaldo exitoso para la IP: {ip}, archivo: {nombre_archivo}")
            
        except paramiko.AuthenticationException:
            self.mostrar_mensaje(f"❌ Error de autenticación para la antena con IP {ip}. Verifica las credenciales.")
        except paramiko.SSHException as e:
            self.mostrar_mensaje(f"⚠️ Error SSH para la antena con IP {ip}: {e}")
        except paramiko.ssh_exception.NoValidConnectionsError:
            self.mostrar_mensaje(f"❌ No se pudo establecer conexión SSH con la antena con IP {ip}")
        except Exception as e:
            self.mostrar_mensaje(f"❌ Error desconocido para la antena con IP {ip}: {e}")
        finally:
            if cliente_ssh:
                cliente_ssh.close()
    
    def limpiar_datos(self):
        self.df = pd.DataFrame(columns=["IP", "Antena", "Usuario", "Contraseña", "Puerto"])
        self.actualizar_tabla()
        self.text_area.delete(1.0, tk.END)
        self.mostrar_mensaje("Datos y mensajes limpiados.")
    
    def mostrar_mensaje(self, mensaje):
        self.text_area.insert(tk.END, mensaje + "\n")
        self.text_area.see(tk.END)
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
