import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import re
import os
import shutil
import sys
import json
import zipfile

class TurboScribeCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TurboScribe Limpiador Automático")
        self.root.geometry("600x600")
        self.root.resizable(False, False)
        
        self.config_file = self.obtener_ruta_config()
        self.config = self.cargar_configuracion()
        
        self.carpeta_actual = self.config.get("carpeta_origen", "")
        self.archivos_en_carpeta = []
        
        bg_color = "#f4f4f9"
        self.root.configure(bg=bg_color)
        
        # Estilo para Notebook
        style = ttk.Style()
        style.configure("TNotebook", background=bg_color)
        style.configure("TFrame", background=bg_color)
        
        # Titulo global
        lbl_titulo = tk.Label(self.root, text="Herramienta de Limpieza de Transcripciones", font=("Segoe UI", 14, "bold"), bg=bg_color, fg="#333333")
        lbl_titulo.pack(pady=(15, 10))
        
        # Notebook para Pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        # Frames para cada pestaña
        self.tab_limpieza = ttk.Frame(self.notebook)
        self.tab_configuracion = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_limpieza, text="✨ Limpieza")
        self.notebook.add(self.tab_configuracion, text="⚙️ Configuración")
        
        self.construir_tab_limpieza()
        self.construir_tab_configuracion()
        
        # Si hay carpeta configurada al abrir, cargarla
        if self.carpeta_actual and os.path.exists(self.carpeta_actual):
            self.lbl_ruta.config(text=f"Carpeta: {self.carpeta_actual}")
            self.actualizar_lista_archivos()

    def obtener_ruta_config(self):
        if getattr(sys, 'frozen', False):
            ruta_script = os.path.dirname(sys.executable)
        else:
            ruta_script = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(ruta_script, "config.json")
        
    def cargar_configuracion(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def guardar_configuracion(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            messagebox.showinfo("Configuración", "Configuración guardada correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la configuración:\n{e}")

    def construir_tab_limpieza(self):
        bg_color = "#f4f4f9"
        
        lbl_info = tk.Label(self.tab_limpieza, text="Selecciona una carpeta y luego los archivos a limpiar.", font=("Segoe UI", 10), bg=bg_color, fg="#666666")
        lbl_info.pack(pady=(15, 15))
        
        # Frame para botones superiores
        frame_botones_sup = tk.Frame(self.tab_limpieza, bg=bg_color)
        frame_botones_sup.pack(pady=(0, 5))

        btn_carpeta = tk.Button(frame_botones_sup, text="📁 Seleccionar Carpeta", command=self.seleccionar_carpeta, font=("Segoe UI", 10), bg="#2196F3", fg="white", relief=tk.FLAT, padx=15, pady=5, cursor="hand2")
        btn_carpeta.pack(side=tk.LEFT, padx=(0, 10))

        btn_refrescar = tk.Button(frame_botones_sup, text="🔄 Refrescar", command=self.actualizar_lista_archivos, font=("Segoe UI", 10), bg="#FF9800", fg="white", relief=tk.FLAT, padx=15, pady=5, cursor="hand2")
        btn_refrescar.pack(side=tk.LEFT)
        
        self.lbl_ruta = tk.Label(self.tab_limpieza, text="Carpeta: Ninguna seleccionada", font=("Segoe UI", 9), bg=bg_color, fg="#666666")
        self.lbl_ruta.pack(pady=(5, 5))
        
        # Frame para Listbox + Scrollbar
        frame_lista = tk.Frame(self.tab_limpieza, bg=bg_color)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        self.scrollbar = tk.Scrollbar(frame_lista)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(frame_lista, selectmode=tk.MULTIPLE, yscrollcommand=self.scrollbar.set, font=("Segoe UI", 10), selectbackground="#4CAF50")
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.listbox.yview)
        
        btn_limpiar = tk.Button(self.tab_limpieza, text="✨ Iniciar Limpieza", command=self.procesar_seleccion, font=("Segoe UI", 11, "bold"), bg="#4CAF50", fg="white", activebackground="#45a049", activeforeground="white", relief=tk.FLAT, padx=15, pady=8, cursor="hand2")
        btn_limpiar.pack(pady=(0, 20))

    def construir_tab_configuracion(self):
        bg_color = "#f4f4f9"
        
        lbl_info = tk.Label(self.tab_configuracion, text="Parámetros por defecto de la aplicación", font=("Segoe UI", 11, "bold"), bg=bg_color, fg="#333333")
        lbl_info.pack(pady=(20, 20))
        
        # --- Carpeta Origen ---
        frame_origen = tk.Frame(self.tab_configuracion, bg=bg_color)
        frame_origen.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(frame_origen, text="Carpeta de Origen por Defecto:", bg=bg_color, font=("Segoe UI", 10)).pack(anchor="w")
        
        self.entry_carpeta_origen = tk.Entry(frame_origen, font=("Segoe UI", 10))
        self.entry_carpeta_origen.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.entry_carpeta_origen.insert(0, self.config.get("carpeta_origen", ""))
        
        btn_browse_origen = tk.Button(frame_origen, text="Examinar", command=self.buscar_carpeta_origen_config, bg="#e0e0e0", relief=tk.FLAT, cursor="hand2")
        btn_browse_origen.pack(side=tk.RIGHT)
        
        # --- Archivo Diccionario ---
        frame_dicc = tk.Frame(self.tab_configuracion, bg=bg_color)
        frame_dicc.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(frame_dicc, text="Archivo de Diccionario (CSV):", bg=bg_color, font=("Segoe UI", 10)).pack(anchor="w")
        
        self.entry_diccionario = tk.Entry(frame_dicc, font=("Segoe UI", 10))
        self.entry_diccionario.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Cargar diccionario default de config si existe, sino, poner el default local
        dicc_guardado = self.config.get("archivo_diccionario", "")
        if not dicc_guardado:
            if getattr(sys, 'frozen', False):
                ruta_script = os.path.dirname(sys.executable)
            else:
                ruta_script = os.path.dirname(os.path.abspath(__file__))
            dicc_guardado = os.path.join(ruta_script, "Palabrascomunes.csv")
            
        self.entry_diccionario.insert(0, dicc_guardado)
        
        btn_browse_dicc = tk.Button(frame_dicc, text="Examinar", command=self.buscar_diccionario_config, bg="#e0e0e0", relief=tk.FLAT, cursor="hand2")
        btn_browse_dicc.pack(side=tk.RIGHT)
        
        # --- Boton Guardar ---
        btn_guardar = tk.Button(self.tab_configuracion, text="💾 Guardar Configuración", command=self.guardar_cambios_configuracion, font=("Segoe UI", 10, "bold"), bg="#2196F3", fg="white", relief=tk.FLAT, padx=15, pady=8, cursor="hand2")
        btn_guardar.pack(pady=40)

    def buscar_carpeta_origen_config(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de origen por defecto")
        if carpeta:
            self.entry_carpeta_origen.delete(0, tk.END)
            self.entry_carpeta_origen.insert(0, carpeta)
            
    def buscar_diccionario_config(self):
        archivo = filedialog.askopenfilename(title="Seleccionar Diccionario (CSV)", filetypes=[("Archivos CSV", "*.csv")])
        if archivo:
            self.entry_diccionario.delete(0, tk.END)
            self.entry_diccionario.insert(0, archivo)

    def guardar_cambios_configuracion(self):
        self.config["carpeta_origen"] = self.entry_carpeta_origen.get()
        self.config["archivo_diccionario"] = self.entry_diccionario.get()
        self.guardar_configuracion()
        
        # Sincronizar con estado actual
        if os.path.exists(self.config["carpeta_origen"]):
            self.carpeta_actual = self.config["carpeta_origen"]
            self.lbl_ruta.config(text=f"Carpeta: {self.carpeta_actual}")
            self.actualizar_lista_archivos()

    def cargar_diccionario(self, ruta_csv):
        diccionario = {}
        try:
            with open(ruta_csv, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                for row in reader:
                    if len(row) >= 2:
                        incorrecto = row[0].strip()
                        correcto = row[1].strip()
                        if incorrecto and correcto:
                            diccionario[incorrecto] = correcto
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el diccionario:\n{e}")
        return diccionario

    def limpiar_archivo(self, ruta_entrada, ruta_salida, diccionario):
        try:
            with open(ruta_entrada, 'r', encoding='utf-8') as f:
                contenido = f.read()

            claves_ordenadas = sorted(diccionario.keys(), key=len, reverse=True)
            
            for incorrecto in claves_ordenadas:
                correcto = diccionario[incorrecto]
                patron = re.compile(r'\b' + re.escape(incorrecto) + r'\b', re.IGNORECASE)
                contenido = patron.sub(correcto.replace('\\', '\\\\'), contenido)

            with open(ruta_salida, 'w', encoding='utf-8') as f:
                f.write(contenido)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar el archivo {os.path.basename(ruta_salida)}:\n{e}")
            return False

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de origen")
        if not carpeta:
            return
            
        self.carpeta_actual = carpeta
        self.lbl_ruta.config(text=f"Carpeta: {self.carpeta_actual}")
        
        self.actualizar_lista_archivos()
        
    def actualizar_lista_archivos(self):
        self.listbox.delete(0, tk.END)
        self.archivos_en_carpeta = []
        
        if not self.carpeta_actual or not os.path.exists(self.carpeta_actual):
            return
            
        extensiones_validas = {'.txt', '.srt', '.csv', '.vtt', '.zip'}
        
        for archivo in os.listdir(self.carpeta_actual):
            ruta_completa = os.path.join(self.carpeta_actual, archivo)
            if os.path.isfile(ruta_completa):
                _, ext = os.path.splitext(archivo)
                if ext.lower() in extensiones_validas:
                    self.archivos_en_carpeta.append(archivo)
                    
        for archivo in self.archivos_en_carpeta:
            self.listbox.insert(tk.END, archivo)
            
        if not self.archivos_en_carpeta:
            messagebox.showinfo("Información", "No se encontraron archivos compatibles (.txt, .srt, .csv, .vtt) en la carpeta seleccionada.")

    def procesar_seleccion(self):
        indices_seleccionados = self.listbox.curselection()
        if not indices_seleccionados:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo de la lista.")
            return
            
        ruta_csv_default = self.config.get("archivo_diccionario", "")
        if not ruta_csv_default or not os.path.exists(ruta_csv_default):
            messagebox.showwarning("Advertencia", f"No se encontró el diccionario en:\n{ruta_csv_default}\n\nPor favor, configúralo en la pestaña 'Configuración'.")
            self.notebook.select(1) # Cambia a la pestaña de configuración
            return

        diccionario = self.cargar_diccionario(ruta_csv_default)
        if not diccionario:
            messagebox.showinfo("Información", "El diccionario está vacío o hubo un error al leerlo.")
            return
            
        carpeta_original = os.path.join(self.carpeta_actual, "original")
        carpeta_limpio = os.path.join(self.carpeta_actual, "limpio")
        
        try:
            os.makedirs(carpeta_original, exist_ok=True)
            os.makedirs(carpeta_limpio, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron crear las subcarpetas:\n{e}")
            return
            
        procesados = 0
        archivos_seleccionados = [self.listbox.get(i) for i in indices_seleccionados]
        archivos_a_procesar = []
        extensiones_limpiables = {'.txt', '.srt', '.csv', '.vtt'}
        
        for archivo in archivos_seleccionados:
            ruta_entrada = os.path.join(self.carpeta_actual, archivo)
            if archivo.lower().endswith('.zip'):
                try:
                    with zipfile.ZipFile(ruta_entrada, 'r') as zip_ref:
                        for info in zip_ref.infolist():
                            _, ext = os.path.splitext(info.filename)
                            if ext.lower() in extensiones_limpiables and not info.is_dir():
                                info.filename = os.path.basename(info.filename)
                                if info.filename:
                                    zip_ref.extract(info, self.carpeta_actual)
                                    archivos_a_procesar.append(info.filename)
                    ruta_destino_zip = os.path.join(carpeta_original, archivo)
                    if os.path.exists(ruta_destino_zip):
                        os.remove(ruta_destino_zip)
                    shutil.move(ruta_entrada, ruta_destino_zip)
                except Exception as e:
                    messagebox.showerror("Error", f"Error al extraer {archivo}:\n{e}")
            else:
                archivos_a_procesar.append(archivo)
                
        archivos_a_procesar = list(set(archivos_a_procesar))
        
        for archivo in archivos_a_procesar:
            ruta_entrada_actual = os.path.join(self.carpeta_actual, archivo)
            ruta_salida = os.path.join(carpeta_limpio, archivo)
            ruta_destino_original = os.path.join(carpeta_original, archivo)
            
            if self.limpiar_archivo(ruta_entrada_actual, ruta_salida, diccionario):
                try:
                    if os.path.exists(ruta_destino_original):
                        os.remove(ruta_destino_original)
                    shutil.move(ruta_entrada_actual, ruta_destino_original)
                    procesados += 1
                except Exception as e:
                    messagebox.showerror("Error", f"Error al mover el archivo {archivo} a originales:\n{e}")
                    
        if procesados > 0:
            messagebox.showinfo("Proceso Terminado", f"¡Éxito! Se procesaron {procesados} archivo(s).\n\nOriginales movidos a 'original\\'\nLimpios guardados en 'limpio\\'")
            self.actualizar_lista_archivos() # Refresca la lista para quitar los que ya se movieron

def main():
    root = tk.Tk()
    app = TurboScribeCleanerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
