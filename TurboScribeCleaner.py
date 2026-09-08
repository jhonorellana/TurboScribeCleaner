import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import re
import os
import shutil
import sys
import json
import zipfile
import docx
from datetime import datetime
import webbrowser

class TurboScribeCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TurboScribe Limpiador Automático")
        self.root.geometry("850x650")
        self.root.resizable(True, True)
        
        self.config_file = self.obtener_ruta_config()
        self.config = self.cargar_configuracion()
        
        self.carpeta_actual = self.config.get("carpeta_origen", "")
        self.carpeta_salida_actual = self.config.get("carpeta_salida", "")
        self.archivos_en_carpeta = []
        self.archivos_limpios_chatgpt = []
        self.archivos_datos = []
        self.archivos_limpios_datos = []
        
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
        self.tab_chatgpt = ttk.Frame(self.notebook)
        self.tab_configuracion = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_limpieza, text="✨ Limpieza")
        self.notebook.add(self.tab_chatgpt, text="🤖 Integración ChatGPT")
        self.notebook.add(self.tab_configuracion, text="⚙️ Configuración")
        
        self.construir_tab_limpieza()
        self.construir_tab_chatgpt()
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
        lbl_info.pack(pady=(10, 10))
        
        frame_botones_sup = tk.Frame(self.tab_limpieza, bg=bg_color)
        frame_botones_sup.pack(fill=tk.X, padx=20, pady=(0, 5))

        btn_carpeta = tk.Button(frame_botones_sup, text="📁 Seleccionar Carpeta", command=self.seleccionar_carpeta, font=("Segoe UI", 10), bg="#2196F3", fg="white", relief=tk.FLAT, padx=15, pady=5, cursor="hand2")
        btn_carpeta.pack(side=tk.LEFT, padx=(0, 10))

        btn_refrescar = tk.Button(frame_botones_sup, text="🔄 Refrescar", command=self.actualizar_lista_archivos, font=("Segoe UI", 10), bg="#FF9800", fg="white", relief=tk.FLAT, padx=15, pady=5, cursor="hand2")
        btn_refrescar.pack(side=tk.LEFT)
        
        frame_buscar = tk.Frame(frame_botones_sup, bg=bg_color)
        frame_buscar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(20, 0))
        tk.Label(frame_buscar, text="🔍 Buscar:", bg=bg_color, font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 5))
        self.entry_buscar = tk.Entry(frame_buscar, font=("Segoe UI", 10))
        self.entry_buscar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.entry_buscar.bind("<KeyRelease>", self.filtrar_lista)
        
        self.lbl_ruta = tk.Label(self.tab_limpieza, text="Carpeta: Ninguna seleccionada", font=("Segoe UI", 9), bg=bg_color, fg="#666666")
        self.lbl_ruta.pack(pady=(5, 5))
        
        frame_lista = tk.Frame(self.tab_limpieza, bg=bg_color)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        self.scrollbar = tk.Scrollbar(frame_lista)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columnas = ("Nombre", "Fecha", "Tamaño")
        self.tree = ttk.Treeview(frame_lista, columns=columnas, show="headings", yscrollcommand=self.scrollbar.set, selectmode="extended")
        
        self.tree.heading("Nombre", text="Nombre", command=lambda: self.treeview_sort_column(self.tree, "Nombre", False))
        self.tree.heading("Fecha", text="Fecha de modificación", command=lambda: self.treeview_sort_column(self.tree, "Fecha", False))
        self.tree.heading("Tamaño", text="Tamaño", command=lambda: self.treeview_sort_column(self.tree, "Tamaño", False))
        
        self.tree.column("Nombre", width=400, minwidth=200)
        self.tree.column("Fecha", width=150, minwidth=100, anchor=tk.CENTER)
        self.tree.column("Tamaño", width=100, minwidth=80, anchor=tk.E)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.tree.yview)
        
        frame_botones_inf = tk.Frame(self.tab_limpieza, bg=bg_color)
        frame_botones_inf.pack(pady=(0, 20))

        btn_limpiar = tk.Button(frame_botones_inf, text="✨ Iniciar Limpieza", command=self.procesar_seleccion, font=("Segoe UI", 11, "bold"), bg="#4CAF50", fg="white", relief=tk.FLAT, padx=15, pady=8, cursor="hand2")
        btn_limpiar.pack(side=tk.LEFT, padx=(0, 10))

        btn_abrir_salida = tk.Button(frame_botones_inf, text="📂 Abrir Carpeta de Salida", command=self.abrir_carpeta_salida, font=("Segoe UI", 11, "bold"), bg="#9C27B0", fg="white", relief=tk.FLAT, padx=15, pady=8, cursor="hand2")
        btn_abrir_salida.pack(side=tk.LEFT)

    def construir_tab_configuracion(self):
        bg_color = "#f4f4f9"
        
        lbl_info = tk.Label(self.tab_configuracion, text="Parámetros por defecto de la aplicación", font=("Segoe UI", 11, "bold"), bg=bg_color, fg="#333333")
        lbl_info.pack(pady=(20, 20))
        
        # --- Carpeta Origen ---
        frame_origen = tk.Frame(self.tab_configuracion, bg=bg_color)
        frame_origen.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(frame_origen, text="Carpeta de Origen por Defecto:", bg=bg_color, font=("Segoe UI", 10)).pack(anchor="w")
        
        self.entry_carpeta_origen = tk.Entry(frame_origen, font=("Segoe UI", 10))
        self.entry_carpeta_origen.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.entry_carpeta_origen.insert(0, self.config.get("carpeta_origen", ""))
        
        btn_browse_origen = tk.Button(frame_origen, text="Examinar", command=self.buscar_carpeta_origen_config, bg="#e0e0e0", relief=tk.FLAT, cursor="hand2")
        btn_browse_origen.pack(side=tk.RIGHT)
        
        # --- Carpeta Salida ---
        frame_salida = tk.Frame(self.tab_configuracion, bg=bg_color)
        frame_salida.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(frame_salida, text="Carpeta de Salida por Defecto:", bg=bg_color, font=("Segoe UI", 10)).pack(anchor="w")
        
        self.entry_carpeta_salida = tk.Entry(frame_salida, font=("Segoe UI", 10))
        self.entry_carpeta_salida.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.entry_carpeta_salida.insert(0, self.config.get("carpeta_salida", ""))
        
        btn_browse_salida = tk.Button(frame_salida, text="Examinar", command=self.buscar_carpeta_salida_config, bg="#e0e0e0", relief=tk.FLAT, cursor="hand2")
        btn_browse_salida.pack(side=tk.RIGHT)
        
        # --- Archivo Diccionario ---
        frame_dicc = tk.Frame(self.tab_configuracion, bg=bg_color)
        frame_dicc.pack(fill=tk.X, padx=20, pady=5)
        
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
        
        # --- Archivo Proyectos ChatGPT ---
        frame_chatgpt = tk.Frame(self.tab_configuracion, bg=bg_color)
        frame_chatgpt.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(frame_chatgpt, text="Archivo Proyectos ChatGPT (CSV):", bg=bg_color, font=("Segoe UI", 10)).pack(anchor="w")
        
        self.entry_diccionario_chatgpt = tk.Entry(frame_chatgpt, font=("Segoe UI", 10))
        self.entry_diccionario_chatgpt.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        dicc_chatgpt_guardado = self.config.get("archivo_proyectos_chatgpt", "")
        self.entry_diccionario_chatgpt.insert(0, dicc_chatgpt_guardado)
        
        btn_browse_chatgpt = tk.Button(frame_chatgpt, text="Examinar", command=self.buscar_diccionario_chatgpt_config, bg="#e0e0e0", relief=tk.FLAT, cursor="hand2")
        btn_browse_chatgpt.pack(side=tk.RIGHT)
        
        # --- Boton Guardar ---
        btn_guardar = tk.Button(self.tab_configuracion, text="💾 Guardar Configuración", command=self.guardar_cambios_configuracion, font=("Segoe UI", 10, "bold"), bg="#2196F3", fg="white", relief=tk.FLAT, padx=15, pady=8, cursor="hand2")
        btn_guardar.pack(pady=40)

    def buscar_carpeta_origen_config(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de origen por defecto")
        if carpeta:
            self.entry_carpeta_origen.delete(0, tk.END)
            self.entry_carpeta_origen.insert(0, carpeta)
            
    def buscar_carpeta_salida_config(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida por defecto")
        if carpeta:
            self.entry_carpeta_salida.delete(0, tk.END)
            self.entry_carpeta_salida.insert(0, carpeta)
            
    def buscar_diccionario_config(self):
        archivo = filedialog.askopenfilename(title="Seleccionar Diccionario (CSV)", filetypes=[("Archivos CSV", "*.csv")])
        if archivo:
            self.entry_diccionario.delete(0, tk.END)
            self.entry_diccionario.insert(0, archivo)

    def buscar_diccionario_chatgpt_config(self):
        archivo = filedialog.askopenfilename(title="Seleccionar Proyectos ChatGPT (CSV)", filetypes=[("Archivos CSV", "*.csv")])
        if archivo:
            self.entry_diccionario_chatgpt.delete(0, tk.END)
            self.entry_diccionario_chatgpt.insert(0, archivo)

    def guardar_cambios_configuracion(self):
        self.config["carpeta_origen"] = self.entry_carpeta_origen.get()
        self.config["carpeta_salida"] = self.entry_carpeta_salida.get()
        self.config["archivo_diccionario"] = self.entry_diccionario.get()
        self.config["archivo_proyectos_chatgpt"] = self.entry_diccionario_chatgpt.get()
        self.guardar_configuracion()
        
        # Sincronizar con estado actual
        if os.path.exists(self.config["carpeta_origen"]):
            self.carpeta_actual = self.config["carpeta_origen"]
            self.lbl_ruta.config(text=f"Carpeta: {self.carpeta_actual}")
            self.actualizar_lista_archivos()
            
        self.carpeta_salida_actual = self.config["carpeta_salida"]

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

    def extraer_fecha(self, texto):
        # Formatos numericos: DD/MM/YYYY, YYYY/MM/DD, DD-MM-YYYY, YYYY-MM-DD
        patron_numerico = re.search(r'\b(\d{2})[-/](\d{2})[-/](\d{4})\b', texto)
        if patron_numerico:
            p1, p2, p3 = patron_numerico.groups()
            if int(p2) <= 12:
                try:
                    fecha = datetime(int(p3), int(p2), int(p1))
                    return fecha.strftime('%Y_%m_%d')
                except ValueError:
                    pass
        
        patron_numerico2 = re.search(r'\b(\d{4})[-/](\d{2})[-/](\d{2})\b', texto)
        if patron_numerico2:
            p1, p2, p3 = patron_numerico2.groups()
            try:
                fecha = datetime(int(p1), int(p2), int(p3))
                return fecha.strftime('%Y_%m_%d')
            except ValueError:
                pass

        # Formatos en texto
        meses = {
            'enero': 1, 'january': 1, 'febrero': 2, 'february': 2, 'marzo': 3, 'march': 3,
            'abril': 4, 'april': 4, 'mayo': 5, 'may': 5, 'junio': 6, 'june': 6,
            'julio': 7, 'july': 7, 'agosto': 8, 'august': 8, 'septiembre': 9, 'september': 9,
            'octubre': 10, 'october': 10, 'noviembre': 11, 'november': 11, 'diciembre': 12, 'december': 12
        }
        
        patron_texto = r'\b(\d{1,2})\s*(?:de\s+)?(' + '|'.join(meses.keys()) + r')\s*(?:de\s+|,\s*)?(\d{4})\b'
        match = re.search(patron_texto, texto, re.IGNORECASE)
        if match:
            dia, mes_str, anio = match.groups()
            mes = meses[mes_str.lower()]
            try:
                fecha = datetime(int(anio), int(mes), int(dia))
                return fecha.strftime('%Y_%m_%d')
            except ValueError:
                pass

        patron_texto2 = r'\b(' + '|'.join(meses.keys()) + r')\s+(\d{1,2})\s*(?:de\s+|,\s*)?(\d{4})\b'
        match2 = re.search(patron_texto2, texto, re.IGNORECASE)
        if match2:
            mes_str, dia, anio = match2.groups()
            mes = meses[mes_str.lower()]
            try:
                fecha = datetime(int(anio), int(mes), int(dia))
                return fecha.strftime('%Y_%m_%d')
            except ValueError:
                pass

        return datetime.now().strftime('%Y_%m_%d')

    def limpiar_archivo(self, ruta_entrada, carpeta_limpio, carpeta_original, diccionario):
        try:
            _, ext = os.path.splitext(ruta_entrada)
            archivo_original = os.path.basename(ruta_entrada)
            nombre_base = os.path.splitext(archivo_original)[0]
            
            if ext.lower() == '.docx':
                doc = docx.Document(ruta_entrada)
                parrafos = [p.text for p in doc.paragraphs]
                contenido = '\n'.join(parrafos)
                ext_salida = '.txt'
            else:
                with open(ruta_entrada, 'r', encoding='utf-8') as f:
                    contenido = f.read()
                ext_salida = ext

            if re.match(r'^\d{4}[_-]\d{2}[_-]\d{2}', nombre_base):
                nuevo_nombre_archivo = f"{nombre_base}{ext_salida}"
                nuevo_nombre_original = archivo_original
            else:
                fecha_prefijo = self.extraer_fecha(contenido)
                nuevo_nombre_archivo = f"{fecha_prefijo}_{nombre_base}{ext_salida}"
                nuevo_nombre_original = f"{fecha_prefijo}_{archivo_original}"

            ruta_salida = os.path.join(carpeta_limpio, nuevo_nombre_archivo)

            claves_ordenadas = sorted(diccionario.keys(), key=len, reverse=True)
            
            for incorrecto in claves_ordenadas:
                correcto = diccionario[incorrecto]
                patron = re.compile(r'\b' + re.escape(incorrecto) + r'\b', re.IGNORECASE)
                contenido = patron.sub(correcto.replace('\\', '\\\\'), contenido)

            with open(ruta_salida, 'w', encoding='utf-8') as f:
                f.write(contenido)
                
            ruta_destino_original = os.path.join(carpeta_original, nuevo_nombre_original)
            
            if os.path.exists(ruta_destino_original):
                os.remove(ruta_destino_original)
            shutil.move(ruta_entrada, ruta_destino_original)
                
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar el archivo {os.path.basename(ruta_entrada)}:\n{e}")
            return False

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de origen")
        if not carpeta:
            return
            
        self.carpeta_actual = carpeta
        self.lbl_ruta.config(text=f"Carpeta: {self.carpeta_actual}")
        
        self.actualizar_lista_archivos()
        
    def actualizar_lista_archivos(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.archivos_datos = []
        
        if not self.carpeta_actual or not os.path.exists(self.carpeta_actual):
            return
            
        extensiones_validas = {'.txt', '.srt', '.csv', '.vtt', '.zip', '.docx'}
        
        for archivo in os.listdir(self.carpeta_actual):
            ruta_completa = os.path.join(self.carpeta_actual, archivo)
            if os.path.isfile(ruta_completa):
                _, ext = os.path.splitext(archivo)
                if ext.lower() in extensiones_validas:
                    try:
                        tiempo = os.path.getmtime(ruta_completa)
                        fecha_str = datetime.fromtimestamp(tiempo).strftime('%Y-%m-%d %H:%M')
                        tamano = os.path.getsize(ruta_completa)
                        tamano_str = self.formatear_tamano(tamano)
                    except Exception:
                        tiempo = 0
                        fecha_str = ""
                        tamano_str = ""
                        
                    self.archivos_datos.append({
                        "nombre": archivo,
                        "fecha_str": fecha_str,
                        "tamano_str": tamano_str,
                        "tiempo": tiempo
                    })
                    
        self.archivos_datos.sort(key=lambda x: x["tiempo"], reverse=True)
        self.filtrar_lista()

    def procesar_seleccion(self):
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo de la lista.")
            return
            
        ruta_csv_default = self.config.get("archivo_diccionario", "")
        if not ruta_csv_default or not os.path.exists(ruta_csv_default):
            messagebox.showwarning("Advertencia", f"No se encontró el diccionario en:\n{ruta_csv_default}\n\nPor favor, configúralo en la pestaña 'Configuración'.")
            self.notebook.select(1) # Cambia a la pestaña de configuración
            return

        if not self.carpeta_salida_actual or not os.path.exists(self.carpeta_salida_actual):
            messagebox.showwarning("Advertencia", "No has configurado una Carpeta de Salida válida.\n\nPor favor, configúrala en la pestaña 'Configuración'.")
            self.notebook.select(1)
            return

        diccionario = self.cargar_diccionario(ruta_csv_default)
        if not diccionario:
            messagebox.showinfo("Información", "El diccionario está vacío o hubo un error al leerlo.")
            return
            
        carpeta_original = os.path.join(self.carpeta_salida_actual, "Original")
        carpeta_limpio = self.carpeta_salida_actual
        
        try:
            os.makedirs(carpeta_original, exist_ok=True)
            # carpeta_limpio (que es carpeta_salida_actual) ya validamos que existe, pero por si acaso
            os.makedirs(carpeta_limpio, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron crear las subcarpetas:\n{e}")
            return
            
        procesados = 0
        archivos_seleccionados = [self.tree.item(item, "values")[0] for item in seleccion]
        archivos_a_procesar = []
        extensiones_limpiables = {'.txt', '.srt', '.csv', '.vtt', '.docx'}
        
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
            if self.limpiar_archivo(ruta_entrada_actual, carpeta_limpio, carpeta_original, diccionario):
                procesados += 1
                    
        if procesados > 0:
            messagebox.showinfo("Proceso Terminado", f"¡Éxito! Se procesaron {procesados} archivo(s).\n\nOriginales movidos a la subcarpeta 'Original'\nLimpios guardados en la raíz de la carpeta de salida.")
            self.actualizar_lista_archivos() # Refresca la lista para quitar los que ya se movieron

    def abrir_carpeta_salida(self):
        if not self.carpeta_salida_actual or not os.path.exists(self.carpeta_salida_actual):
            messagebox.showwarning("Advertencia", "La carpeta de salida configurada no es válida o no existe.")
            return
            
        try:
            os.startfile(self.carpeta_salida_actual)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{e}")

    def construir_tab_chatgpt(self):
        bg_color = "#f4f4f9"
        
        lbl_info = tk.Label(self.tab_chatgpt, text="Selecciona los archivos limpios que deseas subir a ChatGPT.", font=("Segoe UI", 10), bg=bg_color, fg="#666666")
        lbl_info.pack(pady=(10, 10))
        
        frame_botones_sup = tk.Frame(self.tab_chatgpt, bg=bg_color)
        frame_botones_sup.pack(fill=tk.X, padx=20, pady=(0, 5))

        btn_refrescar = tk.Button(frame_botones_sup, text="🔄 Refrescar Lista", command=self.actualizar_lista_archivos_limpios, font=("Segoe UI", 10), bg="#FF9800", fg="white", relief=tk.FLAT, padx=15, pady=5, cursor="hand2")
        btn_refrescar.pack(side=tk.LEFT)
        
        frame_buscar = tk.Frame(frame_botones_sup, bg=bg_color)
        frame_buscar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(20, 0))
        tk.Label(frame_buscar, text="🔍 Buscar:", bg=bg_color, font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 5))
        self.entry_buscar_chatgpt = tk.Entry(frame_buscar, font=("Segoe UI", 10))
        self.entry_buscar_chatgpt.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.entry_buscar_chatgpt.bind("<KeyRelease>", self.filtrar_lista_chatgpt)
        
        self.lbl_ruta_chatgpt = tk.Label(self.tab_chatgpt, text="Carpeta Salida: Ninguna configurada", font=("Segoe UI", 9), bg=bg_color, fg="#666666")
        self.lbl_ruta_chatgpt.pack(pady=(5, 5))
        
        frame_lista = tk.Frame(self.tab_chatgpt, bg=bg_color)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        self.scrollbar_chatgpt = tk.Scrollbar(frame_lista)
        self.scrollbar_chatgpt.pack(side=tk.RIGHT, fill=tk.Y)
        
        columnas = ("Nombre", "Fecha", "Tamaño")
        self.tree_chatgpt = ttk.Treeview(frame_lista, columns=columnas, show="headings", yscrollcommand=self.scrollbar_chatgpt.set, selectmode="extended")
        
        self.tree_chatgpt.heading("Nombre", text="Nombre", command=lambda: self.treeview_sort_column(self.tree_chatgpt, "Nombre", False))
        self.tree_chatgpt.heading("Fecha", text="Fecha de modificación", command=lambda: self.treeview_sort_column(self.tree_chatgpt, "Fecha", False))
        self.tree_chatgpt.heading("Tamaño", text="Tamaño", command=lambda: self.treeview_sort_column(self.tree_chatgpt, "Tamaño", False))
        
        self.tree_chatgpt.column("Nombre", width=400, minwidth=200)
        self.tree_chatgpt.column("Fecha", width=150, minwidth=100, anchor=tk.CENTER)
        self.tree_chatgpt.column("Tamaño", width=100, minwidth=80, anchor=tk.E)
        
        self.tree_chatgpt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_chatgpt.config(command=self.tree_chatgpt.yview)
        
        frame_botones_inf = tk.Frame(self.tab_chatgpt, bg=bg_color)
        frame_botones_inf.pack(pady=(15, 20))
        
        btn_enviar = tk.Button(frame_botones_inf, text="🚀 Enviar a ChatGPT", command=self.enviar_a_chatgpt, font=("Segoe UI", 11, "bold"), bg="#9C27B0", fg="white", relief=tk.FLAT, padx=20, pady=10, cursor="hand2")
        btn_enviar.pack(side=tk.LEFT, padx=10)
        
        btn_abrir_cargado = tk.Button(frame_botones_inf, text="📂 Abrir Carpeta 'Cargado'", command=self.abrir_carpeta_cargado, font=("Segoe UI", 11, "bold"), bg="#2196F3", fg="white", relief=tk.FLAT, padx=20, pady=10, cursor="hand2")
        btn_abrir_cargado.pack(side=tk.LEFT, padx=10)

        self.actualizar_lista_archivos_limpios()

    def actualizar_lista_archivos_limpios(self):
        for item in self.tree_chatgpt.get_children():
            self.tree_chatgpt.delete(item)
        self.archivos_limpios_datos = []
        
        if not self.carpeta_salida_actual or not os.path.exists(self.carpeta_salida_actual):
            self.lbl_ruta_chatgpt.config(text="Carpeta Salida: Ninguna configurada")
            return
            
        self.lbl_ruta_chatgpt.config(text=f"Carpeta Salida: {self.carpeta_salida_actual}")
        
        for archivo in os.listdir(self.carpeta_salida_actual):
            ruta_completa = os.path.join(self.carpeta_salida_actual, archivo)
            if os.path.isfile(ruta_completa) and archivo.lower().endswith('.txt'):
                try:
                    tiempo = os.path.getmtime(ruta_completa)
                    fecha_str = datetime.fromtimestamp(tiempo).strftime('%Y-%m-%d %H:%M')
                    tamano = os.path.getsize(ruta_completa)
                    tamano_str = self.formatear_tamano(tamano)
                except Exception:
                    tiempo = 0
                    fecha_str = ""
                    tamano_str = ""
                
                self.archivos_limpios_datos.append({
                    "nombre": archivo,
                    "fecha_str": fecha_str,
                    "tamano_str": tamano_str,
                    "tiempo": tiempo
                })
                
        self.archivos_limpios_datos.sort(key=lambda x: x["tiempo"], reverse=True)
        self.filtrar_lista_chatgpt()

    def cargar_proyectos_chatgpt(self):
        ruta_csv = self.config.get("archivo_proyectos_chatgpt", "")
        proyectos = []
        if not ruta_csv or not os.path.exists(ruta_csv):
            return proyectos
            
        try:
            with open(ruta_csv, mode='r', encoding='utf-8-sig') as f:
                linea = f.readline()
                f.seek(0)
                delim = ';' if ';' in linea else ','
                reader = csv.reader(f, delimiter=delim)
                next(reader, None) # Saltar cabecera
                
                for row in reader:
                    if len(row) >= 2:
                        nombre = row[0].strip()
                        url = row[1].strip()
                        palabra = row[2].strip().lower() if len(row) > 2 else ""
                        
                        if nombre and url:
                            proyectos.append({
                                "nombre": nombre,
                                "url": url,
                                "clave": palabra
                            })
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo de proyectos ChatGPT:\n{e}")
        return proyectos

    def abrir_selector_proyecto(self, proyectos, callback):
        modal = tk.Toplevel(self.root)
        modal.title("Seleccionar Proyecto ChatGPT")
        modal.geometry("450x300")
        modal.grab_set()
        
        lbl = tk.Label(modal, text="No se detectó palabra clave en el archivo.\nSelecciona el proyecto manualmente:", font=("Segoe UI", 10))
        lbl.pack(pady=10)
        
        frame_lista = tk.Frame(modal)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        scrollbar = tk.Scrollbar(frame_lista)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(frame_lista, yscrollcommand=scrollbar.set, font=("Segoe UI", 10), selectmode=tk.SINGLE, selectbackground="#9C27B0")
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        proyectos_unicos = {}
        for p in proyectos:
            proyectos_unicos[p["nombre"]] = p["url"]
            
        nombres_ordenados = sorted(proyectos_unicos.keys())
        
        for nombre in nombres_ordenados:
            listbox.insert(tk.END, nombre)
            
        def on_continuar():
            seleccion = listbox.curselection()
            if not seleccion:
                messagebox.showwarning("Advertencia", "Selecciona un proyecto.")
                return
            idx = seleccion[0]
            nombre_sel = nombres_ordenados[idx]
            url = proyectos_unicos[nombre_sel]
            modal.destroy()
            callback(url)
            
        def on_cancelar():
            modal.destroy()
            
        frame_btn = tk.Frame(modal)
        frame_btn.pack(pady=10)
        
        btn_ok = tk.Button(frame_btn, text="Continuar", command=on_continuar, bg="#9C27B0", fg="white", font=("Segoe UI", 10, "bold"), width=10, relief=tk.FLAT, cursor="hand2")
        btn_ok.pack(side=tk.LEFT, padx=10)
        
        btn_cancel = tk.Button(frame_btn, text="Cancelar", command=on_cancelar, bg="#f44336", fg="white", font=("Segoe UI", 10, "bold"), width=10, relief=tk.FLAT, cursor="hand2")
        btn_cancel.pack(side=tk.LEFT, padx=10)

    def enviar_a_chatgpt(self):
        seleccion = self.tree_chatgpt.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo limpio de la lista.")
            return
            
        archivo_seleccionado = self.tree_chatgpt.item(seleccion[0], "values")[0]
        ruta_archivo = os.path.join(self.carpeta_salida_actual, archivo_seleccionado)
        
        proyectos = self.cargar_proyectos_chatgpt()
        if not proyectos:
            messagebox.showwarning("Advertencia", "No hay proyectos configurados en el archivo CSV.")
            return
            
        url_destino = None
        nombre_lower = archivo_seleccionado.lower()
        
        for p in proyectos:
            if p["clave"] and p["clave"] in nombre_lower:
                url_destino = p["url"]
                break
                
        def ejecutar_envio(url):
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append("Analiza paso a paso el archivo adjunto")
                self.root.update()
                
                webbrowser.open(url)
                
                carpeta_cargado = os.path.join(self.carpeta_salida_actual, "Cargado")
                os.makedirs(carpeta_cargado, exist_ok=True)
                
                ruta_destino = os.path.join(carpeta_cargado, archivo_seleccionado)
                if os.path.exists(ruta_destino):
                    os.remove(ruta_destino)
                shutil.move(ruta_archivo, ruta_destino)
                
                self.actualizar_lista_archivos_limpios()
                
                messagebox.showinfo("Enviado", f"El archivo '{archivo_seleccionado}' se movió a 'Cargado'.\n\nEl navegador se ha abierto en el Proyecto.\n¡La instrucción ya está en tu portapapeles! (Presiona Ctrl+V para pegar)")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al enviar a ChatGPT:\n{e}")

        if url_destino:
            ejecutar_envio(url_destino)
        else:
            self.abrir_selector_proyecto(proyectos, ejecutar_envio)


    def formatear_tamano(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.0f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"

    def treeview_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        if col == "Tamaño":
            def sort_key(item):
                val = item[0]
                if "KB" in val: return float(val.replace(" KB", "")) * 1024
                if "MB" in val: return float(val.replace(" MB", "")) * 1024 * 1024
                if " B" in val: return float(val.replace(" B", ""))
                return 0
            l.sort(key=sort_key, reverse=reverse)
        else:
            l.sort(key=lambda t: t[0].lower(), reverse=reverse)
            
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def filtrar_lista(self, event=None):
        query = self.entry_buscar.get().lower()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for archivo in self.archivos_datos:
            if query in archivo["nombre"].lower():
                self.tree.insert("", tk.END, values=(archivo["nombre"], archivo["fecha_str"], archivo["tamano_str"]))

    def filtrar_lista_chatgpt(self, event=None):
        query = self.entry_buscar_chatgpt.get().lower()
        for item in self.tree_chatgpt.get_children():
            self.tree_chatgpt.delete(item)
        for archivo in self.archivos_limpios_datos:
            if query in archivo["nombre"].lower():
                self.tree_chatgpt.insert("", tk.END, values=(archivo["nombre"], archivo["fecha_str"], archivo["tamano_str"]))

    def abrir_carpeta_cargado(self):
        if not self.carpeta_salida_actual or not os.path.exists(self.carpeta_salida_actual):
            messagebox.showwarning("Advertencia", "La carpeta de salida configurada no es válida o no existe.")
            return
            
        carpeta_cargado = os.path.join(self.carpeta_salida_actual, "Cargado")
        if not os.path.exists(carpeta_cargado):
            os.makedirs(carpeta_cargado, exist_ok=True)
            
        try:
            os.startfile(carpeta_cargado)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\\n{e}")

def main():
    root = tk.Tk()
    app = TurboScribeCleanerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
