import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import time
from typing import List, Dict
import sys
import os

# Agregar directorio al path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

class EmailSenderGUI:
    """GUI Email Sender Pro - PARTE 1"""
    
    def __init__(self):
        print("🔧 Inicializando GUI...")
        
        # Crear ventana
        self.root = tk.Tk()
        self.root.title("📧 Email Sender Pro - FUNCIONAL")
        
        # Variables
        self.enviando = False
        self.correos_procesados = []
        
        # Configurar
        self._configurar_ventana()
        self._definir_colores()
        self._aplicar_tema_seguro()
        self._inicializar_managers_seguro()
        
        print("✅ GUI inicializada")
    
    def _configurar_ventana(self):
        """Configurar ventana"""
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            width = int(screen_width * 0.85)
            height = int(screen_height * 0.85)
            
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            self.root.geometry(f"{width}x{height}+{x}+{y}")
            self.root.minsize(1200, 800)
            self.root.resizable(True, True)
            
            print(f"📐 Ventana: {width}x{height}")
        except Exception as e:
            print(f"⚠️ Error ventana: {e}")
            self.root.geometry("1200x800")
    
    def _definir_colores(self):
        """Colores Dracula"""
        self.colors = {
            'bg': '#282a36',
            'current_line': '#44475a',
            'selection': '#44475a',
            'foreground': '#f8f8f2',
            'comment': '#6272a4',
            'cyan': '#8be9fd',
            'green': '#50fa7b',
            'orange': '#ffb86c',
            'pink': '#ff79c6',
            'purple': '#bd93f9',
            'red': '#ff5555',
            'yellow': '#f1fa8c'
        }
    
    def _aplicar_tema_seguro(self):
        """Aplicar tema SIN ERRORES"""
        try:
            self.root.configure(bg=self.colors['bg'])
            
            self.style = ttk.Style()
            
            try:
                self.style.theme_use('clam')
            except:
                print("⚠️ Usando tema default")
            
            try:
                self.style.configure('TButton',
                                   background=self.colors['current_line'],
                                   foreground=self.colors['foreground'])
                
                self.style.map('TButton',
                              background=[('active', self.colors['purple'])])
                
                print("✅ Estilos aplicados")
            except Exception as e:
                print(f"⚠️ Error estilos: {e}")
            
        except Exception as e:
            print(f"⚠️ Error tema: {e}")
    
    def _inicializar_managers_seguro(self):
        """Inicializar managers"""
        print("🔧 Inicializando managers...")
        
        # ExcelManager
        try:
            from excel_manager import ExcelManager
            self.excel_mgr = ExcelManager()
            print("✅ ExcelManager OK")
        except Exception as e:
            print(f"❌ ExcelManager: {e}")
            self.excel_mgr = None
        
        # FileManager
        try:
            from file_manager import FileManager
            self.file_mgr = FileManager()
            print("✅ FileManager OK")
        except Exception as e:
            print(f"❌ FileManager: {e}")
            self.file_mgr = None
        
        # EmailProcessor
        try:
            from email_processor import EmailProcessor
            self.email_processor = EmailProcessor()
            print("✅ EmailProcessor OK")
        except Exception as e:
            print(f"❌ EmailProcessor: {e}")
            self.email_processor = None
        
        # EmailSender ESPECIAL
        try:
            from email_sender import EmailSender
            self.email_sender = EmailSender()
            print("✅ EmailSender REAL OK")
        except Exception as e:
            print(f"❌ EmailSender error: {e}")
            print("🔧 Creando EmailSender funcional...")
            self.email_sender = self._crear_email_sender_funcional()
        
        print("✅ Managers listos")
    
    def _crear_email_sender_funcional(self):
        """EmailSender funcional integrado"""
        class EmailSenderFuncional:
            def __init__(self):
                self.conectado = False
                print("📧 EmailSender funcional creado")
            
            def conectar_outlook(self):
                try:
                    import win32com.client
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    namespace = outlook.GetNamespace("MAPI")
                    inbox = namespace.GetDefaultFolder(6)
                    
                    self.conectado = True
                    return {
                        'exitoso': True,
                        'mensaje': 'Conectado a Outlook',
                        'cuenta': 'Outlook OK'
                    }
                except Exception as e:
                    return {
                        'exitoso': False,
                        'mensaje': f'Error Outlook: {str(e)}',
                        'sugerencia': 'Abre Outlook primero'
                    }
            
            def enviar_correo(self, correo_data, adjuntos=None):
                try:
                    import win32com.client
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    
                    mail = outlook.CreateItem(0)
                    mail.To = correo_data['email']
                    mail.Subject = correo_data['asunto']
                    mail.Body = correo_data['contenido']
                    
                    if adjuntos:
                        for adjunto in adjuntos:
                            if os.path.exists(adjunto):
                                mail.Attachments.Add(adjunto)
                    
                    mail.Send()
                    
                    return {
                        'exitoso': True,
                        'timestamp': time.strftime('%H:%M:%S'),
                        'email': correo_data['email'],
                        'nombre': correo_data.get('nombre', 'Sin nombre')
                    }
                    
                except Exception as e:
                    return {
                        'exitoso': False,
                        'error': str(e),
                        'email': correo_data['email'],
                        'nombre': correo_data.get('nombre', 'Sin nombre')
                    }
            
            def envio_por_lotes(self, correos, adjuntos, callback_progreso=None, detener_callback=None):
                resultados = {
                    'exitosos': [],
                    'fallidos': [],
                    'total_procesados': 0,
                    'inicio': time.time()
                }
                
                for i, correo in enumerate(correos):
                    if detener_callback and detener_callback():
                        break
                    
                    if callback_progreso:
                        progreso = (i / len(correos)) * 100
                        nombre = correo.get('nombre', 'Sin nombre')
                        callback_progreso(progreso, f"Enviando a {nombre} ({i+1}/{len(correos)})")
                    
                    resultado = self.enviar_correo(correo, adjuntos)
                    
                    if resultado['exitoso']:
                        resultados['exitosos'].append(resultado)
                    else:
                        resultados['fallidos'].append(resultado)
                    
                    resultados['total_procesados'] += 1
                    
                    # Pausa entre correos (6 minutos)
                    if i < len(correos) - 1:
                        for segundo in range(360):
                            if detener_callback and detener_callback():
                                break
                            time.sleep(1)
                            
                            if segundo % 30 == 0 and callback_progreso:
                                tiempo_restante = 360 - segundo
                                callback_progreso(progreso, f"Pausa: {tiempo_restante}s")
                
                resultados['fin'] = time.time()
                resultados['duracion'] = f"{resultados['fin'] - resultados['inicio']:.1f}s"
                
                return resultados
        
        return EmailSenderFuncional()
    
    def crear_interfaz(self):
        """Crear interfaz"""
        print("🏗️ Creando interfaz...")
        
        # Frame principal
        self.main_frame = tk.Frame(self.root, bg=self.colors['bg'], padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Grid
        self.main_frame.columnconfigure(0, weight=2)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(0, weight=0)
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=0)
        self.main_frame.rowconfigure(3, weight=0)
        
        # Crear secciones
        self.crear_titulo()
        self.crear_contenido()
        self.crear_botones()
        self.crear_progreso_seguro()
        
        print("✅ Interfaz creada")
    
    def crear_titulo(self):
        """Título"""
        titulo_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        titulo_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 20))
        
        tk.Label(titulo_frame, text="📧 EMAIL SENDER PRO", 
                font=('Arial', 24, 'bold'), bg=self.colors['bg'], fg=self.colors['purple']).pack()
        
        tk.Label(titulo_frame, text="Envío masivo REAL con Outlook", 
                font=('Arial', 11), bg=self.colors['bg'], fg=self.colors['comment']).pack(pady=(5,0))
    
    def crear_contenido(self):
        """Área de contenido"""
        # Izquierda
        left_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        left_frame.grid(row=1, column=0, sticky='nsew', padx=(0, 15))
        left_frame.rowconfigure(0, weight=0)
        left_frame.rowconfigure(1, weight=0) 
        left_frame.rowconfigure(2, weight=1)
        left_frame.columnconfigure(0, weight=1)
        
        self.crear_estado_archivos(left_frame)
        self.crear_campana_activa(left_frame)
        self.crear_vista_previa(left_frame)
        
        # Derecha
        right_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        right_frame.grid(row=1, column=1, sticky='nsew')
        right_frame.rowconfigure(0, weight=0)
        right_frame.rowconfigure(1, weight=1)
        right_frame.columnconfigure(0, weight=1)
        
        self.crear_adjuntos(right_frame)
        self.crear_log(right_frame)
    
    def crear_estado_archivos(self, parent):
        """Estado archivos"""
        frame = tk.LabelFrame(parent, text="📊 Estado de Archivos Excel",
                             font=('Arial', 11, 'bold'), bg=self.colors['bg'], 
                             fg=self.colors['foreground'], bd=2, relief='solid')
        frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        frame.columnconfigure(0, weight=1)
        
        self.text_archivos = scrolledtext.ScrolledText(frame, height=6, font=('Consolas', 9),
                                                      bg=self.colors['current_line'], 
                                                      fg=self.colors['foreground'], relief='flat')
        self.text_archivos.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def crear_campana_activa(self, parent):
        """Campaña activa"""
        frame = tk.LabelFrame(parent, text="🎯 Campaña Activa",
                             font=('Arial', 11, 'bold'), bg=self.colors['bg'], 
                             fg=self.colors['foreground'], bd=2, relief='solid')
        frame.grid(row=1, column=0, sticky='ew', pady=(0, 15))
        frame.columnconfigure(0, weight=1)
        
        info_frame = tk.Frame(frame, bg=self.colors['bg'])
        info_frame.pack(fill=tk.BOTH, padx=15, pady=15)
        info_frame.columnconfigure(0, weight=1)
        
        self.label_campana_nombre = tk.Label(info_frame, text="📋 Campaña: Cargando...",
                                           font=('Arial', 11, 'bold'), bg=self.colors['bg'], 
                                           fg=self.colors['cyan'], anchor='w')
        self.label_campana_nombre.grid(row=0, column=0, sticky='ew', pady=(0,5))
        
        self.label_campana_asunto = tk.Label(info_frame, text="📧 Asunto: Cargando...",
                                           font=('Arial', 10), bg=self.colors['bg'], 
                                           fg=self.colors['foreground'], anchor='w', wraplength=500)
        self.label_campana_asunto.grid(row=1, column=0, sticky='ew', pady=(0,5))
        
        self.label_campana_info = tk.Label(info_frame, text="📝 Info: Cargando...",
                                         font=('Arial', 9), bg=self.colors['bg'], 
                                         fg=self.colors['comment'], anchor='w', wraplength=500)
        self.label_campana_info.grid(row=2, column=0, sticky='ew')
    
    def crear_vista_previa(self, parent):
        """Vista previa"""
        frame = tk.LabelFrame(parent, text="📧 Vista Previa del Primer Correo",
                             font=('Arial', 11, 'bold'), bg=self.colors['bg'], 
                             fg=self.colors['foreground'], bd=2, relief='solid')
        frame.grid(row=2, column=0, sticky='nsew')
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        self.text_preview = scrolledtext.ScrolledText(frame, font=('Consolas', 9),
                                                     bg=self.colors['current_line'], 
                                                     fg=self.colors['foreground'], relief='flat')
        self.text_preview.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def crear_adjuntos(self, parent):
        """Adjuntos"""
        frame = tk.LabelFrame(parent, text="📎 Archivos Adjuntos",
                             font=('Arial', 11, 'bold'), bg=self.colors['bg'], 
                             fg=self.colors['foreground'], bd=2, relief='solid')
        frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        frame.columnconfigure(0, weight=1)
        
        self.text_adjuntos = tk.Text(frame, height=8, font=('Consolas', 8),
                                    bg=self.colors['current_line'], fg=self.colors['orange'], 
                                    relief='flat', wrap=tk.WORD)
        self.text_adjuntos.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def crear_log(self, parent):
        """Log"""
        frame = tk.LabelFrame(parent, text="📋 Log en Tiempo Real",
                             font=('Arial', 11, 'bold'), bg=self.colors['bg'], 
                             fg=self.colors['foreground'], bd=2, relief='solid')
        frame.grid(row=1, column=0, sticky='nsew')
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        self.text_log = scrolledtext.ScrolledText(frame, font=('Consolas', 8),
                                                 bg=self.colors['current_line'], 
                                                 fg=self.colors['orange'], relief='flat')
        self.text_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def crear_botones(self):
        """Botones"""
        buttons_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        buttons_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(20, 0))
        
        for i in range(4):
            buttons_frame.columnconfigure(i, weight=1)
        
        self.btn_actualizar = ttk.Button(buttons_frame, text="🔄 Actualizar Datos", command=self.actualizar_datos)
        self.btn_actualizar.grid(row=0, column=0, padx=5, pady=10, sticky='ew')
        
        self.btn_preview = ttk.Button(buttons_frame, text="👁️ Vista Previa", command=self.vista_previa_completa)
        self.btn_preview.grid(row=0, column=1, padx=5, pady=10, sticky='ew')
        
        self.btn_enviar = ttk.Button(buttons_frame, text="🚀 ENVIAR CORREOS", command=self.enviar_correos)
        self.btn_enviar.grid(row=0, column=2, padx=5, pady=10, sticky='ew')
        
        self.btn_detener = ttk.Button(buttons_frame, text="⏹️ DETENER", state="disabled", command=self.detener_envio)
        self.btn_detener.grid(row=0, column=3, padx=5, pady=10, sticky='ew')
    
    def crear_progreso_seguro(self):
        """Progreso SIN ERRORES"""
        progress_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        progress_frame.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(15, 0))
        
        self.label_estado = tk.Label(progress_frame, text="✅ Sistema listo",
                                   font=('Arial', 11, 'bold'), bg=self.colors['bg'], fg=self.colors['green'])
        self.label_estado.pack(pady=(0, 10))
        
        try:
            self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=400)
            self.progress_bar.pack(fill=tk.X, padx=50)
            print("✅ ProgressBar OK")
        except Exception as e:
            print(f"⚠️ Error ProgressBar: {e}")
            self.progress_canvas = tk.Canvas(progress_frame, height=20, bg=self.colors['current_line'])
            self.progress_canvas.pack(fill=tk.X, padx=50)
            self.progress_bar = None
            print("✅ Canvas como ProgressBar")
    
    def log_mensaje(self, mensaje):
        """Log"""
        try:
            timestamp = time.strftime('%H:%M:%S')
            self.text_log.insert(tk.END, f"[{timestamp}] {mensaje}\n")
            self.text_log.see(tk.END)
            self.root.update_idletasks()
        except Exception as e:
            print(f"Error log: {e}")

# CONTINUACIÓN DE LA PARTE 1
# Agregar estos métodos a la clase EmailSenderGUI

    def actualizar_datos(self):
        """Actualizar datos"""
        self.log_mensaje("🔄 Actualizando datos...")
        
        if not self.excel_mgr:
            self.log_mensaje("❌ ExcelManager no disponible")
            messagebox.showerror("Error", "ExcelManager no inicializado")
            return
        
        self.btn_actualizar.config(state='disabled', text='🔄 Cargando...')
        
        try:
            # Estado archivos
            self.log_mensaje("📊 Leyendo Excel...")
            resumen = self.excel_mgr.obtener_resumen()
            
            self.text_archivos.config(state='normal')
            self.text_archivos.delete(1.0, tk.END)
            self.text_archivos.insert(1.0, resumen)
            self.text_archivos.config(state='disabled')
            
            # Campaña
            self.log_mensaje("🎯 Cargando campaña...")
            campanas = self.excel_mgr.cargar_campanas()
            
            if 'error' in campanas:
                self.label_campana_nombre.config(text="❌ Error campañas", fg=self.colors['red'])
                self.label_campana_asunto.config(text=f"Error: {campanas['error']}", fg=self.colors['red'])
                self.label_campana_info.config(text="Revisa CAMPAÑAS.xlsx", fg=self.colors['orange'])
            elif campanas['activa']:
                campana = campanas['activa']
                self.label_campana_nombre.config(text=f"📋 {campana['nombre']}", fg=self.colors['cyan'])
                self.label_campana_asunto.config(text=f"📧 {campana['asunto']}", fg=self.colors['foreground'])
                self.label_campana_info.config(text=f"📝 {len(campana['contenido'])} chars | ID: {campana['id']}", fg=self.colors['green'])
            else:
                self.label_campana_nombre.config(text="⚠️ Sin campaña activa", fg=self.colors['yellow'])
                self.label_campana_asunto.config(text="Marca 'SÍ' en alguna", fg=self.colors['comment'])
                self.label_campana_info.config(text=f"Total: {campanas['total']}", fg=self.colors['orange'])
            
            # Adjuntos
            if self.file_mgr:
                self.log_mensaje("📎 Adjuntos...")
                resumen_adj = self.file_mgr.obtener_resumen()
                
                self.text_adjuntos.config(state='normal')
                self.text_adjuntos.delete(1.0, tk.END)
                self.text_adjuntos.insert(1.0, resumen_adj)
                self.text_adjuntos.config(state='disabled')
            
            # Vista previa
            self.actualizar_vista_previa()
            
            self.btn_actualizar.config(state='normal', text='🔄 Actualizar Datos')
            self.label_estado.config(text="✅ Datos actualizados", fg=self.colors['green'])
            self.log_mensaje("✅ Actualización completa")
            
        except Exception as e:
            self.btn_actualizar.config(state='normal', text='🔄 Actualizar Datos')
            self.label_estado.config(text="❌ Error", fg=self.colors['red'])
            self.log_mensaje(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Error:\n{str(e)}")
    
    def actualizar_vista_previa(self):
        """Vista previa"""
        try:
            if not all([self.excel_mgr, self.email_processor]):
                self.text_preview.delete(1.0, tk.END)
                self.text_preview.insert(1.0, "❌ Managers no disponibles")
                return
            
            campanas = self.excel_mgr.cargar_campanas()
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            self.text_preview.delete(1.0, tk.END)
            
            if 'error' in campanas or 'error' in clientes or 'error' in config:
                self.text_preview.insert(1.0, "❌ Error en Excel")
                return
            
            if not campanas['activa']:
                self.text_preview.insert(1.0, "⚠️ Sin campaña activa")
                return
            
            if not clientes['clientes']:
                self.text_preview.insert(1.0, "⚠️ Sin clientes")
                return
            
            # Procesar primer correo
            primer_cliente = clientes['clientes'][0]
            correo = self.email_processor.procesar_lista_clientes([primer_cliente], campanas['activa'], config['config'])
            
            if correo:
                c = correo[0]
                preview = f"📧 VISTA PREVIA:\n"
                preview += "="*40 + "\n\n"
                preview += f"📮 Para: {c['email']}\n"
                preview += f"👤 Nombre: {c['nombre']}\n"
                preview += f"🏢 Empresa: {c['empresa']}\n"
                preview += f"📋 Asunto: {c['asunto']}\n\n"
                preview += "📝 CONTENIDO:\n"
                preview += "-"*25 + "\n"
                preview += c['contenido']
                
                self.text_preview.insert(1.0, preview)
            else:
                self.text_preview.insert(1.0, "❌ No procesado")
                
        except Exception as e:
            self.text_preview.delete(1.0, tk.END)
            self.text_preview.insert(1.0, f"❌ Error: {str(e)}")
    
    def vista_previa_completa(self):
        """Vista previa completa"""
        try:
            if not all([self.excel_mgr, self.email_processor]):
                messagebox.showerror("Error", "Managers no disponibles")
                return
            
            campanas = self.excel_mgr.cargar_campanas()
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            if 'error' in campanas or 'error' in clientes or 'error' in config or not campanas['activa']:
                messagebox.showerror("Error", "Datos incompletos")
                return
            
            correos = self.email_processor.procesar_lista_clientes(clientes['clientes'], campanas['activa'], config['config'])
            
            if not correos:
                messagebox.showwarning("Sin correos", "No hay correos válidos")
                return
            
            # Ventana
            ventana = tk.Toplevel(self.root)
            ventana.title(f"Vista Previa - {len(correos)} correos")
            ventana.geometry("900x600")
            ventana.configure(bg=self.colors['bg'])
            
            frame = tk.Frame(ventana, bg=self.colors['bg'], padx=15, pady=15)
            frame.pack(fill=tk.BOTH, expand=True)
            
            tk.Label(frame, text=f"📧 {len(correos)} correos listos", 
                    font=('Arial', 14, 'bold'), bg=self.colors['bg'], fg=self.colors['purple']).pack(pady=(0,15))
            
            text_widget = scrolledtext.ScrolledText(frame, font=('Consolas', 9), 
                                                   bg=self.colors['current_line'], fg=self.colors['foreground'])
            text_widget.pack(fill=tk.BOTH, expand=True)
            
            contenido = f"VISTA PREVIA COMPLETA\nTotal: {len(correos)} correos\n" + "="*50 + "\n\n"
            
            for i, correo in enumerate(correos[:5]):
                contenido += f"📩 CORREO #{correo['indice']}:\n"
                contenido += f"   Para: {correo['email']}\n"
                contenido += f"   Nombre: {correo['nombre']}\n"
                contenido += f"   Asunto: {correo['asunto']}\n"
                contenido += f"   Contenido: {correo['contenido'][:100]}...\n\n"
            
            if len(correos) > 5:
                contenido += f"... y {len(correos) - 5} correos más\n"
            
            text_widget.insert(1.0, contenido)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")

# CONTINUACIÓN DE LA PARTE 2
# Agregar estos métodos a la clase EmailSenderGUI

    def enviar_correos(self):
        """Enviar correos REALES"""
        self.log_mensaje("🚀 Iniciando envío...")
        
        # Verificar managers
        if not all([self.excel_mgr, self.email_processor, self.email_sender]):
            mensaje_error = "❌ Componentes no disponibles:\n"
            if not self.excel_mgr:
                mensaje_error += "• ExcelManager\n"
            if not self.email_processor:
                mensaje_error += "• EmailProcessor\n"
            if not self.email_sender:
                mensaje_error += "• EmailSender\n"
            
            self.log_mensaje("❌ Envío cancelado")
            messagebox.showerror("Error", mensaje_error)
            return
        
        # Verificar Outlook
        self.log_mensaje("🔄 Verificando Outlook...")
        conexion = self.email_sender.conectar_outlook()
        if not conexion['exitoso']:
            error_msg = f"❌ Error Outlook:\n{conexion['mensaje']}"
            if 'sugerencia' in conexion:
                error_msg += f"\n\nSugerencia:\n{conexion['sugerencia']}"
            
            self.log_mensaje("❌ Error Outlook")
            messagebox.showerror("Error Outlook", error_msg)
            return
        
        self.log_mensaje("✅ Outlook conectado")
        
        try:
            # Cargar datos
            self.log_mensaje("📊 Cargando datos...")
            campanas = self.excel_mgr.cargar_campanas()
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            if 'error' in campanas or 'error' in clientes or 'error' in config:
                self.log_mensaje("❌ Error en Excel")
                messagebox.showerror("Error", "Problemas con Excel")
                return
            
            if not campanas['activa']:
                self.log_mensaje("❌ Sin campaña activa")
                messagebox.showerror("Error", "Sin campaña activa")
                return
            
            # Procesar correos
            self.log_mensaje("📧 Procesando correos...")
            correos = self.email_processor.procesar_lista_clientes(clientes['clientes'], campanas['activa'], config['config'])
            
            if not correos:
                self.log_mensaje("❌ Sin correos válidos")
                messagebox.showerror("Error", "Sin correos válidos")
                return
            
            # Adjuntos
            adjuntos = []
            if self.file_mgr:
                adjuntos = self.file_mgr.obtener_archivos_validos()
                self.log_mensaje(f"📎 {len(adjuntos)} adjuntos")
            
            # Confirmar
            self.log_mensaje(f"✅ {len(correos)} correos listos")
            respuesta = messagebox.askyesno("🚀 Confirmar Envío REAL", 
                                          f"¿ENVIAR {len(correos)} correos REALES?\n\n"
                                          f"Campaña: {campanas['activa']['nombre']}\n"
                                          f"Adjuntos: {len(adjuntos)} archivos\n\n"
                                          f"⚠️ CORREOS REALES desde Outlook\n"
                                          f"⚠️ NO se puede deshacer")
            
            if respuesta:
                self.log_mensaje("🚀 Confirmado - iniciando...")
                self.iniciar_envio_real(correos, adjuntos)
            else:
                self.log_mensaje("❌ Cancelado por usuario")
                
        except Exception as e:
            self.log_mensaje(f"❌ Error preparando: {str(e)}")
            messagebox.showerror("Error", f"Error:\n{str(e)}")
    
    def iniciar_envio_real(self, correos, adjuntos):
        """Envío real en hilo"""
        self.log_mensaje(f"🚀 Enviando {len(correos)} correos...")
        
        # Cambiar estado
        self.enviando = True
        self.btn_enviar.config(state='disabled')
        self.btn_detener.config(state='normal')
        self.btn_actualizar.config(state='disabled')
        self.btn_preview.config(state='disabled')
        
        def proceso_envio():
            try:
                def callback_progreso(progreso, mensaje):
                    self.root.after(0, lambda: self.actualizar_progreso_seguro(progreso, mensaje))
                
                def detener_callback():
                    return not self.enviando
                
                self.log_mensaje("📤 Ejecutando envío...")
                resultados = self.email_sender.envio_por_lotes(
                    correos, 
                    adjuntos, 
                    callback_progreso=callback_progreso,
                    detener_callback=detener_callback
                )
                
                # Resultados en hilo principal
                self.root.after(0, lambda: self.procesar_resultados(resultados))
                
            except Exception as e:
                error_msg = f"Error durante envío: {str(e)}"
                self.root.after(0, lambda: self.log_mensaje(f"❌ {error_msg}"))
                self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                self.root.after(0, self.finalizar_envio)
        
        # Hilo separado
        threading.Thread(target=proceso_envio, daemon=True).start()
    
    def actualizar_progreso_seguro(self, progreso, mensaje):
        """Actualizar progreso SIN ERRORES"""
        try:
            # TTK ProgressBar si está disponible
            if hasattr(self, 'progress_bar') and self.progress_bar:
                self.progress_bar['value'] = progreso
            
            # Canvas fallback si TTK falló
            elif hasattr(self, 'progress_canvas'):
                width = self.progress_canvas.winfo_width()
                if width > 10:
                    self.progress_canvas.delete("all")
                    # Barra de fondo
                    self.progress_canvas.create_rectangle(0, 0, width, 20, fill=self.colors['current_line'], outline="")
                    # Barra de progreso
                    prog_width = int((progreso / 100) * width)
                    self.progress_canvas.create_rectangle(0, 0, prog_width, 20, fill=self.colors['purple'], outline="")
            
            self.label_estado.config(text=mensaje, fg=self.colors['cyan'])
            self.log_mensaje(f"📊 {progreso:.1f}% - {mensaje}")
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"Error progreso: {e}")
    
    def procesar_resultados(self, resultados):
        """Procesar resultados"""
        self.finalizar_envio()
        
        if 'error' in resultados:
            self.log_mensaje(f"❌ Error: {resultados['error']}")
            messagebox.showerror("Error", f"Error:\n{resultados['error']}")
            return
        
        # Estadísticas
        exitosos = len(resultados.get('exitosos', []))
        fallidos = len(resultados.get('fallidos', []))
        total = exitosos + fallidos
        
        porcentaje = (exitosos / total * 100) if total > 0 else 0
        
        # Log detallado
        self.log_mensaje("="*50)
        self.log_mensaje("📊 RESUMEN FINAL")
        self.log_mensaje("="*50)
        self.log_mensaje(f"✅ Exitosos: {exitosos}")
        self.log_mensaje(f"❌ Fallidos: {fallidos}")
        self.log_mensaje(f"📊 Total: {total}")
        self.log_mensaje(f"📈 Éxito: {porcentaje:.1f}%")
        
        if 'duracion' in resultados:
            self.log_mensaje(f"⏱️ Duración: {resultados['duracion']}")
        
        # Errores (primeros 3)
        if fallidos > 0:
            self.log_mensaje("\n❌ ERRORES:")
            for i, fallo in enumerate(resultados.get('fallidos', [])[:3]):
                email = fallo.get('email', 'desconocido')
                error = fallo.get('error', 'error desconocido')
                self.log_mensaje(f"   {i+1}. {email}: {error}")
            
            if len(resultados.get('fallidos', [])) > 3:
                restantes = len(resultados['fallidos']) - 3
                self.log_mensaje(f"   ... y {restantes} más")
        
        self.log_mensaje("="*50)
        
        # Mensaje usuario
        if porcentaje >= 90:
            icono, titulo, color = "🎉", "¡Éxito!", self.colors['green']
        elif porcentaje >= 70:
            icono, titulo, color = "✅", "Completado", self.colors['yellow']
        else:
            icono, titulo, color = "⚠️", "Con Problemas", self.colors['red']
        
        mensaje_final = f"{icono} {titulo}\n\n"
        mensaje_final += f"📊 ESTADÍSTICAS:\n"
        mensaje_final += f"• Total: {total}\n"
        mensaje_final += f"• Exitosos: {exitosos}\n"
        mensaje_final += f"• Fallidos: {fallidos}\n"
        mensaje_final += f"• Éxito: {porcentaje:.1f}%\n\n"
        mensaje_final += f"📋 Ver detalles en el log."
        
        if fallidos > 0:
            mensaje_final += f"\n\n⚠️ {fallidos} errores - revisar log."
        
        self.label_estado.config(text=f"{icono} Completado ({porcentaje:.0f}%)", fg=color)
        messagebox.showinfo(titulo, mensaje_final)
    
    def finalizar_envio(self):
        """Finalizar y restaurar"""
        self.enviando = False
        self.btn_enviar.config(state='normal')
        self.btn_detener.config(state='disabled')
        self.btn_actualizar.config(state='normal')
        self.btn_preview.config(state='normal')
        
        # Progreso al 100%
        try:
            if hasattr(self, 'progress_bar') and self.progress_bar:
                self.progress_bar['value'] = 100
            elif hasattr(self, 'progress_canvas'):
                width = self.progress_canvas.winfo_width()
                if width > 10:
                    self.progress_canvas.delete("all")
                    self.progress_canvas.create_rectangle(0, 0, width, 20, fill=self.colors['green'], outline="")
        except:
            pass
        
        self.log_mensaje("🏁 Proceso finalizado")
    
    def detener_envio(self):
        """Detener envío"""
        if self.enviando:
            respuesta = messagebox.askyesno("⏹️ Detener", 
                                          "¿Detener envío?\n\n"
                                          "Los enviados no se recuperan.")
            if respuesta:
                self.enviando = False
                self.log_mensaje("⏹️ DETENIDO por usuario")
                self.finalizar_envio()
                self.label_estado.config(text="⏹️ Detenido", fg=self.colors['red'])
    
    def on_closing(self):
        """Cerrar ventana"""
        if self.enviando:
            respuesta = messagebox.askyesno("⚠️ Envío en Progreso", 
                                          "¿Cerrar?\n\nSe detendrá el envío.")
            if not respuesta:
                return
            
            self.enviando = False
            self.log_mensaje("🔄 Cerrando - detenido")
        
        self.log_mensaje("👋 Cerrado")
        self.root.destroy()
    
    def ejecutar(self):
        """Ejecutar aplicación"""
        print("🚀 Iniciando...")
        
        # Crear interfaz
        self.crear_interfaz()
        
        # Log inicial
        self.log_mensaje("📧 Email Sender Pro iniciado")
        self.log_mensaje("💡 Haz clic en 'Actualizar Datos'")
        
        # Verificar
        if not all([self.excel_mgr, self.file_mgr, self.email_processor, self.email_sender]):
            self.log_mensaje("⚠️ Algunos componentes fallaron")
        
        # Configurar cierre
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Listo
        self.root.after(1000, lambda: self.log_mensaje("✅ Sistema listo"))
        
        # Iniciar
        print("✅ Loop principal")
        self.root.mainloop()

# FUNCIÓN PRINCIPAL
if __name__ == "__main__":
    print("="*60)
    print("📧 EMAIL SENDER PRO - VERSIÓN FUNCIONAL")
    print("🔥 SIN ERRORES TTK NI IMPORTACIÓN")
    print("="*60)
    
    try:
        print("🔧 Verificando...")
        print(f"📁 Directorio: {os.getcwd()}")
        print(f"📁 Script: {os.path.dirname(os.path.abspath(__file__))}")
        
        app = EmailSenderGUI()
        app.ejecutar()
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("\n🔧 SOLUCIONES:")
        print("1. Ejecuta desde directorio correcto")
        print("2. Verifica archivos en 'src/':")
        print("   • excel_manager.py")
        print("   • file_manager.py") 
        print("   • email_processor.py")
        print("   • email_sender.py")
        print("3. Comando: python src/gui.py")
        input("\nEnter para cerrar...")
        
    except Exception as e:
        print(f"❌ Error crítico: {e}")
        print(f"❌ Tipo: {type(e).__name__}")
        import traceback
        print("❌ Traceback:")
        traceback.print_exc()
        input("\nEnter para cerrar...")
    
    finally:
        print("🔚 Terminado")

# FIN DEL ARCHIVO COMPLETO