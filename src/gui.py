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

class EmailSenderInteligenteGUI:
    """GUI Email Sender Pro - CON ENVÍO INTELIGENTE ANTI-SPAM - PARTE 1"""
    
    def __init__(self):
        print("🔧 Inicializando GUI Inteligente...")
        
        # Crear ventana
        self.root = tk.Tk()
        self.root.title("📧 Email Sender Pro - INTELIGENTE ANTI-SPAM")
        
        # Variables
        self.enviando = False
        self.correos_procesados = []
        self.estrategia_actual = None
        
        # Configurar
        self._configurar_ventana()
        self._definir_colores()
        self._aplicar_tema_seguro()
        self._inicializar_managers_seguro()
        
        print("✅ GUI Inteligente inicializada")
    
    def _configurar_ventana(self):
        """Configurar ventana"""
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            width = int(screen_width * 0.9)
            height = int(screen_height * 0.9)
            
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            self.root.geometry(f"{width}x{height}+{x}+{y}")
            self.root.minsize(1300, 900)
            self.root.resizable(True, True)
            
            print(f"📐 Ventana: {width}x{height}")
        except Exception as e:
            print(f"⚠️ Error ventana: {e}")
            self.root.geometry("1300x900")
    
    def _definir_colores(self):
        """Colores profesionales estilo Apple con morado como accent"""
        self.colors = {
            # Backgrounds principales
            'bg': '#1a1a1a',                    # Fondo principal oscuro elegante
            'bg_secondary': '#2d2d2d',          # Fondo secundario
            'bg_card': '#353535',               # Fondo de tarjetas/panels
            'bg_input': '#404040',              # Fondo de inputs
            
            # Texto
            'text_primary': '#ffffff',          # Texto principal blanco
            'text_secondary': '#b0b0b0',        # Texto secundario gris claro
            'text_tertiary': '#808080',         # Texto terciario gris medio
            
            # Purple theme (Email Sender Pro)
            'purple_primary': '#8b5cf6',        # Morado principal
            'purple_light': '#a78bfa',          # Morado claro
            'purple_dark': '#7c3aed',           # morado oscuro
            'purple_bg': '#2d1b69',             # Morado de fondo
            
            # Status colors
            'success': '#10b981',               # Verde éxito
            'warning': '#f59e0b',               # Amarillo advertencia
            'error': '#ef4444',                 # Rojo error
            'info': '#3b82f6',                  # Azul información
            
            # Colores legacy para compatibilidad
            'current_line': '#404040',
            'selection': '#404040',
            'foreground': '#ffffff',
            'comment': '#808080',
            'cyan': '#06b6d4',
            'green': '#10b981',
            'orange': '#f59e0b',
            'pink': '#ec4899',
            'purple': '#8b5cf6',
            'red': '#ef4444',
            'yellow': '#eab308'
        }
    
    def _aplicar_tema_seguro(self):
        """Aplicar tema profesional estilo Apple"""
        try:
            self.root.configure(bg=self.colors['bg'])
            
            self.style = ttk.Style()
            
            try:
                self.style.theme_use('clam')
            except:
                print("⚠️ Usando tema default")
            
            try:
                # Configurar estilos profesionales
                
                # Botones principales con estilo Apple
                self.style.configure('Professional.TButton',
                                   background=self.colors['purple_primary'],
                                   foreground='white',
                                   borderwidth=0,
                                   focuscolor='none',
                                   padding=(20, 12),
                                   font=('SF Pro Display', 11, 'normal'))
                
                self.style.map('Professional.TButton',
                              background=[('active', self.colors['purple_light']),
                                        ('pressed', self.colors['purple_dark'])])
                
                # Botón de acción principal (ENVÍO INTELIGENTE)
                self.style.configure('Primary.TButton',
                                   background=self.colors['purple_primary'],
                                   foreground='white',
                                   borderwidth=0,
                                   focuscolor='none',
                                   padding=(25, 15),
                                   font=('SF Pro Display', 12, 'bold'))
                
                self.style.map('Primary.TButton',
                              background=[('active', self.colors['purple_light']),
                                        ('pressed', self.colors['purple_dark'])])
                
                # Botón secundario
                self.style.configure('Secondary.TButton',
                                   background=self.colors['bg_card'],
                                   foreground=self.colors['text_primary'],
                                   borderwidth=1,
                                   focuscolor='none',
                                   padding=(18, 10),
                                   font=('SF Pro Display', 10, 'normal'))
                
                self.style.map('Secondary.TButton',
                              background=[('active', self.colors['bg_input']),
                                        ('pressed', self.colors['bg_secondary'])])
                
                # Botón de peligro (DETENER)
                self.style.configure('Danger.TButton',
                                   background=self.colors['error'],
                                   foreground='white',
                                   borderwidth=0,
                                   focuscolor='none',
                                   padding=(20, 12),
                                   font=('SF Pro Display', 11, 'normal'))
                
                self.style.map('Danger.TButton',
                              background=[('active', '#dc2626'),
                                        ('pressed', '#b91c1c')])
                
                # Progress bar con morado
                self.style.configure('Purple.Horizontal.TProgressbar',
                                   background=self.colors['purple_primary'],
                                   troughcolor=self.colors['bg_card'],
                                   borderwidth=0,
                                   lightcolor=self.colors['purple_primary'],
                                   darkcolor=self.colors['purple_primary'])
                
                print("✅ Estilos profesionales aplicados")
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
        
        # EmailSender - CORREGIDO PARA USAR email_sender
        try:
            from email_sender import EmailSender
            self.email_sender = EmailSender()
            print("✅ EmailSender OK")
        except Exception as e:
            print(f"❌ EmailSender error: {e}")
            print("🔧 Creando EmailSender funcional...")
            self.email_sender = self._crear_email_sender_funcional()
        
        print("✅ Managers listos")
    
    def _crear_email_sender_funcional(self):
        """EmailSender funcional integrado con lógica inteligente"""
        class EmailSenderFuncional:
            def __init__(self):
                self.conectado = False
                self.MAX_CORREOS_DIARIOS = 400
                self.HORAS_TRABAJO = 8
                self.LIMITE_RAPIDO = 25
                print("📧 EmailSender funcional con lógica inteligente creado")
            
            def conectar_outlook(self):
                try:
                    import win32com.client
                    import pythoncom
                    
                    # ⭐ Inicializar COM
                    pythoncom.CoInitialize()
                    
                    try:
                        # Intentar conectar a instancia existente
                        outlook = win32com.client.GetActiveObject("Outlook.Application")
                        print("✅ Conectado a instancia existente")
                    except:
                        # Crear nueva instancia
                        outlook = win32com.client.Dispatch("Outlook.Application")
                        print("✅ Nueva instancia creada")
                    
                    # Verificar funcionamiento
                    namespace = outlook.GetNamespace("MAPI")
                    inbox = namespace.GetDefaultFolder(6)
                    
                    # Verificar cuentas
                    accounts = namespace.Accounts
                    if accounts.Count == 0:
                        raise Exception("No hay cuentas configuradas en Outlook")
                    
                    cuenta_principal = accounts.Item(1)
                    email_cuenta = getattr(cuenta_principal, 'SmtpAddress', cuenta_principal.DisplayName)
                    
                    self.conectado = True
                    
                    return {
                        'exitoso': True,
                        'mensaje': f'Conectado a Outlook correctamente',
                        'cuenta': email_cuenta,
                        'total_cuentas': accounts.Count
                    }
                    
                except Exception as e:
                    return {
                        'exitoso': False,
                        'mensaje': f'Error Outlook: {str(e)}',
                        'sugerencia': 'Abre Outlook primero'
                    }
            
            def calcular_estrategia_envio(self, total_correos):
                """Calcular estrategia simple"""
                if total_correos <= 2:
                    return {
                        'modo': 'INMEDIATO',
                        'descripcion': 'Envío inmediato sin pausas',
                        'pausa_entre_correos': 5
                    }
                elif total_correos <= self.LIMITE_RAPIDO:
                    return {
                        'modo': 'RÁPIDO',
                        'descripcion': f'Envío rápido con pausas de 30s',
                        'pausa_entre_correos': 30
                    }
                else:
                    return {
                        'modo': 'DISTRIBUIDO',
                        'descripcion': f'Envío distribuido con pausas de 6 minutos',
                        'pausa_entre_correos': 360
                    }
            
            def enviar_correo(self, correo_data, adjuntos=None):
                try:
                    import win32com.client
                    import pythoncom
                    
                    # ⭐ CLAVE: Inicializar COM en cada envío
                    pythoncom.CoInitialize()
                    
                    try:
                        outlook = win32com.client.Dispatch("Outlook.Application")
                        
                        mail = outlook.CreateItem(0)
                        mail.To = correo_data['email']
                        mail.Subject = correo_data['asunto']
                        
                        # ⭐ CONFIGURAR CONTENIDO CON TEXTO NEGRO FORZADO
                        contenido = correo_data['contenido']
                        if '\n' in contenido and '<br>' not in contenido.lower():
                            contenido_html = contenido.replace('\n', '<br>')
                            mail.HTMLBody = f"""
                            <html>
                            <body style="font-family: Arial, sans-serif; font-size: 12pt; color: #000000; background-color: #ffffff;">
                            {contenido_html}
                            </body>
                            </html>
                            """
                        else:
                            if '<html>' in contenido.lower():
                                mail.HTMLBody = contenido
                            else:
                                mail.Body = contenido
                        
                        # ⭐ AGREGAR ADJUNTOS - DEBUGGING COMPLETO
                        adjuntos_agregados = 0
                        if adjuntos:
                            print(f"🔍 DEBUG: Procesando {len(adjuntos)} adjuntos")
                            for i, ruta_adjunto in enumerate(adjuntos):
                                print(f"🔍 DEBUG: Adjunto {i+1}: {ruta_adjunto}")
                                
                                # Convertir a ruta absoluta
                                ruta_absoluta = os.path.abspath(ruta_adjunto)
                                print(f"🔍 DEBUG: Ruta absoluta: {ruta_absoluta}")
                                
                                if os.path.exists(ruta_absoluta):
                                    try:
                                        print(f"📎 Agregando adjunto: {os.path.basename(ruta_absoluta)}")
                                        mail.Attachments.Add(ruta_absoluta)
                                        adjuntos_agregados += 1
                                        print(f"✅ Adjunto agregado exitosamente: {os.path.basename(ruta_absoluta)}")
                                    except Exception as attach_error:
                                        print(f"❌ Error adjuntando {ruta_absoluta}: {attach_error}")
                                        print(f"❌ Tipo error: {type(attach_error).__name__}")
                                else:
                                    print(f"❌ Archivo NO EXISTE: {ruta_absoluta}")
                            
                            print(f"📊 Total adjuntos agregados: {adjuntos_agregados}/{len(adjuntos)}")
                        else:
                            print("📎 No hay adjuntos para procesar")
                        
                        # ⭐ ENVIAR
                        print(f"📤 Enviando correo a {correo_data['email']}...")
                        mail.Send()
                        print(f"✅ Correo enviado exitosamente")
                        
                        return {
                            'exitoso': True,
                            'timestamp': time.strftime('%H:%M:%S'),
                            'email': correo_data['email'],
                            'nombre': correo_data.get('nombre', 'Sin nombre'),
                            'adjuntos_agregados': adjuntos_agregados
                        }
                    
                    finally:
                        # ⭐ LIMPIAR COM
                        try:
                            pythoncom.CoUninitialize()
                        except:
                            pass
                    
                except Exception as e:
                    print(f"❌ ERROR GENERAL en enviar_correo: {e}")
                    print(f"❌ Tipo error: {type(e).__name__}")
                    return {
                        'exitoso': False,
                        'error': str(e),
                        'email': correo_data.get('email', 'desconocido'),
                        'nombre': correo_data.get('nombre', 'Sin nombre')
                    }
            
            def envio_inteligente(self, correos, adjuntos, callback_progreso=None, detener_callback=None):
                """Envío inteligente adaptado con COM corregido"""
                import pythoncom
                
                # ⭐ INICIALIZAR COM en el hilo de envío
                pythoncom.CoInitialize()
                
                try:
                    estrategia = self.calcular_estrategia_envio(len(correos))
                    
                    resultados = {
                        'exitosos': [],
                        'fallidos': [],
                        'total_procesados': 0,
                        'estrategia': estrategia,
                        'inicio': time.time()
                    }
                    
                    pausa = estrategia['pausa_entre_correos']
                    
                    for i, correo in enumerate(correos):
                        if detener_callback and detener_callback():
                            break
                        
                        if callback_progreso:
                            progreso = (i / len(correos)) * 100
                            nombre = correo.get('nombre', 'Sin nombre')
                            callback_progreso(progreso, f"[{estrategia['modo']}] Enviando a {nombre} ({i+1}/{len(correos)})")
                        
                        resultado = self.enviar_correo(correo, adjuntos)
                        
                        if resultado['exitoso']:
                            resultados['exitosos'].append(resultado)
                        else:
                            resultados['fallidos'].append(resultado)
                        
                        resultados['total_procesados'] += 1
                        
                        # Pausa inteligente
                        if i < len(correos) - 1:
                            for segundo in range(pausa):
                                if detener_callback and detener_callback():
                                    break
                                time.sleep(1)
                                
                                if segundo % 30 == 0 and callback_progreso:
                                    tiempo_restante = pausa - segundo
                                    if tiempo_restante > 60:
                                        tiempo_texto = f"{tiempo_restante // 60}m {tiempo_restante % 60}s"
                                    else:
                                        tiempo_texto = f"{tiempo_restante}s"
                                    callback_progreso(progreso, f"Pausa {estrategia['modo'].lower()}: {tiempo_texto}")
                    
                    resultados['fin'] = time.time()
                    resultados['duracion'] = f"{resultados['fin'] - resultados['inicio']:.1f}s"
                    
                    return resultados
                
                finally:
                    # ⭐ LIMPIAR COM al final
                    try:
                        pythoncom.CoUninitialize()
                    except:
                        pass
        
        return EmailSenderFuncional()

# FIN DE LA PARTE 1 - Configuración, colores, estilos y EmailSender funcional

# PARTE 2 - INTERFAZ GRÁFICA Y PANELES

    def crear_interfaz(self):
        """Crear interfaz profesional estilo Apple"""
        print("🏗️ Creando interfaz profesional...")
        
        # Frame principal con padding estilo Apple
        self.main_frame = tk.Frame(self.root, bg=self.colors['bg'], padx=30, pady=25)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Grid con proporciones Apple-like
        self.main_frame.columnconfigure(0, weight=3, minsize=600)  # Contenido principal más ancho
        self.main_frame.columnconfigure(1, weight=2, minsize=400)  # Panel lateral
        
        # Filas con espaciado proporcional
        self.main_frame.rowconfigure(0, weight=0, minsize=80)   # Título
        self.main_frame.rowconfigure(1, weight=0, minsize=140)  # Estrategia
        self.main_frame.rowconfigure(2, weight=1, minsize=400)  # Contenido principal
        self.main_frame.rowconfigure(3, weight=0, minsize=80)   # Botones
        self.main_frame.rowconfigure(4, weight=0, minsize=60)   # Progreso
        
        # Crear secciones con diseño Apple
        self.crear_titulo_profesional()
        self.crear_estrategia_profesional()
        self.crear_contenido_profesional()
        self.crear_botones_profesionales()
        self.crear_progreso_profesional()
        
        print("✅ Interfaz profesional creada")
    
    def crear_titulo_profesional(self):
        """Título con diseño Apple"""
        titulo_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        titulo_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 25))
        
        # Título principal con tipografía Apple
        titulo_principal = tk.Label(titulo_frame, 
                                  text="📧 EMAIL SENDER PRO", 
                                  font=('SF Pro Display', 32, 'bold'), 
                                  bg=self.colors['bg'], 
                                  fg=self.colors['purple_primary'])
        titulo_principal.pack(pady=(0, 8))
        
        # Subtítulo elegante
        subtitulo = tk.Label(titulo_frame, 
                           text="🧠 Envío Inteligente Anti-Spam • Distribución Automática", 
                           font=('SF Pro Text', 14, 'normal'), 
                           bg=self.colors['bg'], 
                           fg=self.colors['text_secondary'])
        subtitulo.pack()
    
    def crear_estrategia_profesional(self):
        """Panel de estrategia con diseño Apple"""
        # Frame contenedor con esquinas redondeadas simuladas
        strategy_container = tk.Frame(self.main_frame, bg=self.colors['bg'])
        strategy_container.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 20))
        strategy_container.columnconfigure(0, weight=1)
        
        # Header del panel
        header_frame = tk.Frame(strategy_container, bg=self.colors['bg_card'], height=50)
        header_frame.pack(fill=tk.X, pady=(0, 2))
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(header_frame, 
                              text="🎯 Estrategia de Envío Inteligente",
                              font=('SF Pro Display', 16, 'bold'), 
                              bg=self.colors['bg_card'], 
                              fg=self.colors['purple_primary'])
        header_label.pack(pady=12)
        
        # Contenido del panel
        content_frame = tk.Frame(strategy_container, bg=self.colors['bg_card'])
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        
        self.text_estrategia = scrolledtext.ScrolledText(content_frame, 
                                                        height=6, 
                                                        font=('SF Mono', 11),
                                                        bg=self.colors['bg_input'], 
                                                        fg=self.colors['text_primary'], 
                                                        relief='flat',
                                                        borderwidth=0,
                                                        insertbackground=self.colors['purple_primary'])
        self.text_estrategia.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        # Mensaje inicial estilizado
        mensaje_inicial = """🎯 ESTRATEGIA INTELIGENTE ANTI-SPAM

✅ 1-2 correos → Envío INMEDIATO (sin pausas)
⚡ 3-25 correos → Envío RÁPIDO (pausas de 30 segundos)  
📦 26+ correos → Envío DISTRIBUIDO (lotes con pausas de 6 minutos)

💡 Haz clic en 'Actualizar Datos' para calcular la estrategia específica"""
        
        self.text_estrategia.insert(1.0, mensaje_inicial)
        self.text_estrategia.config(state='disabled')
    
    def crear_contenido_profesional(self):
        """Área de contenido con diseño Apple"""
        # Panel principal izquierdo
        left_panel = tk.Frame(self.main_frame, bg=self.colors['bg'])
        left_panel.grid(row=2, column=0, sticky='nsew', padx=(0, 15))
        left_panel.rowconfigure(0, weight=0, minsize=160)  # Estado archivos
        left_panel.rowconfigure(1, weight=0, minsize=120)  # Campaña activa  
        left_panel.rowconfigure(2, weight=1, minsize=300)  # Vista previa
        left_panel.columnconfigure(0, weight=1)
        
        self.crear_panel_archivos(left_panel)
        self.crear_panel_campana(left_panel)
        self.crear_panel_vista_previa(left_panel)
        
        # Panel lateral derecho
        right_panel = tk.Frame(self.main_frame, bg=self.colors['bg'])
        right_panel.grid(row=2, column=1, sticky='nsew')
        right_panel.rowconfigure(0, weight=0, minsize=200)  # Adjuntos
        right_panel.rowconfigure(1, weight=1, minsize=350)  # Log
        right_panel.columnconfigure(0, weight=1)
        
        self.crear_panel_adjuntos(right_panel)
        self.crear_panel_log(right_panel)
    
    def crear_panel_archivos(self, parent):
        """Panel de estado de archivos estilo Apple"""
        # Frame principal con fondo tipo card
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        card_frame.columnconfigure(0, weight=1)
        
        # Header del panel
        header = tk.Label(card_frame, 
                         text="📊 Estado de Archivos Excel",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido scrollable
        self.text_archivos = scrolledtext.ScrolledText(card_frame, 
                                                      height=6, 
                                                      font=('SF Mono', 10),
                                                      bg=self.colors['bg_input'], 
                                                      fg=self.colors['text_primary'], 
                                                      relief='flat',
                                                      borderwidth=0)
        self.text_archivos.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_panel_campana(self, parent):
        """Panel de campaña activa estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=1, column=0, sticky='ew', pady=(0, 15))
        card_frame.columnconfigure(0, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="🎯 Campaña Activa",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Información de campaña
        info_frame = tk.Frame(card_frame, bg=self.colors['bg_card'])
        info_frame.pack(fill=tk.BOTH, padx=20, pady=(0, 15))
        info_frame.columnconfigure(0, weight=1)
        
        self.label_campana_nombre = tk.Label(info_frame, 
                                           text="📋 Campaña: Cargando...",
                                           font=('SF Pro Text', 12, 'bold'), 
                                           bg=self.colors['bg_card'], 
                                           fg=self.colors['text_primary'], 
                                           anchor='w')
        self.label_campana_nombre.grid(row=0, column=0, sticky='ew', pady=(0, 5))
        
        self.label_campana_asunto = tk.Label(info_frame, 
                                           text="📧 Asunto: Cargando...",
                                           font=('SF Pro Text', 11, 'normal'), 
                                           bg=self.colors['bg_card'], 
                                           fg=self.colors['text_secondary'], 
                                           anchor='w', 
                                           wraplength=500)
        self.label_campana_asunto.grid(row=1, column=0, sticky='ew', pady=(0, 5))
        
        self.label_campana_info = tk.Label(info_frame, 
                                         text="📝 Info: Cargando...",
                                         font=('SF Pro Text', 10, 'normal'), 
                                         bg=self.colors['bg_card'], 
                                         fg=self.colors['text_tertiary'], 
                                         anchor='w', 
                                         wraplength=500)
        self.label_campana_info.grid(row=2, column=0, sticky='ew')
    
    def crear_panel_vista_previa(self, parent):
        """Panel de vista previa estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=2, column=0, sticky='nsew')
        card_frame.columnconfigure(0, weight=1)
        card_frame.rowconfigure(1, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="📧 Vista Previa del Primer Correo",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido
        self.text_preview = scrolledtext.ScrolledText(card_frame, 
                                                     font=('SF Mono', 10),
                                                     bg=self.colors['bg_input'], 
                                                     fg=self.colors['text_primary'], 
                                                     relief='flat',
                                                     borderwidth=0)
        self.text_preview.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_panel_adjuntos(self, parent):
        """Panel de adjuntos estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        card_frame.columnconfigure(0, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="📎 Archivos Adjuntos",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido
        self.text_adjuntos = tk.Text(card_frame, 
                                    height=8, 
                                    font=('SF Mono', 9),
                                    bg=self.colors['bg_input'], 
                                    fg=self.colors['warning'], 
                                    relief='flat', 
                                    borderwidth=0,
                                    wrap=tk.WORD)
        self.text_adjuntos.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_panel_log(self, parent):
        """Panel de log estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=1, column=0, sticky='nsew')
        card_frame.columnconfigure(0, weight=1)
        card_frame.rowconfigure(1, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="📋 Log en Tiempo Real",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido
        self.text_log = scrolledtext.ScrolledText(card_frame, 
                                                 font=('SF Mono', 9),
                                                 bg=self.colors['bg_input'], 
                                                 fg=self.colors['info'], 
                                                 relief='flat',
                                                 borderwidth=0)
        self.text_log.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_botones_profesionales(self):
        """Botones con diseño Apple profesional"""
        buttons_container = tk.Frame(self.main_frame, bg=self.colors['bg'])
        buttons_container.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(25, 0))
        
        # Grid para botones con espaciado proporcional
        for i in range(5):
            buttons_container.columnconfigure(i, weight=1, minsize=180)
        
        # Botones con estilos específicos
        self.btn_actualizar = ttk.Button(buttons_container, 
                                        text="🔄 Actualizar Datos", 
                                        command=self.actualizar_datos,
                                        style='Secondary.TButton')
        self.btn_actualizar.grid(row=0, column=0, padx=8, pady=15, sticky='ew')
        
        self.btn_estrategia = ttk.Button(buttons_container, 
                                        text="🎯 Ver Estrategia", 
                                        command=self.mostrar_estrategia,
                                        style='Secondary.TButton')
        self.btn_estrategia.grid(row=0, column=1, padx=8, pady=15, sticky='ew')
        
        self.btn_preview = ttk.Button(buttons_container, 
                                     text="👁️ Vista Previa", 
                                     command=self.vista_previa_completa,
                                     style='Professional.TButton')
        self.btn_preview.grid(row=0, column=2, padx=8, pady=15, sticky='ew')
        
        self.btn_enviar = ttk.Button(buttons_container, 
                                    text="🚀 ENVÍO INTELIGENTE", 
                                    command=self.enviar_correos_inteligente,
                                    style='Primary.TButton')
        self.btn_enviar.grid(row=0, column=3, padx=8, pady=15, sticky='ew')
        
        self.btn_detener = ttk.Button(buttons_container, 
                                     text="⏹️ DETENER", 
                                     state="disabled", 
                                     command=self.detener_envio,
                                     style='Danger.TButton')
        self.btn_detener.grid(row=0, column=4, padx=8, pady=15, sticky='ew')
    
    def crear_progreso_profesional(self):
        """Barra de progreso estilo Apple"""
        progress_container = tk.Frame(self.main_frame, bg=self.colors['bg'])
        progress_container.grid(row=4, column=0, columnspan=2, sticky='ew', pady=(20, 0))
        
        # Label de estado con tipografía Apple
        self.label_estado = tk.Label(progress_container, 
                                   text="✅ Sistema listo - Haz clic en 'Actualizar Datos'",
                                   font=('SF Pro Text', 13, 'normal'), 
                                   bg=self.colors['bg'], 
                                   fg=self.colors['success'])
        self.label_estado.pack(pady=(0, 15))
        
        # Progress bar container con fondo
        progress_bg = tk.Frame(progress_container, bg=self.colors['bg_card'], height=8)
        progress_bg.pack(fill=tk.X, padx=60, pady=(0, 5))
        
        try:
            # Progress bar con estilo morado
            self.progress_bar = ttk.Progressbar(progress_bg, 
                                              mode='determinate', 
                                              length=400,
                                              style='Purple.Horizontal.TProgressbar')
            self.progress_bar.pack(fill=tk.X, padx=2, pady=2)
            print("✅ ProgressBar profesional OK")
        except Exception as e:
            print(f"⚠️ Error ProgressBar: {e}")
            # Fallback a canvas con diseño Apple
            self.progress_canvas = tk.Canvas(progress_bg, 
                                           height=8, 
                                           bg=self.colors['bg_card'],
                                           highlightthickness=0)
            self.progress_canvas.pack(fill=tk.X, padx=2, pady=2)
            self.progress_bar = None
            print("✅ Canvas como ProgressBar profesional")
    
    def log_mensaje(self, mensaje):
        """Log"""
        try:
            timestamp = time.strftime('%H:%M:%S')
            self.text_log.insert(tk.END, f"[{timestamp}] {mensaje}\n")
            self.text_log.see(tk.END)
            self.root.update_idletasks()
        except Exception as e:
            print(f"Error log: {e}")

# FIN DE LA PARTE 2 - Interfaz gráfica completa
# FIN DE LA PARTE 2 - Interfaz gráfica completa

# PARTE 3 - LÓGICA DE FUNCIONES Y OPERACIONES

    def actualizar_datos(self):
        """Actualizar datos con estrategia"""
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
                self.label_campana_nombre.config(text="❌ Error campañas", fg=self.colors['error'])
                self.label_campana_asunto.config(text=f"Error: {campanas['error']}", fg=self.colors['error'])
                self.label_campana_info.config(text="Revisa CAMPAÑAS.xlsx", fg=self.colors['warning'])
            elif campanas['activa']:
                campana = campanas['activa']
                self.label_campana_nombre.config(text=f"📋 {campana['nombre']}", fg=self.colors['purple_primary'])
                self.label_campana_asunto.config(text=f"📧 {campana['asunto']}", fg=self.colors['text_primary'])
                self.label_campana_info.config(text=f"📝 {len(campana['contenido'])} chars | ID: {campana['id']}", fg=self.colors['success'])
            else:
                self.label_campana_nombre.config(text="⚠️ Sin campaña activa", fg=self.colors['warning'])
                self.label_campana_asunto.config(text="Marca 'SÍ' en alguna", fg=self.colors['text_tertiary'])
                self.label_campana_info.config(text=f"Total: {campanas['total']}", fg=self.colors['info'])
            
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
            
            # CALCULAR Y MOSTRAR ESTRATEGIA
            self.actualizar_estrategia()
            
            self.btn_actualizar.config(state='normal', text='🔄 Actualizar Datos')
            self.label_estado.config(text="✅ Datos actualizados - Estrategia calculada", fg=self.colors['success'])
            self.log_mensaje("✅ Actualización completa")
            
        except Exception as e:
            self.btn_actualizar.config(state='normal', text='🔄 Actualizar Datos')
            self.label_estado.config(text="❌ Error", fg=self.colors['error'])
            self.log_mensaje(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Error:\n{str(e)}")
    
    def actualizar_estrategia(self):
        """Actualizar estrategia de envío"""
        try:
            if not all([self.excel_mgr, self.email_processor]):
                return
            
            campanas = self.excel_mgr.cargar_campanas()
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            if 'error' in campanas or 'error' in clientes or 'error' in config:
                return
            
            if not campanas['activa'] or not clientes['clientes']:
                return
            
            # Procesar correos para obtener total
            correos = self.email_processor.procesar_lista_clientes(clientes['clientes'], campanas['activa'], config['config'])
            total_correos = len(correos)
            
            # Calcular estrategia
            if hasattr(self.email_sender, 'calcular_estrategia_envio'):
                self.estrategia_actual = self.email_sender.calcular_estrategia_envio(total_correos)
                
                # Fallback para sender simple
                estrategia_texto = f"🎯 ESTRATEGIA PARA {total_correos} CORREOS\n"
                estrategia_texto += "=" * 40 + "\n\n"
                estrategia_texto += f"🚀 Modo: {self.estrategia_actual.get('modo', 'AUTOMÁTICO')}\n"
                estrategia_texto += f"📝 {self.estrategia_actual.get('descripcion', 'Envío automático')}\n"
                estrategia_texto += f"⏱️ Pausa entre correos: {self.estrategia_actual.get('pausa_entre_correos', 30)}s\n"
            else:
                # Sender simple sin estrategia avanzada
                if total_correos <= 2:
                    modo = "INMEDIATO"
                    desc = "Sin pausas"
                elif total_correos <= 25:
                    modo = "RÁPIDO"
                    desc = "Pausas de 30 segundos"
                else:
                    modo = "DISTRIBUIDO"
                    desc = "Pausas de 6 minutos"
                
                estrategia_texto = f"🎯 ESTRATEGIA PARA {total_correos} CORREOS\n"
                estrategia_texto += "=" * 40 + "\n\n"
                estrategia_texto += f"🚀 Modo: {modo}\n"
                estrategia_texto += f"📝 {desc}\n"
                estrategia_texto += f"✅ Anti-spam activado\n"
            
            # Mostrar en panel
            self.text_estrategia.config(state='normal')
            self.text_estrategia.delete(1.0, tk.END)
            self.text_estrategia.insert(1.0, estrategia_texto)
            self.text_estrategia.config(state='disabled')
            
            self.log_mensaje(f"🎯 Estrategia calculada para {total_correos} correos")
            
        except Exception as e:
            self.log_mensaje(f"⚠️ Error calculando estrategia: {str(e)}")
    
    def mostrar_estrategia(self):
        """Mostrar estrategia en ventana separada"""
        if not self.estrategia_actual:
            messagebox.showinfo("Info", "Primero actualiza los datos para calcular la estrategia")
            return
        
        ventana = tk.Toplevel(self.root)
        ventana.title("🎯 Estrategia de Envío Detallada")
        ventana.geometry("800x600")
        ventana.configure(bg=self.colors['bg'])
        
        frame = tk.Frame(ventana, bg=self.colors['bg'], padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(frame, text="🎯 ESTRATEGIA DE ENVÍO INTELIGENTE", 
                font=('Arial', 16, 'bold'), bg=self.colors['bg'], fg=self.colors['purple_primary']).pack(pady=(0,20))
        
        text_widget = scrolledtext.ScrolledText(frame, font=('Consolas', 10), 
                                               bg=self.colors['bg_input'], fg=self.colors['text_primary'])
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        contenido = "Estrategia calculada automáticamente\n\nRevisa el panel principal para más detalles."
        if hasattr(self.estrategia_actual, 'modo'):
            contenido = f"Modo: {self.estrategia_actual.get('modo', 'AUTO')}\n"
            contenido += f"Descripción: {self.estrategia_actual.get('descripcion', 'Envío automático')}\n"
            contenido += f"Pausa entre correos: {self.estrategia_actual.get('pausa_entre_correos', 30)}s"
        
        text_widget.insert(1.0, contenido)
        text_widget.config(state='disabled')
    
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
                    font=('Arial', 14, 'bold'), bg=self.colors['bg'], fg=self.colors['purple_primary']).pack(pady=(0,15))
            
            text_widget = scrolledtext.ScrolledText(frame, font=('Consolas', 9), 
                                                   bg=self.colors['bg_input'], fg=self.colors['text_primary'])
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
    
    def enviar_correos_inteligente(self):
        """Enviar correos con LÓGICA INTELIGENTE"""
        self.log_mensaje("🚀 Iniciando envío inteligente...")
        
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
            
            # CALCULAR ESTRATEGIA
            total_correos = len(correos)
            if hasattr(self.email_sender, 'calcular_estrategia_envio'):
                estrategia = self.email_sender.calcular_estrategia_envio(total_correos)
                modo = estrategia.get('modo', 'AUTOMÁTICO')
                descripcion = estrategia.get('descripcion', 'Envío automático')
            else:
                # Fallback para sender simple
                if total_correos <= 2:
                    modo = "INMEDIATO"
                    descripcion = "Sin pausas"
                elif total_correos <= 25:
                    modo = "RÁPIDO" 
                    descripcion = "Pausas cortas"
                else:
                    modo = "DISTRIBUIDO"
                    descripcion = "Pausas largas"
            
            # Confirmar con estrategia
            self.log_mensaje(f"✅ {total_correos} correos listos - Modo: {modo}")
            respuesta = messagebox.askyesno("🚀 Confirmar Envío INTELIGENTE", 
                                          f"¿ENVIAR {total_correos} correos REALES?\n\n"
                                          f"🎯 ESTRATEGIA: {modo}\n"
                                          f"📝 {descripcion}\n"
                                          f"📧 Campaña: {campanas['activa']['nombre']}\n"
                                          f"📎 Adjuntos: {len(adjuntos)} archivos\n\n"
                                          f"🧠 ENVÍO INTELIGENTE ANTI-SPAM\n"
                                          f"⚠️ CORREOS REALES desde Outlook\n"
                                          f"⚠️ NO se puede deshacer")
            
            if respuesta:
                self.log_mensaje(f"🚀 Confirmado - iniciando envío {modo}...")
                self.iniciar_envio_inteligente(correos, adjuntos)
            else:
                self.log_mensaje("❌ Cancelado por usuario")
                
        except Exception as e:
            self.log_mensaje(f"❌ Error preparando: {str(e)}")
            messagebox.showerror("Error", f"Error:\n{str(e)}")
    
    def iniciar_envio_inteligente(self, correos, adjuntos):
        """Envío inteligente en hilo"""
        total = len(correos)
        self.log_mensaje(f"🧠 Envío inteligente: {total} correos")
        
        # Cambiar estado
        self.enviando = True
        self.btn_enviar.config(state='disabled')
        self.btn_detener.config(state='normal')
        self.btn_actualizar.config(state='disabled')
        self.btn_preview.config(state='disabled')
        self.btn_estrategia.config(state='disabled')
        
        def proceso_envio_inteligente():
            try:
                def callback_progreso(progreso, mensaje):
                    # Verificar si la ventana aún existe antes de usar after
                    try:
                        if self.root.winfo_exists():
                            self.root.after(0, lambda: self.actualizar_progreso_seguro(progreso, mensaje))
                    except tk.TclError:
                        # La ventana se cerró, no hacer nada
                        pass
                
                def detener_callback():
                    return not self.enviando
                
                self.log_mensaje("🧠 Ejecutando envío inteligente...")
                
                # Usar envío inteligente si está disponible
                if hasattr(self.email_sender, 'envio_inteligente'):
                    resultados = self.email_sender.envio_inteligente(
                        correos, 
                        adjuntos, 
                        callback_progreso=callback_progreso,
                        detener_callback=detener_callback
                    )
                else:
                    # Fallback al método original
                    resultados = self.email_sender.envio_por_lotes(
                        correos, 
                        adjuntos, 
                        callback_progreso=callback_progreso,
                        detener_callback=detener_callback
                    )
                
                # Resultados en hilo principal - verificar si ventana existe
                try:
                    if self.root.winfo_exists():
                        self.root.after(0, lambda: self.procesar_resultados_inteligente(resultados))
                except tk.TclError:
                    # La ventana se cerró, no hacer nada
                    pass
                
            except Exception as e:
                error_msg = f"Error durante envío: {str(e)}"
                try:
                    if self.root.winfo_exists():
                        self.root.after(0, lambda: self.log_mensaje(f"❌ {error_msg}"))
                        self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                        self.root.after(0, self.finalizar_envio)
                except tk.TclError:
                    # La ventana se cerró, solo imprimir el error
                    print(f"❌ {error_msg}")
        
        # Hilo separado
        threading.Thread(target=proceso_envio_inteligente, daemon=True).start()
    
    def actualizar_progreso_seguro(self, progreso, mensaje):
        """Actualizar progreso con diseño Apple"""
        try:
            # Si progreso es None (pausas), no actualizar barra
            if progreso is not None:
                # TTK ProgressBar con estilo morado
                if hasattr(self, 'progress_bar') and self.progress_bar:
                    self.progress_bar['value'] = progreso
                
                # Canvas fallback con diseño Apple
                elif hasattr(self, 'progress_canvas'):
                    width = self.progress_canvas.winfo_width()
                    if width > 10:
                        self.progress_canvas.delete("all")
                        # Fondo de la barra (track)
                        self.progress_canvas.create_rectangle(0, 0, width, 8, 
                                                            fill=self.colors['bg_input'], 
                                                            outline="")
                        # Barra de progreso morada
                        prog_width = int((progreso / 100) * width)
                        if prog_width > 0:
                            self.progress_canvas.create_rectangle(0, 0, prog_width, 8, 
                                                                fill=self.colors['purple_primary'], 
                                                                outline="")
            
            # Actualizar label con colores dinámicos
            if "error" in mensaje.lower() or "❌" in mensaje:
                color = self.colors['error']
            elif "pausa" in mensaje.lower() or "⏳" in mensaje:
                color = self.colors['warning']
            elif "enviando" in mensaje.lower() or "📊" in mensaje:
                color = self.colors['info']
            elif "completado" in mensaje.lower() or "✅" in mensaje:
                color = self.colors['success']
            else:
                color = self.colors['text_primary']
            
            self.label_estado.config(text=mensaje, fg=color)
            
            if progreso is not None:
                self.log_mensaje(f"📊 {progreso:.1f}% - {mensaje}")
            else:
                self.log_mensaje(f"⏳ {mensaje}")
            
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"Error progreso: {e}")
    
    def procesar_resultados_inteligente(self, resultados):
        """Procesar resultados del envío inteligente"""
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
        self.log_mensaje("🧠 RESUMEN ENVÍO INTELIGENTE")
        self.log_mensaje("="*50)
        
        # Mostrar estrategia usada
        if 'estrategia' in resultados:
            estrategia = resultados['estrategia']
            self.log_mensaje(f"🎯 Estrategia: {estrategia.get('modo', 'AUTOMÁTICO')}")
            self.log_mensaje(f"📝 {estrategia.get('descripcion', 'Envío automático')}")
        
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
        
        # Mensaje usuario con colores actualizados
        if porcentaje >= 90:
            icono, titulo, color = "🎉", "¡Éxito Inteligente!", self.colors['success']
        elif porcentaje >= 70:
            icono, titulo, color = "✅", "Completado Inteligente", self.colors['warning']
        else:
            icono, titulo, color = "⚠️", "Con Problemas", self.colors['error']
        
        mensaje_final = f"{icono} {titulo}\n\n"
        mensaje_final += f"🧠 ENVÍO INTELIGENTE ANTI-SPAM\n\n"
        mensaje_final += f"📊 ESTADÍSTICAS:\n"
        mensaje_final += f"• Total: {total}\n"
        mensaje_final += f"• Exitosos: {exitosos}\n"
        mensaje_final += f"• Fallidos: {fallidos}\n"
        mensaje_final += f"• Éxito: {porcentaje:.1f}%\n"
        
        if 'estrategia' in resultados:
            mensaje_final += f"• Estrategia: {resultados['estrategia'].get('modo', 'AUTO')}\n"
        
        mensaje_final += f"\n📋 Ver detalles en el log."
        
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
        self.btn_estrategia.config(state='normal')
        
        # Progreso al 100%
        try:
            if hasattr(self, 'progress_bar') and self.progress_bar:
                self.progress_bar['value'] = 100
            elif hasattr(self, 'progress_canvas'):
                width = self.progress_canvas.winfo_width()
                if width > 10:
                    self.progress_canvas.delete("all")
                    self.progress_canvas.create_rectangle(0, 0, width, 8, fill=self.colors['success'], outline="")
        except:
            pass
        
        self.log_mensaje("🏁 Proceso finalizado")
    
    def detener_envio(self):
        """Detener envío"""
        if self.enviando:
            respuesta = messagebox.askyesno("⏹️ Detener", 
                                          "¿Detener envío inteligente?\n\n"
                                          "Los enviados no se recuperan.")
            if respuesta:
                self.enviando = False
                self.log_mensaje("⏹️ DETENIDO por usuario")
                self.finalizar_envio()
                self.label_estado.config(text="⏹️ Detenido", fg=self.colors['error'])
    
    def on_closing(self):
        """Cerrar ventana de forma segura"""
        if self.enviando:
            respuesta = messagebox.askyesno("⚠️ Envío en Progreso", 
                                          "¿Cerrar?\n\nSe detendrá el envío inteligente.")
            if not respuesta:
                return
            
            self.enviando = False
            self.log_mensaje("🔄 Cerrando - detenido")
            
            # Esperar un momento para que el hilo termine
            time.sleep(0.5)
        
        self.log_mensaje("👋 Cerrado")
        
        # Destruir la ventana de forma segura
        try:
            self.root.quit()  # Salir del mainloop
            self.root.destroy()  # Destruir la ventana
        except tk.TclError:
            pass  # La ventana ya se cerró
    
    def ejecutar(self):
        """Ejecutar aplicación"""
        print("🚀 Iniciando Email Sender Inteligente...")
        
        # Crear interfaz
        self.crear_interfaz()
        
        # Log inicial
        self.log_mensaje("🧠 Email Sender Pro Inteligente iniciado")
        self.log_mensaje("💡 Haz clic en 'Actualizar Datos' para calcular estrategia")
        
        # Verificar
        if not all([self.excel_mgr, self.file_mgr, self.email_processor, self.email_sender]):
            self.log_mensaje("⚠️ Algunos componentes fallaron")
        else:
            self.log_mensaje("✅ Todos los componentes cargados")
        
        # Configurar cierre
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Listo
        self.root.after(1000, lambda: self.log_mensaje("✅ Sistema inteligente listo"))
        
        # Iniciar
        print("✅ Loop principal")
        self.root.mainloop()

# FUNCIÓN PRINCIPAL
if __name__ == "__main__":
    print("="*60)
    print("🧠 EMAIL SENDER PRO - VERSIÓN INTELIGENTE ANTI-SPAM")
    print("🎯 DISTRIBUCIÓN AUTOMÁTICA SEGÚN CANTIDAD")
    print("="*60)
    
    try:
        print("🔧 Verificando...")
        print(f"📁 Directorio: {os.getcwd()}")
        print(f"📁 Script: {os.path.dirname(os.path.abspath(__file__))}")
        
        app = EmailSenderInteligenteGUI()
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
        print("3. Comando: python src/gui_inteligente.py")
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

# FIN DE LA PARTE 3 - ARCHIVO COMPLETO Y FUNCIONALimport tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import time
from typing import List, Dict
import sys
import os

# Agregar directorio al path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

class EmailSenderInteligenteGUI:
    """GUI Email Sender Pro - CON ENVÍO INTELIGENTE ANTI-SPAM - PARTE 1"""
    
    def __init__(self):
        print("🔧 Inicializando GUI Inteligente...")
        
        # Crear ventana
        self.root = tk.Tk()
        self.root.title("📧 Email Sender Pro - INTELIGENTE ANTI-SPAM")
        
        # Variables
        self.enviando = False
        self.correos_procesados = []
        self.estrategia_actual = None
        
        # Configurar
        self._configurar_ventana()
        self._definir_colores()
        self._aplicar_tema_seguro()
        self._inicializar_managers_seguro()
        
        print("✅ GUI Inteligente inicializada")
    
    def _configurar_ventana(self):
        """Configurar ventana"""
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            width = int(screen_width * 0.9)
            height = int(screen_height * 0.9)
            
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            self.root.geometry(f"{width}x{height}+{x}+{y}")
            self.root.minsize(1300, 900)
            self.root.resizable(True, True)
            
            print(f"📐 Ventana: {width}x{height}")
        except Exception as e:
            print(f"⚠️ Error ventana: {e}")
            self.root.geometry("1300x900")
    
    def _definir_colores(self):
        """Colores profesionales estilo Apple con morado como accent"""
        self.colors = {
            # Backgrounds principales
            'bg': '#1a1a1a',                    # Fondo principal oscuro elegante
            'bg_secondary': '#2d2d2d',          # Fondo secundario
            'bg_card': '#353535',               # Fondo de tarjetas/panels
            'bg_input': '#404040',              # Fondo de inputs
            
            # Texto
            'text_primary': '#ffffff',          # Texto principal blanco
            'text_secondary': '#b0b0b0',        # Texto secundario gris claro
            'text_tertiary': '#808080',         # Texto terciario gris medio
            
            # Purple theme (Email Sender Pro)
            'purple_primary': '#8b5cf6',        # Morado principal
            'purple_light': '#a78bfa',          # Morado claro
            'purple_dark': '#7c3aed',           # morado oscuro
            'purple_bg': '#2d1b69',             # Morado de fondo
            
            # Status colors
            'success': '#10b981',               # Verde éxito
            'warning': '#f59e0b',               # Amarillo advertencia
            'error': '#ef4444',                 # Rojo error
            'info': '#3b82f6',                  # Azul información
            
            # Colores legacy para compatibilidad
            'current_line': '#404040',
            'selection': '#404040',
            'foreground': '#ffffff',
            'comment': '#808080',
            'cyan': '#06b6d4',
            'green': '#10b981',
            'orange': '#f59e0b',
            'pink': '#ec4899',
            'purple': '#8b5cf6',
            'red': '#ef4444',
            'yellow': '#eab308'
        }
    
    def _aplicar_tema_seguro(self):
        """Aplicar tema profesional estilo Apple"""
        try:
            self.root.configure(bg=self.colors['bg'])
            
            self.style = ttk.Style()
            
            try:
                self.style.theme_use('clam')
            except:
                print("⚠️ Usando tema default")
            
            try:
                # Configurar estilos profesionales
                
                # Botones principales con estilo Apple
                self.style.configure('Professional.TButton',
                                   background=self.colors['purple_primary'],
                                   foreground='white',
                                   borderwidth=0,
                                   focuscolor='none',
                                   padding=(20, 12),
                                   font=('SF Pro Display', 11, 'normal'))
                
                self.style.map('Professional.TButton',
                              background=[('active', self.colors['purple_light']),
                                        ('pressed', self.colors['purple_dark'])])
                
                # Botón de acción principal (ENVÍO INTELIGENTE)
                self.style.configure('Primary.TButton',
                                   background=self.colors['purple_primary'],
                                   foreground='white',
                                   borderwidth=0,
                                   focuscolor='none',
                                   padding=(25, 15),
                                   font=('SF Pro Display', 12, 'bold'))
                
                self.style.map('Primary.TButton',
                              background=[('active', self.colors['purple_light']),
                                        ('pressed', self.colors['purple_dark'])])
                
                # Botón secundario
                self.style.configure('Secondary.TButton',
                                   background=self.colors['bg_card'],
                                   foreground=self.colors['text_primary'],
                                   borderwidth=1,
                                   focuscolor='none',
                                   padding=(18, 10),
                                   font=('SF Pro Display', 10, 'normal'))
                
                self.style.map('Secondary.TButton',
                              background=[('active', self.colors['bg_input']),
                                        ('pressed', self.colors['bg_secondary'])])
                
                # Botón de peligro (DETENER)
                self.style.configure('Danger.TButton',
                                   background=self.colors['error'],
                                   foreground='white',
                                   borderwidth=0,
                                   focuscolor='none',
                                   padding=(20, 12),
                                   font=('SF Pro Display', 11, 'normal'))
                
                self.style.map('Danger.TButton',
                              background=[('active', '#dc2626'),
                                        ('pressed', '#b91c1c')])
                
                # Progress bar con morado
                self.style.configure('Purple.Horizontal.TProgressbar',
                                   background=self.colors['purple_primary'],
                                   troughcolor=self.colors['bg_card'],
                                   borderwidth=0,
                                   lightcolor=self.colors['purple_primary'],
                                   darkcolor=self.colors['purple_primary'])
                
                print("✅ Estilos profesionales aplicados")
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
        
        # EmailSender - CORREGIDO PARA USAR email_sender
        try:
            from email_sender import EmailSender
            self.email_sender = EmailSender()
            print("✅ EmailSender OK")
        except Exception as e:
            print(f"❌ EmailSender error: {e}")
            print("🔧 Creando EmailSender funcional...")
            self.email_sender = self._crear_email_sender_funcional()
        
        print("✅ Managers listos")
    
    def _crear_email_sender_funcional(self):
        """EmailSender funcional integrado con lógica inteligente"""
        class EmailSenderFuncional:
            def __init__(self):
                self.conectado = False
                self.MAX_CORREOS_DIARIOS = 400
                self.HORAS_TRABAJO = 8
                self.LIMITE_RAPIDO = 25
                print("📧 EmailSender funcional con lógica inteligente creado")
            
            def conectar_outlook(self):
                try:
                    import win32com.client
                    import pythoncom
                    
                    # ⭐ Inicializar COM
                    pythoncom.CoInitialize()
                    
                    try:
                        # Intentar conectar a instancia existente
                        outlook = win32com.client.GetActiveObject("Outlook.Application")
                        print("✅ Conectado a instancia existente")
                    except:
                        # Crear nueva instancia
                        outlook = win32com.client.Dispatch("Outlook.Application")
                        print("✅ Nueva instancia creada")
                    
                    # Verificar funcionamiento
                    namespace = outlook.GetNamespace("MAPI")
                    inbox = namespace.GetDefaultFolder(6)
                    
                    # Verificar cuentas
                    accounts = namespace.Accounts
                    if accounts.Count == 0:
                        raise Exception("No hay cuentas configuradas en Outlook")
                    
                    cuenta_principal = accounts.Item(1)
                    email_cuenta = getattr(cuenta_principal, 'SmtpAddress', cuenta_principal.DisplayName)
                    
                    self.conectado = True
                    
                    return {
                        'exitoso': True,
                        'mensaje': f'Conectado a Outlook correctamente',
                        'cuenta': email_cuenta,
                        'total_cuentas': accounts.Count
                    }
                    
                except Exception as e:
                    return {
                        'exitoso': False,
                        'mensaje': f'Error Outlook: {str(e)}',
                        'sugerencia': 'Abre Outlook primero'
                    }
            
            def calcular_estrategia_envio(self, total_correos):
                """Calcular estrategia simple"""
                if total_correos <= 2:
                    return {
                        'modo': 'INMEDIATO',
                        'descripcion': 'Envío inmediato sin pausas',
                        'pausa_entre_correos': 5
                    }
                elif total_correos <= self.LIMITE_RAPIDO:
                    return {
                        'modo': 'RÁPIDO',
                        'descripcion': f'Envío rápido con pausas de 30s',
                        'pausa_entre_correos': 30
                    }
                else:
                    return {
                        'modo': 'DISTRIBUIDO',
                        'descripcion': f'Envío distribuido con pausas de 6 minutos',
                        'pausa_entre_correos': 360
                    }
            
            def enviar_correo(self, correo_data, adjuntos=None):
                try:
                    import win32com.client
                    import pythoncom
                    
                    # ⭐ CLAVE: Inicializar COM en cada envío
                    pythoncom.CoInitialize()
                    
                    try:
                        outlook = win32com.client.Dispatch("Outlook.Application")
                        
                        mail = outlook.CreateItem(0)
                        mail.To = correo_data['email']
                        mail.Subject = correo_data['asunto']
                        
                        # ⭐ CONFIGURAR CONTENIDO CON TEXTO NEGRO FORZADO
                        contenido = correo_data['contenido']
                        if '\n' in contenido and '<br>' not in contenido.lower():
                            contenido_html = contenido.replace('\n', '<br>')
                            mail.HTMLBody = f"""
                            <html>
                            <body style="font-family: Arial, sans-serif; font-size: 12pt; color: #000000; background-color: #ffffff;">
                            {contenido_html}
                            </body>
                            </html>
                            """
                        else:
                            if '<html>' in contenido.lower():
                                mail.HTMLBody = contenido
                            else:
                                mail.Body = contenido
                        
                        # ⭐ AGREGAR ADJUNTOS - DEBUGGING COMPLETO
                        adjuntos_agregados = 0
                        if adjuntos:
                            print(f"🔍 DEBUG: Procesando {len(adjuntos)} adjuntos")
                            for i, ruta_adjunto in enumerate(adjuntos):
                                print(f"🔍 DEBUG: Adjunto {i+1}: {ruta_adjunto}")
                                
                                # Convertir a ruta absoluta
                                ruta_absoluta = os.path.abspath(ruta_adjunto)
                                print(f"🔍 DEBUG: Ruta absoluta: {ruta_absoluta}")
                                
                                if os.path.exists(ruta_absoluta):
                                    try:
                                        print(f"📎 Agregando adjunto: {os.path.basename(ruta_absoluta)}")
                                        mail.Attachments.Add(ruta_absoluta)
                                        adjuntos_agregados += 1
                                        print(f"✅ Adjunto agregado exitosamente: {os.path.basename(ruta_absoluta)}")
                                    except Exception as attach_error:
                                        print(f"❌ Error adjuntando {ruta_absoluta}: {attach_error}")
                                        print(f"❌ Tipo error: {type(attach_error).__name__}")
                                else:
                                    print(f"❌ Archivo NO EXISTE: {ruta_absoluta}")
                            
                            print(f"📊 Total adjuntos agregados: {adjuntos_agregados}/{len(adjuntos)}")
                        else:
                            print("📎 No hay adjuntos para procesar")
                        
                        # ⭐ ENVIAR
                        print(f"📤 Enviando correo a {correo_data['email']}...")
                        mail.Send()
                        print(f"✅ Correo enviado exitosamente")
                        
                        return {
                            'exitoso': True,
                            'timestamp': time.strftime('%H:%M:%S'),
                            'email': correo_data['email'],
                            'nombre': correo_data.get('nombre', 'Sin nombre'),
                            'adjuntos_agregados': adjuntos_agregados
                        }
                    
                    finally:
                        # ⭐ LIMPIAR COM
                        try:
                            pythoncom.CoUninitialize()
                        except:
                            pass
                    
                except Exception as e:
                    print(f"❌ ERROR GENERAL en enviar_correo: {e}")
                    print(f"❌ Tipo error: {type(e).__name__}")
                    return {
                        'exitoso': False,
                        'error': str(e),
                        'email': correo_data.get('email', 'desconocido'),
                        'nombre': correo_data.get('nombre', 'Sin nombre')
                    }
            
            def envio_inteligente(self, correos, adjuntos, callback_progreso=None, detener_callback=None):
                """Envío inteligente adaptado con COM corregido"""
                import pythoncom
                
                # ⭐ INICIALIZAR COM en el hilo de envío
                pythoncom.CoInitialize()
                
                try:
                    estrategia = self.calcular_estrategia_envio(len(correos))
                    
                    resultados = {
                        'exitosos': [],
                        'fallidos': [],
                        'total_procesados': 0,
                        'estrategia': estrategia,
                        'inicio': time.time()
                    }
                    
                    pausa = estrategia['pausa_entre_correos']
                    
                    for i, correo in enumerate(correos):
                        if detener_callback and detener_callback():
                            break
                        
                        if callback_progreso:
                            progreso = (i / len(correos)) * 100
                            nombre = correo.get('nombre', 'Sin nombre')
                            callback_progreso(progreso, f"[{estrategia['modo']}] Enviando a {nombre} ({i+1}/{len(correos)})")
                        
                        resultado = self.enviar_correo(correo, adjuntos)
                        
                        if resultado['exitoso']:
                            resultados['exitosos'].append(resultado)
                        else:
                            resultados['fallidos'].append(resultado)
                        
                        resultados['total_procesados'] += 1
                        
                        # Pausa inteligente
                        if i < len(correos) - 1:
                            for segundo in range(pausa):
                                if detener_callback and detener_callback():
                                    break
                                time.sleep(1)
                                
                                if segundo % 30 == 0 and callback_progreso:
                                    tiempo_restante = pausa - segundo
                                    if tiempo_restante > 60:
                                        tiempo_texto = f"{tiempo_restante // 60}m {tiempo_restante % 60}s"
                                    else:
                                        tiempo_texto = f"{tiempo_restante}s"
                                    callback_progreso(progreso, f"Pausa {estrategia['modo'].lower()}: {tiempo_texto}")
                    
                    resultados['fin'] = time.time()
                    resultados['duracion'] = f"{resultados['fin'] - resultados['inicio']:.1f}s"
                    
                    return resultados
                
                finally:
                    # ⭐ LIMPIAR COM al final
                    try:
                        pythoncom.CoUninitialize()
                    except:
                        pass
        
        return EmailSenderFuncional()

# FIN DE LA PARTE 1 - Configuración, colores, estilos y EmailSender funcional

# PARTE 2 - INTERFAZ GRÁFICA Y PANELES

    def crear_interfaz(self):
        """Crear interfaz profesional estilo Apple"""
        print("🏗️ Creando interfaz profesional...")
        
        # Frame principal con padding estilo Apple
        self.main_frame = tk.Frame(self.root, bg=self.colors['bg'], padx=30, pady=25)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Grid con proporciones Apple-like
        self.main_frame.columnconfigure(0, weight=3, minsize=600)  # Contenido principal más ancho
        self.main_frame.columnconfigure(1, weight=2, minsize=400)  # Panel lateral
        
        # Filas con espaciado proporcional
        self.main_frame.rowconfigure(0, weight=0, minsize=80)   # Título
        self.main_frame.rowconfigure(1, weight=0, minsize=140)  # Estrategia
        self.main_frame.rowconfigure(2, weight=1, minsize=400)  # Contenido principal
        self.main_frame.rowconfigure(3, weight=0, minsize=80)   # Botones
        self.main_frame.rowconfigure(4, weight=0, minsize=60)   # Progreso
        
        # Crear secciones con diseño Apple
        self.crear_titulo_profesional()
        self.crear_estrategia_profesional()
        self.crear_contenido_profesional()
        self.crear_botones_profesionales()
        self.crear_progreso_profesional()
        
        print("✅ Interfaz profesional creada")
    
    def crear_titulo_profesional(self):
        """Título con diseño Apple"""
        titulo_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        titulo_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 25))
        
        # Título principal con tipografía Apple
        titulo_principal = tk.Label(titulo_frame, 
                                  text="📧 EMAIL SENDER PRO", 
                                  font=('SF Pro Display', 32, 'bold'), 
                                  bg=self.colors['bg'], 
                                  fg=self.colors['purple_primary'])
        titulo_principal.pack(pady=(0, 8))
        
        # Subtítulo elegante
        subtitulo = tk.Label(titulo_frame, 
                           text="🧠 Envío Inteligente Anti-Spam • Distribución Automática", 
                           font=('SF Pro Text', 14, 'normal'), 
                           bg=self.colors['bg'], 
                           fg=self.colors['text_secondary'])
        subtitulo.pack()
    
    def crear_estrategia_profesional(self):
        """Panel de estrategia con diseño Apple"""
        # Frame contenedor con esquinas redondeadas simuladas
        strategy_container = tk.Frame(self.main_frame, bg=self.colors['bg'])
        strategy_container.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 20))
        strategy_container.columnconfigure(0, weight=1)
        
        # Header del panel
        header_frame = tk.Frame(strategy_container, bg=self.colors['bg_card'], height=50)
        header_frame.pack(fill=tk.X, pady=(0, 2))
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(header_frame, 
                              text="🎯 Estrategia de Envío Inteligente",
                              font=('SF Pro Display', 16, 'bold'), 
                              bg=self.colors['bg_card'], 
                              fg=self.colors['purple_primary'])
        header_label.pack(pady=12)
        
        # Contenido del panel
        content_frame = tk.Frame(strategy_container, bg=self.colors['bg_card'])
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        
        self.text_estrategia = scrolledtext.ScrolledText(content_frame, 
                                                        height=6, 
                                                        font=('SF Mono', 11),
                                                        bg=self.colors['bg_input'], 
                                                        fg=self.colors['text_primary'], 
                                                        relief='flat',
                                                        borderwidth=0,
                                                        insertbackground=self.colors['purple_primary'])
        self.text_estrategia.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        # Mensaje inicial estilizado
        mensaje_inicial = """🎯 ESTRATEGIA INTELIGENTE ANTI-SPAM

✅ 1-2 correos → Envío INMEDIATO (sin pausas)
⚡ 3-25 correos → Envío RÁPIDO (pausas de 30 segundos)  
📦 26+ correos → Envío DISTRIBUIDO (lotes con pausas de 6 minutos)

💡 Haz clic en 'Actualizar Datos' para calcular la estrategia específica"""
        
        self.text_estrategia.insert(1.0, mensaje_inicial)
        self.text_estrategia.config(state='disabled')
    
    def crear_contenido_profesional(self):
        """Área de contenido con diseño Apple"""
        # Panel principal izquierdo
        left_panel = tk.Frame(self.main_frame, bg=self.colors['bg'])
        left_panel.grid(row=2, column=0, sticky='nsew', padx=(0, 15))
        left_panel.rowconfigure(0, weight=0, minsize=160)  # Estado archivos
        left_panel.rowconfigure(1, weight=0, minsize=120)  # Campaña activa  
        left_panel.rowconfigure(2, weight=1, minsize=300)  # Vista previa
        left_panel.columnconfigure(0, weight=1)
        
        self.crear_panel_archivos(left_panel)
        self.crear_panel_campana(left_panel)
        self.crear_panel_vista_previa(left_panel)
        
        # Panel lateral derecho
        right_panel = tk.Frame(self.main_frame, bg=self.colors['bg'])
        right_panel.grid(row=2, column=1, sticky='nsew')
        right_panel.rowconfigure(0, weight=0, minsize=200)  # Adjuntos
        right_panel.rowconfigure(1, weight=1, minsize=350)  # Log
        right_panel.columnconfigure(0, weight=1)
        
        self.crear_panel_adjuntos(right_panel)
        self.crear_panel_log(right_panel)
    
    def crear_panel_archivos(self, parent):
        """Panel de estado de archivos estilo Apple"""
        # Frame principal con fondo tipo card
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        card_frame.columnconfigure(0, weight=1)
        
        # Header del panel
        header = tk.Label(card_frame, 
                         text="📊 Estado de Archivos Excel",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido scrollable
        self.text_archivos = scrolledtext.ScrolledText(card_frame, 
                                                      height=6, 
                                                      font=('SF Mono', 10),
                                                      bg=self.colors['bg_input'], 
                                                      fg=self.colors['text_primary'], 
                                                      relief='flat',
                                                      borderwidth=0)
        self.text_archivos.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_panel_campana(self, parent):
        """Panel de campaña activa estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=1, column=0, sticky='ew', pady=(0, 15))
        card_frame.columnconfigure(0, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="🎯 Campaña Activa",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Información de campaña
        info_frame = tk.Frame(card_frame, bg=self.colors['bg_card'])
        info_frame.pack(fill=tk.BOTH, padx=20, pady=(0, 15))
        info_frame.columnconfigure(0, weight=1)
        
        self.label_campana_nombre = tk.Label(info_frame, 
                                           text="📋 Campaña: Cargando...",
                                           font=('SF Pro Text', 12, 'bold'), 
                                           bg=self.colors['bg_card'], 
                                           fg=self.colors['text_primary'], 
                                           anchor='w')
        self.label_campana_nombre.grid(row=0, column=0, sticky='ew', pady=(0, 5))
        
        self.label_campana_asunto = tk.Label(info_frame, 
                                           text="📧 Asunto: Cargando...",
                                           font=('SF Pro Text', 11, 'normal'), 
                                           bg=self.colors['bg_card'], 
                                           fg=self.colors['text_secondary'], 
                                           anchor='w', 
                                           wraplength=500)
        self.label_campana_asunto.grid(row=1, column=0, sticky='ew', pady=(0, 5))
        
        self.label_campana_info = tk.Label(info_frame, 
                                         text="📝 Info: Cargando...",
                                         font=('SF Pro Text', 10, 'normal'), 
                                         bg=self.colors['bg_card'], 
                                         fg=self.colors['text_tertiary'], 
                                         anchor='w', 
                                         wraplength=500)
        self.label_campana_info.grid(row=2, column=0, sticky='ew')
    
    def crear_panel_vista_previa(self, parent):
        """Panel de vista previa estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=2, column=0, sticky='nsew')
        card_frame.columnconfigure(0, weight=1)
        card_frame.rowconfigure(1, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="📧 Vista Previa del Primer Correo",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido
        self.text_preview = scrolledtext.ScrolledText(card_frame, 
                                                     font=('SF Mono', 10),
                                                     bg=self.colors['bg_input'], 
                                                     fg=self.colors['text_primary'], 
                                                     relief='flat',
                                                     borderwidth=0)
        self.text_preview.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_panel_adjuntos(self, parent):
        """Panel de adjuntos estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        card_frame.columnconfigure(0, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="📎 Archivos Adjuntos",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido
        self.text_adjuntos = tk.Text(card_frame, 
                                    height=8, 
                                    font=('SF Mono', 9),
                                    bg=self.colors['bg_input'], 
                                    fg=self.colors['warning'], 
                                    relief='flat', 
                                    borderwidth=0,
                                    wrap=tk.WORD)
        self.text_adjuntos.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_panel_log(self, parent):
        """Panel de log estilo Apple"""
        card_frame = tk.Frame(parent, bg=self.colors['bg_card'])
        card_frame.grid(row=1, column=0, sticky='nsew')
        card_frame.columnconfigure(0, weight=1)
        card_frame.rowconfigure(1, weight=1)
        
        # Header
        header = tk.Label(card_frame, 
                         text="📋 Log en Tiempo Real",
                         font=('SF Pro Display', 14, 'bold'), 
                         bg=self.colors['bg_card'], 
                         fg=self.colors['purple_primary'],
                         anchor='w')
        header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        # Contenido
        self.text_log = scrolledtext.ScrolledText(card_frame, 
                                                 font=('SF Mono', 9),
                                                 bg=self.colors['bg_input'], 
                                                 fg=self.colors['info'], 
                                                 relief='flat',
                                                 borderwidth=0)
        self.text_log.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
    
    def crear_botones_profesionales(self):
        """Botones con diseño Apple profesional"""
        buttons_container = tk.Frame(self.main_frame, bg=self.colors['bg'])
        buttons_container.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(25, 0))
        
        # Grid para botones con espaciado proporcional
        for i in range(5):
            buttons_container.columnconfigure(i, weight=1, minsize=180)
        
        # Botones con estilos específicos
        self.btn_actualizar = ttk.Button(buttons_container, 
                                        text="🔄 Actualizar Datos", 
                                        command=self.actualizar_datos,
                                        style='Secondary.TButton')
        self.btn_actualizar.grid(row=0, column=0, padx=8, pady=15, sticky='ew')
        
        self.btn_estrategia = ttk.Button(buttons_container, 
                                        text="🎯 Ver Estrategia", 
                                        command=self.mostrar_estrategia,
                                        style='Secondary.TButton')
        self.btn_estrategia.grid(row=0, column=1, padx=8, pady=15, sticky='ew')
        
        self.btn_preview = ttk.Button(buttons_container, 
                                     text="👁️ Vista Previa", 
                                     command=self.vista_previa_completa,
                                     style='Professional.TButton')
        self.btn_preview.grid(row=0, column=2, padx=8, pady=15, sticky='ew')
        
        self.btn_enviar = ttk.Button(buttons_container, 
                                    text="🚀 ENVÍO INTELIGENTE", 
                                    command=self.enviar_correos_inteligente,
                                    style='Primary.TButton')
        self.btn_enviar.grid(row=0, column=3, padx=8, pady=15, sticky='ew')
        
        self.btn_detener = ttk.Button(buttons_container, 
                                     text="⏹️ DETENER", 
                                     state="disabled", 
                                     command=self.detener_envio,
                                     style='Danger.TButton')
        self.btn_detener.grid(row=0, column=4, padx=8, pady=15, sticky='ew')
    
    def crear_progreso_profesional(self):
        """Barra de progreso estilo Apple"""
        progress_container = tk.Frame(self.main_frame, bg=self.colors['bg'])
        progress_container.grid(row=4, column=0, columnspan=2, sticky='ew', pady=(20, 0))
        
        # Label de estado con tipografía Apple
        self.label_estado = tk.Label(progress_container, 
                                   text="✅ Sistema listo - Haz clic en 'Actualizar Datos'",
                                   font=('SF Pro Text', 13, 'normal'), 
                                   bg=self.colors['bg'], 
                                   fg=self.colors['success'])
        self.label_estado.pack(pady=(0, 15))
        
        # Progress bar container con fondo
        progress_bg = tk.Frame(progress_container, bg=self.colors['bg_card'], height=8)
        progress_bg.pack(fill=tk.X, padx=60, pady=(0, 5))
        
        try:
            # Progress bar con estilo morado
            self.progress_bar = ttk.Progressbar(progress_bg, 
                                              mode='determinate', 
                                              length=400,
                                              style='Purple.Horizontal.TProgressbar')
            self.progress_bar.pack(fill=tk.X, padx=2, pady=2)
            print("✅ ProgressBar profesional OK")
        except Exception as e:
            print(f"⚠️ Error ProgressBar: {e}")
            # Fallback a canvas con diseño Apple
            self.progress_canvas = tk.Canvas(progress_bg, 
                                           height=8, 
                                           bg=self.colors['bg_card'],
                                           highlightthickness=0)
            self.progress_canvas.pack(fill=tk.X, padx=2, pady=2)
            self.progress_bar = None
            print("✅ Canvas como ProgressBar profesional")
    
    def log_mensaje(self, mensaje):
        """Log seguro que verifica si la ventana existe"""
        try:
            if hasattr(self, 'text_log') and self.text_log.winfo_exists():
                timestamp = time.strftime('%H:%M:%S')
                self.text_log.insert(tk.END, f"[{timestamp}] {mensaje}\n")
                self.text_log.see(tk.END)
                self.root.update_idletasks()
            else:
                # Si no hay interfaz, imprimir en consola
                print(f"[{time.strftime('%H:%M:%S')}] {mensaje}")
        except (tk.TclError, AttributeError) as e:
            # Si hay error con la interfaz, imprimir en consola
            print(f"[{time.strftime('%H:%M:%S')}] {mensaje}")
            print(f"Log error: {e}")

# FIN DE LA PARTE 2 - Interfaz gráfica completa