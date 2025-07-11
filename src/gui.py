import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import time
from typing import List, Dict
from excel_manager import ExcelManager
from file_manager import FileManager
from email_processor import EmailProcessor
from email_sender import EmailSender

class EmailSenderGUI:
    """Interfaz gráfica principal del Email Sender - TEMA DRACULA COMPACTO"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("📧 Email Sender Pro")
        self.root.geometry("1200x720")  # Tamaño laptop estándar
        self.root.resizable(True, True)
        
        # Colores Dracula Theme
        self.colors = {
            'bg': '#282a36',           # Fondo principal
            'current_line': '#44475a', # Línea actual
            'selection': '#44475a',    # Selección
            'foreground': '#f8f8f2',   # Texto principal
            'comment': '#6272a4',      # Comentarios
            'cyan': '#8be9fd',         # Cyan
            'green': '#50fa7b',        # Verde
            'orange': '#ffb86c',       # Naranja
            'pink': '#ff79c6',         # Rosa
            'purple': '#bd93f9',       # Púrpura
            'red': '#ff5555',          # Rojo
            'yellow': '#f1fa8c'        # Amarillo
        }
        
        # Aplicar tema Dracula
        self.aplicar_tema_dracula()

        # Managers
        self.excel_mgr = ExcelManager()
        self.file_mgr = FileManager()
        self.email_processor = EmailProcessor()
        self.email_sender = EmailSender()

        # Variables
        self.enviando = False
        self.correos_procesados = []

        # Crear interfaz
        self.crear_interfaz()
        self.actualizar_datos()
    
    def aplicar_tema_dracula(self):
        """Aplica el tema Dracula a toda la aplicación"""
        # Configurar ventana principal
        self.root.configure(bg=self.colors['bg'])
        
        # Configurar estilo ttk
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configurar solo estilos básicos que funcionan
        self.style.configure('TButton', 
                       background=self.colors['current_line'],
                       foreground=self.colors['foreground'],
                       borderwidth=1,
                       focuscolor='none')
        
        self.style.map('TButton',
                 background=[('active', self.colors['purple'])])
        
        self.style.configure('Iniciar.TButton', 
                       background=self.colors['green'],
                       foreground=self.colors['bg'],
                       font=('Arial', 10, 'bold'))
        
        self.style.configure('Detener.TButton', 
                       background=self.colors['red'],
                       foreground=self.colors['bg'],
                       font=('Arial', 10, 'bold'))
        
        self.style.configure('TProgressbar',
                       background=self.colors['purple'],
                       troughcolor=self.colors['current_line'])
    
    def crear_interfaz(self):
        """Crea todos los elementos de la interfaz - VERSIÓN COMPACTA DRACULA"""
        
        # Frame principal con scroll
        canvas = tk.Canvas(self.root, bg=self.colors['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Frame principal
        main_frame = tk.Frame(scrollable_frame, bg=self.colors['bg'], padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === TÍTULO COMPACTO ===
        titulo_label = ttk.Label(
            main_frame, 
            text="📧 EMAIL SENDER PRO", 
            font=('Arial', 14, 'bold'),
            foreground=self.colors['purple'],
            background=self.colors['bg']
        )
        titulo_label.pack(pady=(0, 10))
        
        # === SECCIÓN 1: ESTADO DE ARCHIVOS (ULTRA COMPACTA) ===
        archivos_frame = tk.LabelFrame(
            main_frame, 
            text="📊 Archivos", 
            bg=self.colors['bg'],
            fg=self.colors['foreground'],
            bd=1,
            relief='solid'
        )
        archivos_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.label_archivos = tk.Label(
            archivos_frame, 
            text="Cargando...", 
            font=('Consolas', 8),
            justify=tk.LEFT,
            bg=self.colors['bg'],
            fg=self.colors['foreground'],
            wraplength=1000
        )
        self.label_archivos.pack(anchor=tk.W, fill=tk.X, padx=8, pady=8)
        
        # === SECCIÓN 2: CAMPAÑA ACTIVA (COMPACTA) ===
        campana_frame = tk.LabelFrame(
            main_frame, 
            text="🎯 Campaña", 
            bg=self.colors['bg'],
            fg=self.colors['foreground'],
            bd=1,
            relief='solid'
        )
        campana_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.label_campana = tk.Label(
            campana_frame, 
            text="Cargando...", 
            font=('Consolas', 8),
            justify=tk.LEFT,
            bg=self.colors['bg'],
            fg=self.colors['cyan'],
            wraplength=1000
        )
        self.label_campana.pack(anchor=tk.W, fill=tk.X, padx=8, pady=8)
        
        # === SECCIÓN 3: ADJUNTOS (COMPACTA) ===
        adjuntos_frame = tk.LabelFrame(
            main_frame, 
            text="📎 Adjuntos", 
            bg=self.colors['bg'],
            fg=self.colors['foreground'],
            bd=1,
            relief='solid'
        )
        adjuntos_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.label_adjuntos = tk.Label(
            adjuntos_frame, 
            text="Cargando...", 
            font=('Consolas', 8),
            justify=tk.LEFT,
            bg=self.colors['bg'],
            fg=self.colors['orange'],
            wraplength=1000
        )
        self.label_adjuntos.pack(anchor=tk.W, fill=tk.X, padx=8, pady=8)
        
        # === SECCIÓN 4: VISTA PREVIA (MINI) ===
        preview_frame = tk.LabelFrame(
            main_frame, 
            text="📧 Preview", 
            bg=self.colors['bg'],
            fg=self.colors['foreground'],
            bd=1,
            relief='solid'
        )
        preview_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.text_preview = tk.Text(
            preview_frame, 
            height=4,  # Solo 4 líneas
            font=('Consolas', 8),
            wrap=tk.WORD,
            bg=self.colors['current_line'],
            fg=self.colors['foreground'],
            insertbackground=self.colors['foreground'],
            selectbackground=self.colors['selection'],
            relief='flat'
        )
        self.text_preview.pack(fill=tk.X, padx=8, pady=8)
        
        # === SECCIÓN 5: CONTROLES (COMPACTOS) ===
        controles_frame = tk.LabelFrame(
            main_frame, 
            text="🎮 Controles", 
            bg=self.colors['bg'],
            fg=self.colors['foreground'],
            bd=1,
            relief='solid'
        )
        controles_frame.pack(fill=tk.X, pady=(5, 5))
        
        # Grid de botones 2x2
        botones_grid = tk.Frame(controles_frame, bg=self.colors['bg'])
        botones_grid.pack(fill=tk.X, padx=10, pady=10)
        
        # Configurar grid
        for i in range(4):
            botones_grid.columnconfigure(i, weight=1)
        
        # Primera fila
        self.btn_actualizar = ttk.Button(
            botones_grid, 
            text="🔄 Actualizar", 
            command=self.actualizar_datos,
            width=15
        )
        self.btn_actualizar.grid(row=0, column=0, padx=2, pady=2, sticky='ew')
        
        self.btn_vista_previa = ttk.Button(
            botones_grid, 
            text="👁️ Preview", 
            command=self.mostrar_vista_previa_completa,
            width=15
        )
        self.btn_vista_previa.grid(row=0, column=1, padx=2, pady=2, sticky='ew')
        
        # Segunda fila - Botones principales
        self.btn_iniciar = ttk.Button(
            botones_grid, 
            text="🚀 ENVIAR", 
            command=self.iniciar_envio,
            style='Iniciar.TButton',
            width=15
        )
        self.btn_iniciar.grid(row=1, column=0, padx=2, pady=5, sticky='ew')
        
        self.btn_detener = ttk.Button(
            botones_grid, 
            text="⏹️ STOP", 
            command=self.detener_envio,
            state="disabled",
            style='Detener.TButton',
            width=15
        )
        self.btn_detener.grid(row=1, column=1, padx=2, pady=5, sticky='ew')
        
        # === SECCIÓN 6: PROGRESO (MINI) ===
        progress_frame = tk.LabelFrame(
            main_frame, 
            text="📊 Progreso", 
            bg=self.colors['bg'],
            fg=self.colors['foreground'],
            bd=1,
            relief='solid'
        )
        progress_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.label_progreso = tk.Label(
            progress_frame, 
            text="✅ Sistema listo", 
            font=('Arial', 9, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['green']
        )
        self.label_progreso.pack(pady=(8, 5), padx=8)
        
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, padx=8)
        
        # === ESTADÍSTICAS COMPACTAS ===
        self.label_stats = tk.Label(
            progress_frame, 
            text="", 
            font=('Consolas', 7),
            bg=self.colors['bg'],
            fg=self.colors['comment']
        )
        self.label_stats.pack(pady=(3, 8), padx=8)
        
        # Configurar scroll con mouse
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    def actualizar_datos(self):
        """Actualiza todos los datos - VERSIÓN COMPACTA"""
        try:
            # Estado de archivos (solo lo esencial)
            resumen_excel = self.excel_mgr.obtener_resumen()
            lineas_importantes = []
            for linea in resumen_excel.split('\n'):
                if any(x in linea for x in ['✅', '❌', 'Campaña activa:', 'Total clientes:', 'Email:']):
                    lineas_importantes.append(linea)
            texto_compacto = '\n'.join(lineas_importantes[:6])
            self.label_archivos.config(text=texto_compacto)
            
            # Campaña (una línea)
            campanas = self.excel_mgr.cargar_campanas()
            if 'error' in campanas:
                self.label_campana.config(text=f"❌ Error: {campanas['error']}")
            elif campanas['activa']:
                campana = campanas['activa']
                texto_campana = f"📋 {campana['nombre']} | 📧 {campana['asunto'][:60]}..."
                self.label_campana.config(text=texto_campana)
            else:
                self.label_campana.config(text="⚠️ Sin campaña activa")
            
            # Adjuntos (resumen)
            resumen_adjuntos = self.file_mgr.obtener_resumen()
            lineas_adj = []
            for linea in resumen_adjuntos.split('\n'):
                if any(x in linea for x in ['archivo(s) encontrado', 'Tamaño total:', '✅', '❌']):
                    lineas_adj.append(linea)
            self.label_adjuntos.config(text='\n'.join(lineas_adj[:3]))
            
            # Vista previa mini
            self.actualizar_vista_previa_mini()
            
            # Estadísticas
            self.actualizar_estadisticas_mini()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")
    
    def actualizar_vista_previa_mini(self):
        """Vista previa ultra compacta"""
        try:
            campanas = self.excel_mgr.cargar_campanas()
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            if ('error' in campanas or 'error' in clientes or 'error' in config 
                or not campanas['activa']):
                self.text_preview.delete(1.0, tk.END)
                self.text_preview.insert(1.0, "❌ Datos incompletos\n\nRevisa archivos Excel")
                return
            
            # Un solo correo de muestra
            if clientes['clientes']:
                cliente = clientes['clientes'][0]
                texto_mini = (
                    f"Para: {cliente.get('Email', 'N/A')[:40]}...\n"
                    f"Asunto: {campanas['activa']['asunto'][:50]}...\n"
                    f"💡 Ver completo con 'Preview'"
                )
                self.text_preview.delete(1.0, tk.END)
                self.text_preview.insert(1.0, texto_mini)
            
        except Exception as e:
            self.text_preview.delete(1.0, tk.END)
            self.text_preview.insert(1.0, f"❌ Error: {str(e)[:100]}")
    
    def actualizar_estadisticas_mini(self):
        """Estadísticas en una línea"""
        try:
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            if 'error' not in clientes and 'error' not in config:
                total = len(clientes['clientes'])
                horas = config['config'].get('Horas_Para_Enviar_Todo', '?')
                stats = f"👥 {total} destinatarios | ⏱️ {horas}h estimadas"
                self.label_stats.config(text=stats)
            else:
                self.label_stats.config(text="⚠️ Estadísticas no disponibles")
                
        except Exception as e:
            self.label_stats.config(text=f"❌ Error stats: {str(e)[:50]}")
    
    def mostrar_vista_previa_completa(self):
        """Vista previa completa en ventana separada con tema Dracula"""
        try:
            campanas = self.excel_mgr.cargar_campanas()
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            if ('error' in campanas or 'error' in clientes or 'error' in config 
                or not campanas['activa']):
                messagebox.showerror("Error", "Datos incompletos")
                return
            
            # Procesar correos
            correos = self.email_processor.procesar_lista_clientes(
                clientes['clientes'],
                campanas['activa'],
                config['config']
            )
            
            if not correos:
                messagebox.showwarning("Advertencia", "Sin correos válidos")
                return
            
            # Ventana con tema Dracula
            ventana = tk.Toplevel(self.root)
            ventana.title("Vista Previa Completa")
            ventana.geometry("900x600")
            ventana.configure(bg=self.colors['bg'])
            
            frame = tk.Frame(ventana, bg=self.colors['bg'], padx=10, pady=10)
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Título
            titulo = tk.Label(
                frame, 
                text=f"📧 {len(correos)} correos preparados",
                font=('Arial', 12, 'bold'),
                bg=self.colors['bg'],
                fg=self.colors['purple']
            )
            titulo.pack(pady=(0, 10))
            
            # Área de texto
            text_widget = scrolledtext.ScrolledText(
                frame, 
                font=('Consolas', 9),
                wrap=tk.WORD,
                bg=self.colors['current_line'],
                fg=self.colors['foreground'],
                insertbackground=self.colors['foreground'],
                selectbackground=self.colors['selection']
            )
            text_widget.pack(fill=tk.BOTH, expand=True)
            
            # Contenido (primeros 5 correos)
            contenido = f"📧 VISTA PREVIA - {len(correos)} CORREOS TOTALES\n"
            contenido += "=" * 60 + "\n\n"
            
            for i, correo in enumerate(correos[:5]):
                contenido += f"📩 CORREO #{correo['indice']}:\n"
                contenido += f"Para: {correo['email']}\n"
                contenido += f"Nombre: {correo['nombre']}\n"
                contenido += f"Asunto: {correo['asunto']}\n"
                contenido += f"Contenido:\n{correo['contenido'][:200]}...\n"
                contenido += "-" * 50 + "\n\n"
            
            if len(correos) > 5:
                contenido += f"... y {len(correos) - 5} correos más\n"
            
            text_widget.insert(1.0, contenido)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")
    
    def iniciar_envio(self):
        """Inicia el proceso de envío"""
        # Verificar Outlook
        conexion = self.email_sender.conectar_outlook()
        if not conexion['exitoso']:
            messagebox.showerror("Error", f"Sin conexión Outlook:\n{conexion['mensaje']}")
            return
        
        try:
            # Cargar datos
            campanas = self.excel_mgr.cargar_campanas()
            clientes = self.excel_mgr.cargar_clientes()
            config = self.excel_mgr.cargar_configuracion()
            
            if 'error' in campanas or 'error' in clientes or 'error' in config:
                messagebox.showerror("Error", "Problemas con archivos Excel")
                return
            
            if not campanas['activa']:
                messagebox.showerror("Error", "Sin campaña activa")
                return
            
            # Procesar
            correos = self.email_processor.procesar_lista_clientes(
                clientes['clientes'],
                campanas['activa'],
                config['config']
            )
            
            if not correos:
                messagebox.showerror("Error", "Sin correos válidos")
                return
            
            # Verificar adjuntos
            adjuntos = self.file_mgr.obtener_archivos_validos()
            scan_adjuntos = self.file_mgr.escanear_adjuntos()
            
            if not scan_adjuntos['tamaño_ok']:
                messagebox.showerror("Error", "Adjuntos muy grandes")
                return
            
        except Exception as e:
            messagebox.showerror("Error", f"Error preparando: {str(e)}")
            return
        
        # Confirmar
        respuesta = messagebox.askyesno(
            "🚀 Confirmar", 
            f"¿Enviar {len(correos)} correos?\n\n"
            f"Campaña: {campanas['activa']['nombre']}\n"
            f"Adjuntos: {len(adjuntos)} archivos"
        )
        
        if respuesta:
            self.correos_procesados = correos
            self.iniciar_envio_real(correos, adjuntos)
    
    def iniciar_envio_real(self, correos: List[Dict], adjuntos: List[str]):
        """Ejecuta el envío real"""
        self.enviando = True
        self.btn_iniciar.config(state="disabled")
        self.btn_detener.config(state="normal")
        self.progress_bar['value'] = 0
        self.label_progreso.config(text="🔄 Conectando...", fg=self.colors['orange'])

        threading.Thread(
            target=self._proceso_envio_real, 
            args=(correos, adjuntos), 
            daemon=True
        ).start()

    def _proceso_envio_real(self, correos: List[Dict], adjuntos: List[str]):
        """Proceso de envío en hilo separado"""
        try:
            def callback_progreso(progreso, mensaje):
                self.root.after(0, self._actualizar_progreso, progreso, mensaje)

            def detener_callback():
                return not self.enviando

            resultados = self.email_sender.envio_por_lotes(
                correos,
                adjuntos,
                callback_progreso=callback_progreso,
                detener_callback=detener_callback
            )

            self.root.after(0, self._procesar_resultados, resultados)

        except Exception as e:
            self.root.after(0, self._mostrar_error, str(e))

    def _procesar_resultados(self, resultados):
        """Procesa resultados finales"""
        self._finalizar_envio()
        
        exitosos = len(resultados['exitosos'])
        fallidos = len(resultados['fallidos'])
        total = exitosos + fallidos
        
        mensaje = f"✅ COMPLETADO\n\n"
        mensaje += f"Total: {total}\n"
        mensaje += f"Exitosos: {exitosos}\n"
        mensaje += f"Fallidos: {fallidos}\n"
        mensaje += f"Éxito: {(exitosos/total*100):.1f}%"
        
        messagebox.showinfo("Resultado", mensaje)

    def _mostrar_error(self, error_msg):
        """Muestra error y restaura interfaz"""
        self._finalizar_envio()
        messagebox.showerror("Error", error_msg)

    def _actualizar_progreso(self, progreso, mensaje):
        """Actualiza progreso"""
        self.progress_bar['value'] = progreso
        # Acortar mensaje si es muy largo
        if len(mensaje) > 40:
            mensaje = mensaje[:37] + "..."
        self.label_progreso.config(text=mensaje, fg=self.colors['cyan'])
        self.root.update_idletasks()

    def _finalizar_envio(self):
        """Restaura interfaz"""
        self.btn_iniciar.config(state="normal")
        self.btn_detener.config(state="disabled")
        self.enviando = False
        self.label_progreso.config(text="✅ Finalizado", fg=self.colors['green'])
    
    def detener_envio(self):
        """Detiene el envío"""
        self.enviando = False
        self.label_progreso.config(text="⏹️ Detenido", fg=self.colors['red'])
        self.btn_iniciar.config(state="normal")
        self.btn_detener.config(state="disabled")

    def ejecutar(self):
        """Inicia la aplicación"""
        self.root.mainloop()

# Función principal
if __name__ == "__main__":
    print("🚀 Iniciando Email Sender Pro...")
    app = EmailSenderGUI()
    app.ejecutar()