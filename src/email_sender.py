import win32com.client
import pythoncom
import time
import os
from typing import List, Dict, Optional, Callable
import logging
from datetime import datetime

class EmailSender:
    """EmailSender CORREGIDO - Soluciona error COM"""
    
    def __init__(self):
        self.outlook = None
        self.conectado = False
        self.logger = self._configurar_logger()
        
    def _configurar_logger(self):
        """Configurar logging"""
        logger = logging.getLogger('EmailSender')
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            os.makedirs('reportes', exist_ok=True)
            fecha = datetime.now().strftime('%Y-%m-%d')
            archivo_log = f'reportes/envios_{fecha}.log'
            
            file_handler = logging.FileHandler(archivo_log, encoding='utf-8')
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            
        return logger
    
    def conectar_outlook(self) -> Dict:
        """Conectar con Outlook - COM INICIALIZADO"""
        try:
            self.logger.info("🔄 Conectando con Outlook...")
            
            # ⭐ CLAVE: Inicializar COM en el hilo actual
            pythoncom.CoInitialize()
            
            try:
                # Intentar conectar a instancia existente
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                self.logger.info("✅ Conectado a instancia existente")
            except:
                # Crear nueva instancia
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.logger.info("✅ Nueva instancia creada")
            
            # Verificar funcionamiento
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # Bandeja entrada
            
            # Verificar cuentas
            accounts = namespace.Accounts
            if accounts.Count == 0:
                raise Exception("No hay cuentas configuradas en Outlook")
            
            cuenta_principal = accounts.Item(1)
            email_cuenta = getattr(cuenta_principal, 'SmtpAddress', cuenta_principal.DisplayName)
            
            self.conectado = True
            self.logger.info(f"✅ Conectado - Cuenta: {email_cuenta}")
            
            return {
                'exitoso': True,
                'mensaje': f'Conectado a Outlook correctamente',
                'cuenta': email_cuenta,
                'total_cuentas': accounts.Count
            }
            
        except Exception as e:
            self.conectado = False
            error_msg = str(e)
            self.logger.error(f"❌ Error conexión: {error_msg}")
            
            return {
                'exitoso': False,
                'mensaje': error_msg,
                'sugerencia': self._obtener_sugerencia_error(error_msg)
            }
    
    def _obtener_sugerencia_error(self, error_msg: str) -> str:
        """Sugerencias según el error"""
        error_lower = error_msg.lower()
        
        if "class not registered" in error_lower:
            return "Instala/repara Microsoft Office"
        elif "rpc server" in error_lower:
            return "Abre Outlook manualmente primero"
        elif "no hay cuentas" in error_lower:
            return "Configura una cuenta en Outlook"
        elif "access denied" in error_lower:
            return "Ejecuta como administrador"
        else:
            return "Verifica que Outlook esté instalado y funcionando"
    
    def enviar_correo(self, correo_data: Dict, adjuntos: List[str] = None) -> Dict:
        """Enviar correo individual - COM SEGURO"""
        if not self.conectado:
            return {
                'exitoso': False,
                'error': 'No hay conexión con Outlook'
            }
        
        try:
            # ⭐ CLAVE: Asegurar COM en cada envío
            pythoncom.CoInitialize()
            
            # Validar datos
            if not correo_data.get('email'):
                raise ValueError("Email del destinatario requerido")
            if not correo_data.get('asunto'):
                raise ValueError("Asunto del correo requerido")
            if not correo_data.get('contenido'):
                raise ValueError("Contenido del correo requerido")
            
            # Crear correo
            mail = self.outlook.CreateItem(0)  # olMailItem
            
            # Configurar destinatario
            mail.To = correo_data['email'].strip()
            
            # Configurar asunto
            mail.Subject = correo_data['asunto'].strip()
            
            # Configurar contenido
            contenido = correo_data['contenido']
            
            # Usar HTMLBody para mejor formato
            if '\n' in contenido and '<br>' not in contenido.lower():
                contenido_html = contenido.replace('\n', '<br>')
                mail.HTMLBody = f"""
                <html>
                <body style="font-family: Arial, sans-serif; font-size: 12pt;">
                {contenido_html}
                </body>
                </html>
                """
            else:
                if '<html>' in contenido.lower():
                    mail.HTMLBody = contenido
                else:
                    mail.Body = contenido
            
            # Agregar adjuntos
            adjuntos_agregados = 0
            if adjuntos:
                for ruta_adjunto in adjuntos:
                    if os.path.exists(ruta_adjunto):
                        try:
                            mail.Attachments.Add(ruta_adjunto)
                            adjuntos_agregados += 1
                            self.logger.info(f"📎 Adjunto: {os.path.basename(ruta_adjunto)}")
                        except Exception as attach_error:
                            self.logger.warning(f"⚠️ No se pudo adjuntar {ruta_adjunto}: {attach_error}")
            
            # ⭐ ENVIAR EL CORREO
            self.logger.info(f"📤 Enviando a {correo_data['email']}...")
            mail.Send()
            
            timestamp = datetime.now().strftime('%H:%M:%S')
            nombre = correo_data.get('nombre', 'Sin nombre')
            
            self.logger.info(f"✅ ENVIADO: {correo_data['email']} - {nombre}")
            
            return {
                'exitoso': True,
                'timestamp': timestamp,
                'email': correo_data['email'],
                'nombre': nombre,
                'adjuntos_agregados': adjuntos_agregados
            }
            
        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"❌ Error enviando a {correo_data.get('email', 'desconocido')}: {error_msg}")
            
            return {
                'exitoso': False,
                'error': error_msg,
                'tipo_error': self._clasificar_error(error_msg),
                'email': correo_data.get('email', 'desconocido'),
                'nombre': correo_data.get('nombre', 'Sin nombre'),
                'reintentar': 'com' not in error_msg.lower()
            }
        finally:
            # ⭐ LIMPIAR COM (opcional pero recomendado)
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _clasificar_error(self, error_msg: str) -> str:
        """Clasificar tipo de error"""
        error_lower = error_msg.lower()
        
        if 'coinitialize' in error_lower or 'com' in error_lower:
            return 'com_error'
        elif any(palabra in error_lower for palabra in ['timeout', 'network', 'connection', 'rpc']):
            return 'red'
        elif any(palabra in error_lower for palabra in ['busy', 'temporary', 'try again']):
            return 'temporal'
        elif any(palabra in error_lower for palabra in ['invalid', 'not found', 'bad recipient']):
            return 'email_invalido'
        elif any(palabra in error_lower for palabra in ['full', 'quota', 'storage']):
            return 'buzon_lleno'
        else:
            return 'desconocido'
    
    def envio_por_lotes(self, correos: List[Dict], adjuntos: List[str], 
                       callback_progreso: Callable = None, 
                       detener_callback: Callable = None) -> Dict:
        """Envío por lotes - COM INICIALIZADO CORRECTAMENTE"""
        
        # ⭐ CLAVE: Inicializar COM en el hilo de envío
        pythoncom.CoInitialize()
        
        resultados = {
            'exitosos': [],
            'fallidos': [],
            'total_procesados': 0,
            'inicio': datetime.now()
        }
        
        self.logger.info(f"🚀 INICIANDO ENVÍO DE {len(correos)} CORREOS")
        self.logger.info("="*60)
        
        try:
            # Verificar adjuntos
            if adjuntos:
                adjuntos_faltantes = [adj for adj in adjuntos if not os.path.exists(adj)]
                if adjuntos_faltantes:
                    error_msg = f"Adjuntos faltantes: {adjuntos_faltantes}"
                    self.logger.error(f"❌ {error_msg}")
                    return {'error': error_msg}
                
                self.logger.info(f"📎 {len(adjuntos)} adjuntos verificados")
            
            # Procesar cada correo
            for i, correo in enumerate(correos):
                # Verificar si detener
                if detener_callback and detener_callback():
                    self.logger.info("⏹️ Envío detenido por el usuario")
                    break
                
                # Actualizar progreso
                progreso_actual = (i / len(correos)) * 100
                mensaje_progreso = f"Enviando a {correo.get('nombre', 'Sin nombre')} ({i+1}/{len(correos)})"
                
                if callback_progreso:
                    callback_progreso(progreso_actual, mensaje_progreso)
                
                self.logger.info(f"📧 [{i+1}/{len(correos)}] Procesando: {correo.get('email', 'desconocido')}")
                
                # ⭐ ENVIAR CORREO (COM ya inicializado)
                resultado = self.enviar_correo(correo, adjuntos)
                
                if resultado['exitoso']:
                    resultados['exitosos'].append(resultado)
                    self.logger.info(f"✅ [{i+1}/{len(correos)}] ÉXITO: {resultado['email']}")
                else:
                    resultados['fallidos'].append(resultado)
                    self.logger.error(f"❌ [{i+1}/{len(correos)}] FALLO: {resultado['email']} - {resultado['error']}")
                
                resultados['total_procesados'] += 1
                
                # Pausa entre correos (excepto el último)
                if i < len(correos) - 1:
                    tiempo_pausa = 360  # 6 minutos
                    
                    if callback_progreso:
                        siguiente_nombre = correos[i+1].get('nombre', 'Sin nombre')
                        callback_progreso(
                            progreso_actual, 
                            f"Pausa de 6 minutos... Siguiente: {siguiente_nombre}"
                        )
                    
                    self.logger.info(f"⏳ Pausa de {tiempo_pausa} segundos...")
                    
                    # Pausa con verificación de detención
                    for segundo in range(tiempo_pausa):
                        if detener_callback and detener_callback():
                            self.logger.info("⏹️ Envío detenido durante pausa")
                            break
                        time.sleep(1)
                        
                        # Actualizar progreso cada 30 segundos
                        if segundo % 30 == 0 and callback_progreso:
                            tiempo_restante = tiempo_pausa - segundo
                            callback_progreso(
                                progreso_actual,
                                f"Pausa: {tiempo_restante}s restantes..."
                            )
                    
                    # Verificar nuevamente si detener
                    if detener_callback and detener_callback():
                        break
            
        finally:
            # ⭐ LIMPIAR COM al final
            try:
                pythoncom.CoUninitialize()
            except:
                pass
        
        # Finalizar
        resultados['fin'] = datetime.now()
        resultados['duracion'] = resultados['fin'] - resultados['inicio']
        
        # Log de resumen
        self.logger.info("="*60)
        self.logger.info("📊 RESUMEN FINAL:")
        self.logger.info(f"   ✅ Exitosos: {len(resultados['exitosos'])}")
        self.logger.info(f"   ❌ Fallidos: {len(resultados['fallidos'])}")
        self.logger.info(f"   📊 Total: {resultados['total_procesados']}")
        self.logger.info(f"   ⏱️ Duración: {resultados['duracion']}")
        
        if resultados['fallidos']:
            self.logger.info("❌ ERRORES PRINCIPALES:")
            tipos_error = {}
            for fallo in resultados['fallidos']:
                tipo = fallo.get('tipo_error', 'desconocido')
                tipos_error[tipo] = tipos_error.get(tipo, 0) + 1
            
            for tipo, cantidad in tipos_error.items():
                self.logger.info(f"   • {tipo}: {cantidad} casos")
        
        self.logger.info("="*60)
        
        return resultados
    
    def probar_conexion(self) -> str:
        """Probar conexión con Outlook"""
        conexion = self.conectar_outlook()
        
        if conexion['exitoso']:
            try:
                # Información detallada
                namespace = self.outlook.GetNamespace("MAPI")
                accounts = namespace.Accounts
                
                reporte = "✅ CONEXIÓN CON OUTLOOK EXITOSA\n"
                reporte += "="*50 + "\n"
                reporte += f"📧 Total cuentas: {accounts.Count}\n"
                
                if accounts.Count > 0:
                    cuenta_principal = accounts.Item(1)
                    reporte += f"📮 Cuenta principal: {cuenta_principal.DisplayName}\n"
                    
                    if hasattr(cuenta_principal, 'SmtpAddress') and cuenta_principal.SmtpAddress:
                        reporte += f"📧 Email: {cuenta_principal.SmtpAddress}\n"
                
                # Verificar carpetas
                try:
                    inbox = namespace.GetDefaultFolder(6)
                    sent = namespace.GetDefaultFolder(5)
                    reporte += f"📥 Bandeja entrada: Accesible\n"
                    reporte += f"📤 Elementos enviados: Accesible\n"
                except:
                    reporte += f"⚠️ Algunas carpetas no accesibles\n"
                
                reporte += f"\n🎯 Estado: LISTO PARA ENVIAR\n"
                reporte += f"🔧 COM: Inicializado correctamente\n"
                
                return reporte
                
            except Exception as e:
                return f"⚠️ Conectado con advertencias:\n{str(e)}"
        else:
            reporte = "❌ ERROR DE CONEXIÓN\n"
            reporte += "="*40 + "\n"
            reporte += f"🚨 Error: {conexion['mensaje']}\n"
            if 'sugerencia' in conexion:
                reporte += f"💡 Sugerencia: {conexion['sugerencia']}\n"
            
            reporte += f"\n🔧 PASOS PARA SOLUCIONAR:\n"
            reporte += f"1. Abre Microsoft Outlook manualmente\n"
            reporte += f"2. Configura al menos una cuenta de email\n"
            reporte += f"3. Verifica que Outlook funcione\n"
            reporte += f"4. Ejecuta como administrador si persiste\n"
            
            return reporte

# Función de prueba
if __name__ == "__main__":
    print("🧪 PROBANDO EMAILSENDER CON CORRECCIÓN COM...")
    
    sender = EmailSender()
    print(sender.probar_conexion())
    
    # Prueba de envío básico
    respuesta = input("\n¿Enviar correo de prueba? (s/n): ").lower().strip()
    if respuesta in ['s', 'si', 'sí', 'y', 'yes']:
        email_test = input("Tu email para prueba: ").strip()
        if '@' in email_test:
            correo_test = {
                'email': email_test,
                'nombre': 'Usuario Test',
                'asunto': '🧪 Prueba EmailSender CORREGIDO',
                'contenido': '''¡Hola!

Este es un correo de prueba del EmailSender CORREGIDO.

✅ Error COM solucionado
✅ Inicialización correcta
✅ Envío funcional

¡Funciona perfectamente!

Saludos,
EmailSender Pro'''
            }
            
            print("📤 Enviando correo de prueba...")
            resultado = sender.enviar_correo(correo_test)
            
            if resultado['exitoso']:
                print("✅ ¡Correo enviado exitosamente!")
            else:
                print(f"❌ Error: {resultado['error']}")
        else:
            print("❌ Email inválido")
    
    input("\nPresiona Enter para cerrar...")