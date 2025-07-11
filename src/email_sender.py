import win32com.client
import time
import os
from typing import List, Dict, Optional
import logging
from datetime import datetime

class EmailSender:
    """Envía correos usando Microsoft Outlook"""
    
    def __init__(self):
        self.outlook = None
        self.conectado = False
        self.logger = self._configurar_logger()
        
    def _configurar_logger(self):
        """Configura el sistema de logging"""
        logger = logging.getLogger('EmailSender')
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            # Log a archivo
            os.makedirs('reportes', exist_ok=True)
            fecha = datetime.now().strftime('%Y-%m-%d')
            archivo_log = f'reportes/envios_{fecha}.log'
            
            file_handler = logging.FileHandler(archivo_log, encoding='utf-8')
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            
        return logger
    
    def conectar_outlook(self) -> Dict:
        """Conecta con Microsoft Outlook"""
        try:
            self.logger.info("Intentando conectar con Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            
            # Verificar que Outlook esté funcionando
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            
            self.conectado = True
            self.logger.info("✅ Conexión con Outlook exitosa")
            
            return {
                'exitoso': True,
                'mensaje': 'Conectado a Outlook correctamente'
            }
            
        except Exception as e:
            self.conectado = False
            error_msg = f"Error conectando con Outlook: {str(e)}"
            self.logger.error(error_msg)
            
            return {
                'exitoso': False,
                'mensaje': error_msg,
                'sugerencia': 'Asegúrate de que Outlook esté instalado y funcionando'
            }
    
    def verificar_adjuntos(self, rutas_adjuntos: List[str]) -> Dict:
        """Verifica que todos los archivos adjuntos existan"""
        archivos_validos = []
        archivos_faltantes = []
        tamaño_total = 0
        
        for ruta in rutas_adjuntos:
            if os.path.exists(ruta):
                tamaño = os.path.getsize(ruta)
                archivos_validos.append({
                    'ruta': ruta,
                    'nombre': os.path.basename(ruta),
                    'tamaño': tamaño
                })
                tamaño_total += tamaño
            else:
                archivos_faltantes.append(ruta)
        
        return {
            'validos': archivos_validos,
            'faltantes': archivos_faltantes,
            'tamaño_total': tamaño_total,
            'total_archivos': len(archivos_validos)
        }
    
    def enviar_correo(self, correo_data: Dict, adjuntos: List[str] = None) -> Dict:
        """Envía un correo individual"""
        if not self.conectado:
            return {
                'exitoso': False,
                'error': 'No hay conexión con Outlook'
            }
        
        try:
            # Crear nuevo correo
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            
            # Configurar correo
            mail.To = correo_data['email']
            mail.Subject = correo_data['asunto']
            mail.Body = correo_data['contenido']
            
            # Agregar adjuntos
            if adjuntos:
                for ruta_adjunto in adjuntos:
                    if os.path.exists(ruta_adjunto):
                        mail.Attachments.Add(ruta_adjunto)
                        self.logger.info(f"Adjunto agregado: {os.path.basename(ruta_adjunto)}")
            
            # Enviar correo
            mail.Send()
            
            timestamp = datetime.now().strftime('%H:%M:%S')
            self.logger.info(f"✅ Correo enviado a {correo_data['email']} - {correo_data['nombre']}")
            
            return {
                'exitoso': True,
                'timestamp': timestamp,
                'email': correo_data['email'],
                'nombre': correo_data['nombre']
            }
            
        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"❌ Error enviando a {correo_data['email']}: {error_msg}")
            
            # Clasificar tipos de error
            tipo_error = self._clasificar_error(error_msg)
            
            return {
                'exitoso': False,
                'error': error_msg,
                'tipo_error': tipo_error,
                'email': correo_data['email'],
                'nombre': correo_data['nombre'],
                'reintentar': tipo_error in ['temporal', 'red']
            }
    
    def _clasificar_error(self, error_msg: str) -> str:
        """Clasifica el tipo de error para determinar si reintentar"""
        error_lower = error_msg.lower()
        
        if any(palabra in error_lower for palabra in ['timeout', 'network', 'connection']):
            return 'red'
        elif any(palabra in error_lower for palabra in ['busy', 'temporary', 'try again']):
            return 'temporal'
        elif any(palabra in error_lower for palabra in ['invalid', 'not found', 'does not exist']):
            return 'email_invalido'
        elif any(palabra in error_lower for palabra in ['full', 'quota', 'storage']):
            return 'buzon_lleno'
        else:
            return 'desconocido'
    
    def envio_por_lotes(self, correos: List[Dict], adjuntos: List[str], 
                       callback_progreso=None, detener_callback=None) -> Dict:
        """Envía correos en lotes con control de velocidad"""
        
        resultados = {
            'exitosos': [],
            'fallidos': [],
            'total_procesados': 0,
            'inicio': datetime.now()
        }
        
        self.logger.info(f"Iniciando envío de {len(correos)} correos")
        
        # Verificar adjuntos antes de empezar
        verificacion_adjuntos = self.verificar_adjuntos(adjuntos)
        if verificacion_adjuntos['faltantes']:
            return {
                'error': f"Archivos adjuntos faltantes: {verificacion_adjuntos['faltantes']}"
            }
        
        for i, correo in enumerate(correos):
            # Verificar si se debe detener
            if detener_callback and detener_callback():
                self.logger.info("Envío detenido por el usuario")
                break
            
            # Actualizar progreso
            if callback_progreso:
                progreso = (i / len(correos)) * 100
                callback_progreso(progreso, f"Enviando a {correo['nombre']} ({i+1}/{len(correos)})")
            
            # Enviar correo
            resultado = self.enviar_correo(correo, adjuntos)
            
            if resultado['exitoso']:
                resultados['exitosos'].append(resultado)
            else:
                resultados['fallidos'].append(resultado)
            
            resultados['total_procesados'] += 1
            
            # Pausa entre correos (6 minutos por defecto)
            if i < len(correos) - 1:  # No pausar después del último
                if callback_progreso:
                    callback_progreso(
                        (i / len(correos)) * 100, 
                        f"Pausa de 6 minutos... (Siguiente: {correos[i+1]['nombre']})"
                    )
                
                # Pausa de 6 minutos = 360 segundos
                for segundo in range(360):
                    if detener_callback and detener_callback():
                        self.logger.info("Envío detenido durante pausa")
                        return resultados
                    time.sleep(1)
        
        # Finalizar
        resultados['fin'] = datetime.now()
        resultados['duracion'] = resultados['fin'] - resultados['inicio']
        
        self.logger.info(f"Envío completado: {len(resultados['exitosos'])} exitosos, {len(resultados['fallidos'])} fallidos")
        
        return resultados
    
    def probar_conexion(self) -> str:
        """Prueba la conexión con Outlook y retorna un reporte"""
        conexion = self.conectar_outlook()
        
        if conexion['exitoso']:
            try:
                # Obtener información de la cuenta
                namespace = self.outlook.GetNamespace("MAPI")
                accounts = namespace.Accounts
                
                reporte = "✅ CONEXIÓN CON OUTLOOK EXITOSA\n"
                reporte += "=" * 40 + "\n"
                reporte += f"📧 Cuentas disponibles: {accounts.Count}\n"
                
                if accounts.Count > 0:
                    cuenta_principal = accounts.Item(1)
                    reporte += f"📮 Cuenta principal: {cuenta_principal.DisplayName}\n"
                
                return reporte
                
            except Exception as e:
                return f"⚠️ Conectado pero con advertencias: {str(e)}"
        else:
            reporte = "❌ ERROR DE CONEXIÓN\n"
            reporte += "=" * 30 + "\n"
            reporte += f"Error: {conexion['mensaje']}\n"
            if 'sugerencia' in conexion:
                reporte += f"Sugerencia: {conexion['sugerencia']}\n"
            return reporte

# Función de prueba
if __name__ == "__main__":
    print("🧪 Probando EmailSender...")
    
    sender = EmailSender()
    print(sender.probar_conexion())