import win32com.client
import pythoncom
import time
import os
from typing import List, Dict, Optional, Callable
import logging
from datetime import datetime, timedelta
import math
import random
import json
import csv

class SmartEmailSender:
    """EmailSender INTELIGENTE que LEE configuración del EXCEL - PARTE 1"""
    
    def __init__(self):
        self.outlook = None
        self.conectado = False
        self.logger = self._configurar_logger()
        
        # ⭐ CONFIGURACIÓN POR DEFECTO (si Excel no está disponible)
        self.config_default = {
            'MAX_CORREOS_DIARIOS': 400,
            'HORAS_TRABAJO': 8,
            'CORREOS_POR_LOTE': 50,
            'MINUTOS_ENTRE_LOTES': 6,
            'EMPEZAR_INMEDIATAMENTE': True
        }
        
        # Variables actuales (se cargan del Excel)
        self.config_actual = self.config_default.copy()
        self.CORREOS_POR_HORA = self.config_actual['MAX_CORREOS_DIARIOS'] // self.config_actual['HORAS_TRABAJO']
        
        # LÍMITES ANTI-SPAM DINÁMICOS
        self.LIMITE_RAPIDO = 25  # Base mínima
        self.PAUSA_CORTA = 30    # 30 segundos entre correos normales
        self.PAUSA_LARGA = self.config_actual['MINUTOS_ENTRE_LOTES'] * 60  # En segundos
        self.PAUSA_ALEATORIA = (15, 60)  # Variación aleatoria
        
        # Para reportes
        self.reportes_folder = "reportes"
        self.ultimo_asunto = ""
        self.ultimo_contenido = ""
        os.makedirs(self.reportes_folder, exist_ok=True)
        
        print(f"📊 CONFIGURACIÓN INICIAL (Default):")
        self.mostrar_configuracion_actual()
        
    def cargar_configuracion_excel(self, config_excel: Dict = None) -> bool:
        """⭐ CARGAR configuración desde CONFIGURACION.xlsx"""
        
        print("⚙️ CARGANDO CONFIGURACIÓN DESDE EXCEL...")
        
        if not config_excel or 'error' in config_excel:
            print("⚠️ No hay configuración Excel válida - usando valores por defecto")
            return False
        
        try:
            config_data = config_excel.get('config', {})
            
            # Mapear campos del Excel a configuración interna
            campos_excel = {
                'Total_Correos_Por_Dia': 'MAX_CORREOS_DIARIOS',
                'Horas_Para_Enviar_Todo': 'HORAS_TRABAJO', 
                'Correos_Por_Lote': 'CORREOS_POR_LOTE',
                'Minutos_Entre_Lotes': 'MINUTOS_ENTRE_LOTES',
                'Empezar_Inmediatamente': 'EMPEZAR_INMEDIATAMENTE'
            }
            
            # Cargar valores del Excel
            cambios = []
            for campo_excel, campo_interno in campos_excel.items():
                if campo_excel in config_data:
                    valor_excel = config_data[campo_excel]
                    
                    # Procesar según tipo
                    if campo_interno == 'EMPEZAR_INMEDIATAMENTE':
                        # Convertir SÍ/NO a boolean
                        nuevo_valor = str(valor_excel).upper() in ['SÍ', 'SI', 'YES', 'TRUE', '1']
                    else:
                        # Convertir a número
                        try:
                            nuevo_valor = int(float(valor_excel))
                        except:
                            print(f"⚠️ Valor inválido para {campo_excel}: {valor_excel}")
                            continue
                    
                    # Validar rangos
                    if campo_interno == 'MAX_CORREOS_DIARIOS' and not (1 <= nuevo_valor <= 1000):
                        print(f"⚠️ {campo_excel} fuera de rango (1-1000): {nuevo_valor}")
                        continue
                    elif campo_interno == 'HORAS_TRABAJO' and not (1 <= nuevo_valor <= 24):
                        print(f"⚠️ {campo_excel} fuera de rango (1-24): {nuevo_valor}")
                        continue
                    elif campo_interno == 'CORREOS_POR_LOTE' and not (1 <= nuevo_valor <= 100):
                        print(f"⚠️ {campo_excel} fuera de rango (1-100): {nuevo_valor}")
                        continue
                    elif campo_interno == 'MINUTOS_ENTRE_LOTES' and not (1 <= nuevo_valor <= 60):
                        print(f"⚠️ {campo_excel} fuera de rango (1-60): {nuevo_valor}")
                        continue
                    
                    # Aplicar cambio
                    valor_anterior = self.config_actual[campo_interno]
                    self.config_actual[campo_interno] = nuevo_valor
                    
                    if valor_anterior != nuevo_valor:
                        cambios.append(f"   • {campo_excel}: {valor_anterior} → {nuevo_valor}")
            
            # Recalcular valores derivados
            self.CORREOS_POR_HORA = self.config_actual['MAX_CORREOS_DIARIOS'] // self.config_actual['HORAS_TRABAJO']
            self.PAUSA_LARGA = self.config_actual['MINUTOS_ENTRE_LOTES'] * 60
            
            # Ajustar límite rápido según configuración
            self.LIMITE_RAPIDO = min(25, self.config_actual['CORREOS_POR_LOTE'])
            
            print(f"✅ CONFIGURACIÓN EXCEL CARGADA")
            if cambios:
                print(f"🔄 CAMBIOS APLICADOS:")
                for cambio in cambios:
                    print(cambio)
            else:
                print(f"📊 Sin cambios - valores coinciden")
            
            self.logger.info(f"✅ Configuración Excel cargada: {len(cambios)} cambios")
            return True
            
        except Exception as e:
            print(f"❌ Error cargando configuración Excel: {e}")
            self.logger.error(f"❌ Error configuración Excel: {e}")
            return False
    
    def mostrar_configuracion_actual(self):
        """Mostrar configuración actual"""
        print(f"📊 CONFIGURACIÓN ACTUAL:")
        print(f"   📈 Max correos diarios: {self.config_actual['MAX_CORREOS_DIARIOS']}")
        print(f"   ⏰ Horas de trabajo: {self.config_actual['HORAS_TRABAJO']}")
        print(f"   📧 Correos por hora: {self.CORREOS_POR_HORA}")
        print(f"   📦 Correos por lote: {self.config_actual['CORREOS_POR_LOTE']}")
        print(f"   ⏳ Minutos entre lotes: {self.config_actual['MINUTOS_ENTRE_LOTES']}")
        print(f"   🚀 Empezar inmediatamente: {self.config_actual['EMPEZAR_INMEDIATAMENTE']}")
        print(f"   🛡️ Límite rápido: {self.LIMITE_RAPIDO}")
        print(f"   📁 Reportes: {self.reportes_folder}")
        
    def _configurar_logger(self):
        """Configurar logging"""
        logger = logging.getLogger('SmartEmailSender')
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            os.makedirs('reportes', exist_ok=True)
            fecha = datetime.now().strftime('%Y-%m-%d')
            archivo_log = f'reportes/envios_inteligente_{fecha}.log'
            
            file_handler = logging.FileHandler(archivo_log, encoding='utf-8')
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            
        return logger
    
    def conectar_outlook(self) -> Dict:
        """Conectar con Outlook"""
        try:
            self.logger.info("🔄 Conectando con Outlook...")
            pythoncom.CoInitialize()
            
            try:
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                self.logger.info("✅ Conectado a instancia existente")
            except:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.logger.info("✅ Nueva instancia creada")
            
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            
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
    
    def calcular_estrategia_envio(self, total_correos: int, config_excel: Dict = None) -> Dict:
        """⭐ CALCULAR ESTRATEGIA usando configuración del EXCEL"""
        
        # Cargar configuración Excel si se proporciona
        if config_excel:
            self.cargar_configuracion_excel(config_excel)
        
        print(f"\n🎯 CALCULANDO ESTRATEGIA para {total_correos} correos...")
        print(f"📊 Usando configuración:")
        print(f"   • Max diarios: {self.config_actual['MAX_CORREOS_DIARIOS']}")
        print(f"   • Horas trabajo: {self.config_actual['HORAS_TRABAJO']}")
        print(f"   • Lote máximo: {self.config_actual['CORREOS_POR_LOTE']}")
        print(f"   • Pausa entre lotes: {self.config_actual['MINUTOS_ENTRE_LOTES']}m")
        
        estrategia = {
            'total_correos': total_correos,
            'modo': '',
            'descripcion': '',
            'tiempo_estimado': '',
            'lotes': [],
            'pausas': {},
            'config_usada': self.config_actual.copy()
        }
        
        # VERIFICAR LÍMITE DIARIO
        if total_correos > self.config_actual['MAX_CORREOS_DIARIOS']:
            print(f"⚠️ ADVERTENCIA: {total_correos} correos superan el límite diario de {self.config_actual['MAX_CORREOS_DIARIOS']}")
            print(f"📊 Se procesarán solo los primeros {self.config_actual['MAX_CORREOS_DIARIOS']}")
            total_correos = self.config_actual['MAX_CORREOS_DIARIOS']
            estrategia['total_correos'] = total_correos
            estrategia['advertencia'] = f'Limitado a {self.config_actual["MAX_CORREOS_DIARIOS"]} correos'
        
        if total_correos <= 2:
            # MODO INMEDIATO: Muy pocos correos
            estrategia.update({
                'modo': 'INMEDIATO',
                'descripcion': 'Envío inmediato sin pausas',
                'tiempo_estimado': '< 1 minuto',
                'lotes': [{'cantidad': total_correos, 'pausa_despues': 0}],
                'pausas': {'entre_correos': 5, 'entre_lotes': 0}
            })
            
        elif total_correos <= self.LIMITE_RAPIDO:
            # MODO RÁPIDO: Hasta el límite rápido con pausas cortas
            tiempo_total = total_correos * self.PAUSA_CORTA
            estrategia.update({
                'modo': 'RÁPIDO',
                'descripcion': f'Envío con pausas cortas de {self.PAUSA_CORTA}s',
                'tiempo_estimado': f'{tiempo_total // 60}m {tiempo_total % 60}s',
                'lotes': [{'cantidad': total_correos, 'pausa_despues': 0}],
                'pausas': {'entre_correos': self.PAUSA_CORTA, 'entre_lotes': 0}
            })
            
        else:
            # MODO DISTRIBUIDO: Usar configuración del Excel
            correos_por_lote = self.config_actual['CORREOS_POR_LOTE']
            lotes_necesarios = math.ceil(total_correos / correos_por_lote)
            
            # Crear lotes
            lotes = []
            correos_restantes = total_correos
            
            for i in range(lotes_necesarios):
                cantidad_lote = min(correos_por_lote, correos_restantes)
                
                # Pausa después del lote (excepto el último)
                pausa_despues = self.PAUSA_LARGA if i < lotes_necesarios - 1 else 0
                
                lotes.append({
                    'numero': i + 1,
                    'cantidad': cantidad_lote,
                    'pausa_despues': pausa_despues
                })
                
                correos_restantes -= cantidad_lote
            
            # Tiempo estimado usando configuración
            tiempo_envios = total_correos * 15  # 15 seg promedio por envío
            tiempo_pausas = (lotes_necesarios - 1) * self.PAUSA_LARGA
            tiempo_total = tiempo_envios + tiempo_pausas
            
            horas = tiempo_total // 3600
            minutos = (tiempo_total % 3600) // 60
            
            # Verificar si cabe en las horas de trabajo
            horas_configuradas = self.config_actual['HORAS_TRABAJO']
            if horas > horas_configuradas:
                print(f"⚠️ TIEMPO EXCEDIDO: {horas}h estimadas > {horas_configuradas}h configuradas")
                estrategia['advertencia_tiempo'] = f'Excede {horas_configuradas}h configuradas'
            
            estrategia.update({
                'modo': 'DISTRIBUIDO',
                'descripcion': f'{lotes_necesarios} lotes de ~{correos_por_lote} correos (según Excel)',
                'tiempo_estimado': f'{horas}h {minutos}m' if horas > 0 else f'{minutos}m',
                'lotes': lotes,
                'pausas': {
                    'entre_correos': random.randint(*self.PAUSA_ALEATORIA),
                    'entre_lotes': self.PAUSA_LARGA
                }
            })
        
        return estrategia
    
    def mostrar_estrategia(self, estrategia: Dict) -> str:
        """⭐ MOSTRAR estrategia con datos del Excel"""
        
        resumen = f"🎯 ESTRATEGIA DE ENVÍO (Configuración Excel)\n"
        resumen += f"=" * 55 + "\n\n"
        
        resumen += f"📊 Total correos: {estrategia['total_correos']}\n"
        resumen += f"🚀 Modo: {estrategia['modo']}\n"
        resumen += f"📝 Descripción: {estrategia['descripcion']}\n"
        resumen += f"⏰ Tiempo estimado: {estrategia['tiempo_estimado']}\n"
        
        # Mostrar advertencias
        if 'advertencia' in estrategia:
            resumen += f"⚠️ Advertencia: {estrategia['advertencia']}\n"
        if 'advertencia_tiempo' in estrategia:
            resumen += f"⚠️ Tiempo: {estrategia['advertencia_tiempo']}\n"
        
        resumen += f"\n⚙️ CONFIGURACIÓN EXCEL USADA:\n"
        config = estrategia.get('config_usada', {})
        resumen += f"   • Max diarios: {config.get('MAX_CORREOS_DIARIOS', 'N/A')}\n"
        resumen += f"   • Horas trabajo: {config.get('HORAS_TRABAJO', 'N/A')}\n"
        resumen += f"   • Correos por lote: {config.get('CORREOS_POR_LOTE', 'N/A')}\n"
        resumen += f"   • Minutos entre lotes: {config.get('MINUTOS_ENTRE_LOTES', 'N/A')}\n"
        
        resumen += f"\n"
        
        if estrategia['modo'] == 'INMEDIATO':
            resumen += f"✅ Envío inmediato sin esperas\n"
            resumen += f"🚀 Perfecto para pocos correos\n"
            
        elif estrategia['modo'] == 'RÁPIDO':
            resumen += f"⚡ Pausa entre correos: {estrategia['pausas']['entre_correos']}s\n"
            resumen += f"✅ Sin riesgo de spam\n"
            
        else:  # DISTRIBUIDO
            resumen += f"📦 LOTES PROGRAMADOS (según Excel):\n"
            for lote in estrategia['lotes']:
                resumen += f"   • Lote {lote['numero']}: {lote['cantidad']} correos"
                if lote['pausa_despues'] > 0:
                    resumen += f" → Pausa {lote['pausa_despues'] // 60}m"
                resumen += f"\n"
            
            resumen += f"\n⏱️ PAUSAS:\n"
            resumen += f"   • Entre correos: {estrategia['pausas']['entre_correos']}s (aleatorio)\n"
            resumen += f"   • Entre lotes: {estrategia['pausas']['entre_lotes'] // 60}m (Excel)\n"
        
        resumen += f"\n🛡️ PROTECCIÓN ANTI-SPAM ACTIVADA\n"
        resumen += f"📄 Reportes automáticos en: {self.reportes_folder}/\n"
        
        return resumen
    
    def _generar_csv_completo(self, resultados: Dict, nombre_base: str) -> str:
            """Generar CSV con TODOS los datos (exitosos y fallidos) + configuración Excel"""
    
            archivo_csv = os.path.join(self.reportes_folder, f"{nombre_base}_COMPLETO.csv")
    
            with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
    
                # Headers con información de configuración Excel
                writer.writerow(['Email', 'Nombre', 'Empresa', 'Estado', 'Hora_Envio', 'Error', 'Adjuntos', 'Config_Excel_Usada'])
    
                # Configuración Excel como string para incluir en cada fila
                config_excel = resultados.get('config_excel_usada', {})
                config_str = f"Lote:{config_excel.get('CORREOS_POR_LOTE', 'N/A')},Pausa:{config_excel.get('MINUTOS_ENTRE_LOTES', 'N/A')}m"
    
                # Exitosos
                for exitoso in resultados.get('exitosos', []):
                    writer.writerow([
                        exitoso.get('email', ''),
                        exitoso.get('nombre', ''),
                        exitoso.get('empresa', ''),
                        'EXITOSO',
                        exitoso.get('timestamp', ''),
                        '',
                        exitoso.get('adjuntos_agregados', 0),
                        config_str
                    ])
    
                # Fallidos
                for fallido in resultados.get('fallidos', []):
                    writer.writerow([
                        fallido.get('email', ''),
                        fallido.get('nombre', ''),
                        fallido.get('empresa', ''),
                        'FALLIDO',
                        '',
                        fallido.get('error', ''),
                        0,
                        config_str
                    ])
    
            return archivo_csv
    
    def _generar_csv_exitosos(self, resultados: Dict, nombre_base: str) -> str:
        """Generar CSV solo con exitosos"""

        archivo_csv = os.path.join(self.reportes_folder, f"{nombre_base}_EXITOSOS.csv")

        with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)

            # Headers compatibles con Excel de clientes
            writer.writerow(['Email', 'Nombre', 'Empresa', 'Mensaje_Personal', 'Hora_Envio', 'Adjuntos'])

            for exitoso in resultados.get('exitosos', []):
                writer.writerow([
                    exitoso.get('email', ''),
                    exitoso.get('nombre', ''),
                    exitoso.get('empresa', ''),
                    '',  # Mensaje personal vacío
                    exitoso.get('timestamp', ''),
                    exitoso.get('adjuntos_agregados', 0)
                ])

        return archivo_csv

    def _generar_csv_fallidos_simple(self, resultados: Dict, nombre_base: str) -> str:
        """Generar CSV solo con correos fallidos para reintento"""

        archivo_csv = os.path.join(self.reportes_folder, f"{nombre_base}_FALLIDOS.csv")

        if not resultados.get('fallidos'):
            # Crear archivo vacío si no hay fallidos
            with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Email', 'Nombre', 'Empresa', 'Mensaje_Personal', 'Error'])
            return archivo_csv

        with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)

            # Headers compatibles con CLIENTES.xlsx
            writer.writerow(['Email', 'Nombre', 'Empresa', 'Mensaje_Personal', 'Error'])

            for fallido in resultados['fallidos']:
                writer.writerow([
                    fallido.get('email', ''),
                    fallido.get('nombre', ''),
                    fallido.get('empresa', ''),
                    fallido.get('mensaje_personal', ''),
                    fallido.get('error', '')
                ])

        return archivo_csv
    
    def _guardar_json_completo_excel(self, resultados: Dict, nombre_base: str) -> str:
        """⭐ Guardar resultados completos en JSON CON configuración Excel"""

        archivo_json = os.path.join(self.reportes_folder, f"{nombre_base}.json")

        # Convertir datetime a string para JSON
        resultados_json = resultados.copy()

        if 'inicio' in resultados_json:
            resultados_json['inicio'] = resultados_json['inicio'].isoformat()

        if 'fin' in resultados_json:
            resultados_json['fin'] = resultados_json['fin'].isoformat()

        if 'duracion' in resultados_json:
            resultados_json['duracion'] = str(resultados_json['duracion'])

        # ⭐ Agregar metadatos de configuración Excel
        resultados_json['metadatos'] = {
            'version': 'SmartEmailSender_Excel_v1.0',
            'fecha_generacion': datetime.now().isoformat(),
            'configuracion_excel_aplicada': True,
            'config_excel_usada': resultados.get('config_excel_usada', {}),
            'archivo_origen': 'CONFIGURACION.xlsx'
        }

        with open(archivo_json, 'w', encoding='utf-8') as f:
            json.dump(resultados_json, f, indent=2, ensure_ascii=False)

        return archivo_json
    
    def reintentar_fallidos_simple(self, archivo_fallidos: str = None, adjuntos: List[str] = None, 
                                  callback_progreso: Callable = None, detener_callback: Callable = None,
                                  config_excel: Dict = None) -> Dict:  # ⭐ NUEVO PARÁMETRO
        """⭐ Reintentar correos fallidos CON configuración Excel"""

        correos_reintento = []

        if archivo_fallidos:
            # Leer CSV de fallidos manualmente
            try:
                with open(archivo_fallidos, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)

                    for row in reader:
                        correo = {
                            'email': row['Email'],
                            'nombre': row['Nombre'] if row['Nombre'] else '',
                            'empresa': row['Empresa'] if row['Empresa'] else '',
                            'asunto': self.ultimo_asunto or 'Reintento - Correo importante',
                            'contenido': self.ultimo_contenido or 'Este es un reintento del correo anterior.'
                        }
                        correos_reintento.append(correo)

            except Exception as e:
                self.logger.error(f"Error leyendo archivo de fallidos: {e}")
                return {'error': f'Error leyendo archivo: {e}'}
        else:
            # Buscar fallidos del último envío
            try:
                # Buscar el JSON más reciente
                archivos_json = [f for f in os.listdir(self.reportes_folder) if f.endswith('.json')]
                if not archivos_json:
                    return {'error': 'No se encontraron archivos de resultados previos'}

                archivo_mas_reciente = max(archivos_json, 
                                         key=lambda x: os.path.getctime(os.path.join(self.reportes_folder, x)))
                archivo_completo = os.path.join(self.reportes_folder, archivo_mas_reciente)

                with open(archivo_completo, 'r', encoding='utf-8') as f:
                    resultados_previos = json.load(f)

                fallidos = resultados_previos.get('fallidos', [])
                if not fallidos:
                    return {'error': 'No hay correos fallidos para reintentar'}

                for fallido in fallidos:
                    correo = {
                        'email': fallido['email'],
                        'nombre': fallido['nombre'],
                        'empresa': fallido.get('empresa', ''),
                        'asunto': self.ultimo_asunto or 'Reintento - Correo importante',
                        'contenido': self.ultimo_contenido or 'Este es un reintento del correo anterior.'
                    }
                    correos_reintento.append(correo)

            except Exception as e:
                return {'error': f'Error buscando fallidos previos: {e}'}

        if not correos_reintento:
            return {'error': 'No hay correos para reintentar'}

        self.logger.info(f"🔄 REINTENTANDO {len(correos_reintento)} correos fallidos CON configuración Excel")
        print(f"\n🔄 REINTENTANDO {len(correos_reintento)} CORREOS FALLIDOS")
        
        # ⭐ Aplicar configuración Excel si se proporciona
        if config_excel:
            print("⚙️ Aplicando configuración Excel para reintento...")

        # Usar envío inteligente para el reintento CON configuración Excel
        return self.envio_inteligente(correos_reintento, adjuntos or [], callback_progreso, detener_callback, config_excel)

    def mostrar_resumen_rapido(self, resultados: Dict):
        """⭐ Mostrar resumen rápido en consola CON configuración Excel"""
        
        exitosos = len(resultados.get('exitosos', []))
        fallidos = len(resultados.get('fallidos', []))
        total = exitosos + fallidos
        porcentaje = (exitosos / total * 100) if total > 0 else 0
        
        print("\n" + "="*55)
        print("📊 RESUMEN RÁPIDO DEL ENVÍO (Configuración Excel)")
        print("="*55)
        
        # ⭐ Mostrar configuración Excel usada
        config_excel = resultados.get('config_excel_usada', {})
        if config_excel:
            print(f"⚙️ CONFIGURACIÓN EXCEL APLICADA:")
            print(f"   📦 Correos por lote: {config_excel.get('CORREOS_POR_LOTE', 'N/A')}")
            print(f"   ⏳ Minutos entre lotes: {config_excel.get('MINUTOS_ENTRE_LOTES', 'N/A')}")
            print(f"   📈 Max diarios: {config_excel.get('MAX_CORREOS_DIARIOS', 'N/A')}")
            print(f"   ⏰ Horas trabajo: {config_excel.get('HORAS_TRABAJO', 'N/A')}")
            print()
        
        print(f"📊 RESULTADOS:")
        print(f"✅ Exitosos: {exitosos}")
        print(f"❌ Fallidos: {fallidos}")
        print(f"📊 Total: {total}")
        print(f"📈 Éxito: {porcentaje:.1f}%")
        
        if 'estrategia' in resultados:
            print(f"🎯 Estrategia: {resultados['estrategia'].get('modo', 'N/A')}")
        
        if fallidos > 0:
            print(f"\n🔄 PARA REINTENTAR:")
            print(f"   1. Revisa el archivo CSV de fallidos")
            print(f"   2. Corrige emails inválidos")
            print(f"   3. Ajusta CONFIGURACION.xlsx si es necesario")
            print(f"   4. Usa función reintentar_fallidos_simple()")
            
            # Mostrar algunos emails fallidos
            print(f"\n❌ EMAILS FALLIDOS:")
            for i, fallido in enumerate(resultados['fallidos'][:5], 1):
                print(f"   {i}. {fallido.get('email', 'N/A')} - {fallido.get('error', 'Error desconocido')}")
            
            if len(resultados['fallidos']) > 5:
                print(f"   ... y {len(resultados['fallidos']) - 5} más")
        
        print("="*55)
    
    def listar_reportes_disponibles(self) -> List[str]:
        """⭐ Listar todos los reportes disponibles CON configuración Excel"""
        try:
            archivos = os.listdir(self.reportes_folder)
            reportes = {
                'texto': [f for f in archivos if f.endswith('.txt') and 'reporte_envio_' in f],
                'csv_fallidos': [f for f in archivos if f.endswith('_FALLIDOS.csv')],
                'csv_completos': [f for f in archivos if f.endswith('_COMPLETO.csv')],
                'csv_exitosos': [f for f in archivos if f.endswith('_EXITOSOS.csv')],
                'json': [f for f in archivos if f.endswith('.json') and 'reporte_envio_' in f]
            }
            
            print(f"\n📁 REPORTES DISPONIBLES (Con Configuración Excel):")
            print(f"   📄 Reportes texto: {len(reportes['texto'])}")
            print(f"   📊 CSVs completos: {len(reportes['csv_completos'])}")
            print(f"   ✅ CSVs exitosos: {len(reportes['csv_exitosos'])}")
            print(f"   🔄 CSVs fallidos: {len(reportes['csv_fallidos'])}")
            print(f"   💾 JSONs completos: {len(reportes['json'])}")
            
            # Mostrar los 5 más recientes
            todos_reportes = (reportes['texto'] + reportes['csv_completos'] + 
                            reportes['csv_exitosos'] + reportes['csv_fallidos'] + reportes['json'])
            
            if todos_reportes:
                todos_reportes.sort(key=lambda x: os.path.getctime(os.path.join(self.reportes_folder, x)), reverse=True)
                print(f"\n📋 ÚLTIMOS 5 REPORTES:")
                for i, archivo in enumerate(todos_reportes[:5], 1):
                    fecha_mod = datetime.fromtimestamp(os.path.getctime(os.path.join(self.reportes_folder, archivo)))
                    tipo = "Excel" if "excel" in archivo else "Clásico"
                    print(f"   {i}. {archivo} - {fecha_mod.strftime('%Y-%m-%d %H:%M')} ({tipo})")
            
            return reportes
            
        except Exception as e:
            print(f"❌ Error listando reportes: {e}")
            return []
    
    def obtener_configuracion_de_reporte(self, archivo_json: str) -> Dict:
        """⭐ NUEVA: Obtener configuración Excel de un reporte previo"""
        try:
            ruta_completa = os.path.join(self.reportes_folder, archivo_json)
            
            if not os.path.exists(ruta_completa):
                return {'error': f'Archivo no encontrado: {archivo_json}'}
            
            with open(ruta_completa, 'r', encoding='utf-8') as f:
                datos = json.load(f)
            
            config_excel = datos.get('config_excel_usada', {})
            metadatos = datos.get('metadatos', {})
            
            if not config_excel:
                return {'error': 'No hay configuración Excel en este reporte'}
            
            print(f"📊 CONFIGURACIÓN EXCEL DEL REPORTE:")
            print(f"   📄 Archivo: {archivo_json}")
            print(f"   📅 Fecha: {metadatos.get('fecha_generacion', 'N/A')}")
            print(f"   📈 Max diarios: {config_excel.get('MAX_CORREOS_DIARIOS')}")
            print(f"   ⏰ Horas trabajo: {config_excel.get('HORAS_TRABAJO')}")
            print(f"   📦 Correos por lote: {config_excel.get('CORREOS_POR_LOTE')}")
            print(f"   ⏳ Minutos entre lotes: {config_excel.get('MINUTOS_ENTRE_LOTES')}")
            print(f"   🚀 Empezar inmediatamente: {config_excel.get('EMPEZAR_INMEDIATAMENTE')}")
            
            return {
                'config': config_excel,
                'metadatos': metadatos,
                'exitoso': True
            }
            
        except Exception as e:
            return {'error': f'Error leyendo configuración: {e}'}
    
    def aplicar_configuracion_desde_reporte(self, archivo_json: str) -> bool:
        """⭐ NUEVA: Aplicar configuración Excel desde un reporte previo"""
        try:
            config_data = self.obtener_configuracion_de_reporte(archivo_json)
            
            if 'error' in config_data:
                print(f"❌ {config_data['error']}")
                return False
            
            # Aplicar configuración
            config_excel = {'config': config_data['config']}
            exito = self.cargar_configuracion_excel(config_excel)
            
            if exito:
                print(f"✅ Configuración aplicada desde: {archivo_json}")
                print("📊 Nueva configuración activa:")
                self.mostrar_configuracion_actual()
            else:
                print(f"❌ Error aplicando configuración desde: {archivo_json}")
            
            return exito
            
        except Exception as e:
            print(f"❌ Error aplicando configuración: {e}")
            return False
    
    def comparar_configuraciones_reportes(self, archivo1: str, archivo2: str):
        """⭐ NUEVA: Comparar configuraciones Excel entre dos reportes"""
        try:
            config1 = self.obtener_configuracion_de_reporte(archivo1)
            config2 = self.obtener_configuracion_de_reporte(archivo2)
            
            if 'error' in config1 or 'error' in config2:
                print("❌ Error leyendo uno de los reportes")
                return
            
            cfg1 = config1['config']
            cfg2 = config2['config']
            
            print(f"\n📊 COMPARACIÓN DE CONFIGURACIONES:")
            print(f"=" * 50)
            print(f"📄 Reporte 1: {archivo1}")
            print(f"📄 Reporte 2: {archivo2}")
            print(f"=" * 50)
            
            campos = ['MAX_CORREOS_DIARIOS', 'HORAS_TRABAJO', 'CORREOS_POR_LOTE', 'MINUTOS_ENTRE_LOTES', 'EMPEZAR_INMEDIATAMENTE']
            
            for campo in campos:
                val1 = cfg1.get(campo, 'N/A')
                val2 = cfg2.get(campo, 'N/A')
                
                if val1 == val2:
                    estado = "="
                elif val1 > val2:
                    estado = "↑"
                else:
                    estado = "↓"
                
                print(f"{campo:25} | {val1:>8} {estado} {val2:<8}")
            
            print(f"=" * 50)
            
        except Exception as e:
            print(f"❌ Error comparando: {e}")

    def probar_conexion(self) -> str:
        """⭐ Probar conexión con Outlook MEJORADO con configuración"""
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
                
                # ⭐ Mostrar configuración actual
                reporte += f"\n⚙️ CONFIGURACIÓN ACTUAL:\n"
                reporte += f"📈 Max correos diarios: {self.config_actual['MAX_CORREOS_DIARIOS']}\n"
                reporte += f"📦 Correos por lote: {self.config_actual['CORREOS_POR_LOTE']}\n"
                reporte += f"⏳ Minutos entre lotes: {self.config_actual['MINUTOS_ENTRE_LOTES']}\n"
                reporte += f"📁 Reportes: {self.reportes_folder}/\n"
                
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

# FUNCIÓN DE PRUEBA COMPLETA CON CONFIGURACIÓN EXCEL
if __name__ == "__main__":
    print("🧪 PROBANDO SMART EMAIL SENDER CON CONFIGURACIÓN EXCEL")
    print("=" * 65)
    
    sender = SmartEmailSender()
    
    # ⭐ PRUEBA CON CONFIGURACIÓN EXCEL SIMULADA
    print("\n🔧 PROBANDO CONFIGURACIÓN EXCEL...")
    
    config_excel_test = {
        'config': {
            'Total_Correos_Por_Dia': 300,
            'Horas_Para_Enviar_Todo': 6,
            'Correos_Por_Lote': 30,
            'Minutos_Entre_Lotes': 8,
            'Empezar_Inmediatamente': 'SÍ'
        },
        'valida': True
    }
    
    print("📊 Configuración Excel de prueba:")
    for key, value in config_excel_test['config'].items():
        print(f"   {key}: {value}")
    
    # Cargar configuración
    sender.cargar_configuracion_excel(config_excel_test)
    
    # Pruebas de estrategia con configuración Excel
    casos_prueba = [2, 10, 25, 50, 100, 300]
    
    for caso in casos_prueba:
        print(f"\n📊 CASO: {caso} correos (CON configuración Excel)")
        print("-" * 50)
        
        estrategia = sender.calcular_estrategia_envio(caso, config_excel_test)
        print(sender.mostrar_estrategia(estrategia))
    
    # Listar reportes disponibles
    print(f"\n" + "="*65)
    sender.listar_reportes_disponibles()
    
    print("\n✅ Pruebas completadas")
    print("\n💡 NUEVAS FUNCIONES CON CONFIGURACIÓN EXCEL:")
    print("   📊 sender.envio_inteligente(..., config_excel) - Envío con Excel")
    print("   🔄 sender.reintentar_fallidos_simple(..., config_excel) - Reintento con Excel")
    print("   📋 sender.obtener_configuracion_de_reporte() - Ver config de reporte")
    print("   ⚙️ sender.aplicar_configuracion_desde_reporte() - Aplicar config previa")
    print("   📊 sender.comparar_configuraciones_reportes() - Comparar configs")
    print("   📄 sender.mostrar_configuracion_actual() - Ver config actual")
    
    print(f"\n🎯 INTEGRACIÓN CON GUI:")
    print("   En gui.py cambiar:")
    print("   resultados = sender.envio_inteligente(correos, adjuntos, callback, detener)")
    print("   POR:")
    print("   config = excel_mgr.cargar_configuracion()")
    print("   resultados = sender.envio_inteligente(correos, adjuntos, callback, detener, config)")
    
    input("\nPresiona Enter para cerrar...")# CONTINUACIÓN DE SmartEmailSender - PARTE 2
    
    def enviar_correo(self, correo_data: Dict, adjuntos: List[str] = None) -> Dict:
        """Enviar correo individual"""
        if not self.conectado:
            return {
                'exitoso': False,
                'error': 'No hay conexión con Outlook'
            }
        
        try:
            pythoncom.CoInitialize()
            
            # Validar datos
            if not correo_data.get('email'):
                raise ValueError("Email del destinatario requerido")
            if not correo_data.get('asunto'):
                raise ValueError("Asunto del correo requerido")
            if not correo_data.get('contenido'):
                raise ValueError("Contenido del correo requerido")
            
            # Crear correo
            mail = self.outlook.CreateItem(0)
            mail.To = correo_data['email'].strip()
            mail.Subject = correo_data['asunto'].strip()
            
            # Configurar contenido
            contenido = correo_data['contenido']
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
            
            # ENVIAR
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
                'empresa': correo_data.get('empresa', ''),
                'adjuntos_agregados': adjuntos_agregados
            }
            
        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"❌ Error enviando a {correo_data.get('email', 'desconocido')}: {error_msg}")
            
            return {
                'exitoso': False,
                'error': error_msg,
                'email': correo_data.get('email', 'desconocido'),
                'nombre': correo_data.get('nombre', 'Sin nombre'),
                'empresa': correo_data.get('empresa', '')
            }
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def envio_inteligente(self, correos: List[Dict], adjuntos: List[str], 
                         callback_progreso: Callable = None, 
                         detener_callback: Callable = None,
                         config_excel: Dict = None) -> Dict:  # ⭐ NUEVO PARÁMETRO
        """⭐ ENVÍO INTELIGENTE usando configuración del EXCEL + REPORTES"""
        
        pythoncom.CoInitialize()
        
        # ⭐ CARGAR CONFIGURACIÓN EXCEL AL INICIO
        if config_excel:
            print("\n⚙️ APLICANDO CONFIGURACIÓN DEL EXCEL...")
            self.cargar_configuracion_excel(config_excel)
            print("✅ Configuración Excel aplicada\n")
        else:
            print("⚠️ Sin configuración Excel - usando valores por defecto")
        
        # ⭐ GUARDAR PARA REINTENTOS
        if correos:
            self.ultimo_asunto = correos[0].get('asunto', '')
            self.ultimo_contenido = correos[0].get('contenido', '')
        
        # ⭐ Calcular estrategia con configuración Excel
        estrategia = self.calcular_estrategia_envio(len(correos), config_excel)
        
        self.logger.info("🎯 ENVÍO INTELIGENTE CON CONFIGURACIÓN EXCEL")
        self.logger.info("=" * 60)
        self.logger.info(f"📊 Total correos: {len(correos)}")
        self.logger.info(f"🚀 Modo: {estrategia['modo']}")
        self.logger.info(f"📝 {estrategia['descripcion']}")
        self.logger.info(f"⏰ Tiempo estimado: {estrategia['tiempo_estimado']}")
        
        # Log de configuración usada
        config_usada = estrategia.get('config_usada', {})
        self.logger.info(f"⚙️ Configuración Excel:")
        self.logger.info(f"   Max diarios: {config_usada.get('MAX_CORREOS_DIARIOS')}")
        self.logger.info(f"   Correos por lote: {config_usada.get('CORREOS_POR_LOTE')}")
        self.logger.info(f"   Minutos entre lotes: {config_usada.get('MINUTOS_ENTRE_LOTES')}")
        
        resultados = {
            'exitosos': [],
            'fallidos': [],
            'total_procesados': 0,
            'inicio': datetime.now(),
            'estrategia': estrategia,
            'config_excel_usada': config_usada
        }
        
        try:
            # Verificar adjuntos
            if adjuntos:
                adjuntos_faltantes = [adj for adj in adjuntos if not os.path.exists(adj)]
                if adjuntos_faltantes:
                    error_msg = f"Adjuntos faltantes: {adjuntos_faltantes}"
                    self.logger.error(f"❌ {error_msg}")
                    return {'error': error_msg}
                
                self.logger.info(f"📎 {len(adjuntos)} adjuntos verificados")
            
            # PROCESAR SEGÚN ESTRATEGIA (usando configuración Excel)
            correo_actual = 0
            
            if estrategia['modo'] == 'INMEDIATO':
                # Envío inmediato
                self.logger.info("🚀 MODO INMEDIATO - Sin pausas")
                
                for i, correo in enumerate(correos):
                    if detener_callback and detener_callback():
                        break
                    
                    if callback_progreso:
                        progreso = ((i + 1) / len(correos)) * 100
                        callback_progreso(progreso, f"Enviando {i+1}/{len(correos)} - {correo.get('nombre', 'Sin nombre')}")
                    
                    resultado = self.enviar_correo(correo, adjuntos)
                    self._procesar_resultado(resultado, resultados)
                    
                    # Solo pausa mínima
                    if i < len(correos) - 1:
                        time.sleep(5)
            
            elif estrategia['modo'] == 'RÁPIDO':
                # Envío rápido con pausas cortas
                self.logger.info(f"⚡ MODO RÁPIDO - Pausas de {estrategia['pausas']['entre_correos']}s")
                
                for i, correo in enumerate(correos):
                    if detener_callback and detener_callback():
                        break
                    
                    if callback_progreso:
                        progreso = ((i + 1) / len(correos)) * 100
                        callback_progreso(progreso, f"Enviando {i+1}/{len(correos)} - {correo.get('nombre', 'Sin nombre')}")
                    
                    resultado = self.enviar_correo(correo, adjuntos)
                    self._procesar_resultado(resultado, resultados)
                    
                    # Pausa corta entre correos
                    if i < len(correos) - 1:
                        self._pausa_inteligente(estrategia['pausas']['entre_correos'], 
                                              f"Pausa rápida", callback_progreso, detener_callback)
            
            else:  # MODO DISTRIBUIDO con configuración Excel
                self.logger.info(f"📦 MODO DISTRIBUIDO - {len(estrategia['lotes'])} lotes (Excel: {config_usada.get('CORREOS_POR_LOTE')} por lote)")
                
                for num_lote, lote in enumerate(estrategia['lotes']):
                    if detener_callback and detener_callback():
                        break
                    
                    self.logger.info(f"📦 LOTE {lote['numero']}/{len(estrategia['lotes'])}: {lote['cantidad']} correos")
                    
                    # Procesar correos del lote
                    fin_lote = min(correo_actual + lote['cantidad'], len(correos))
                    
                    for i in range(correo_actual, fin_lote):
                        if detener_callback and detener_callback():
                            break
                        
                        correo = correos[i]
                        
                        if callback_progreso:
                            progreso = ((i + 1) / len(correos)) * 100
                            lote_info = f"Lote {lote['numero']}/{len(estrategia['lotes'])}"
                            callback_progreso(progreso, f"{lote_info} - {correo.get('nombre', 'Sin nombre')} ({i+1}/{len(correos)})")
                        
                        resultado = self.enviar_correo(correo, adjuntos)
                        self._procesar_resultado(resultado, resultados)
                        
                        # Pausa aleatoria entre correos del mismo lote
                        if i < fin_lote - 1:
                            pausa = random.randint(*self.PAUSA_ALEATORIA)
                            self._pausa_inteligente(pausa, f"Entre correos", callback_progreso, detener_callback)
                    
                    correo_actual = fin_lote
                    
                    # Pausa larga entre lotes (usando configuración Excel)
                    if lote['pausa_despues'] > 0 and not (detener_callback and detener_callback()):
                        minutos_pausa = config_usada.get('MINUTOS_ENTRE_LOTES', 6)
                        self.logger.info(f"⏳ Pausa entre lotes: {minutos_pausa} minutos (según Excel)")
                        
                        siguiente_lote = num_lote + 2  # +2 porque empezamos desde 1
                        if siguiente_lote <= len(estrategia['lotes']):
                            mensaje_pausa = f"Pausa {minutos_pausa}m (Excel) - Siguiente: Lote {siguiente_lote}"
                        else:
                            mensaje_pausa = f"Pausa final de {minutos_pausa}m"
                        
                        self._pausa_inteligente(lote['pausa_despues'], mensaje_pausa, callback_progreso, detener_callback)
        
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
        
        # Finalizar
        resultados['fin'] = datetime.now()
        resultados['duracion'] = resultados['fin'] - resultados['inicio']
        
        self._log_resumen_final(resultados)
        
        # ⭐ GENERAR REPORTE AUTOMÁTICAMENTE con configuración Excel
        try:
            self.logger.info("📊 Generando reporte detallado con configuración Excel...")
            archivo_reporte = self.generar_reporte_detallado(resultados)
            self.logger.info(f"📄 Reporte guardado: {archivo_reporte}")
            
        except Exception as e:
            print(f"❌ Error generando reporte: {e}")
            self.logger.error(f"❌ Error reporte: {e}")
        
        return resultados
    
    def _procesar_resultado(self, resultado: Dict, resultados: Dict):
        """Procesar resultado individual"""
        if resultado['exitoso']:
            resultados['exitosos'].append(resultado)
            self.logger.info(f"✅ ÉXITO: {resultado['email']}")
        else:
            resultados['fallidos'].append(resultado)
            self.logger.error(f"❌ FALLO: {resultado['email']} - {resultado['error']}")
        
        resultados['total_procesados'] += 1
    
    def _pausa_inteligente(self, segundos: int, descripcion: str, 
                          callback_progreso: Callable, detener_callback: Callable):
        """Pausa inteligente con actualizaciones"""
        
        for segundo in range(segundos):
            if detener_callback and detener_callback():
                break
            
            # Actualizar progreso cada 10 segundos
            if segundo % 10 == 0 and callback_progreso:
                tiempo_restante = segundos - segundo
                minutos = tiempo_restante // 60
                segs = tiempo_restante % 60
                
                if minutos > 0:
                    tiempo_texto = f"{minutos}m {segs}s"
                else:
                    tiempo_texto = f"{segs}s"
                
                callback_progreso(None, f"{descripcion} - Restante: {tiempo_texto}")
            
            time.sleep(1)
    
    def _log_resumen_final(self, resultados: Dict):
        """Log del resumen final con configuración Excel"""
        self.logger.info("=" * 60)
        self.logger.info("📊 RESUMEN FINAL DEL ENVÍO INTELIGENTE (Excel)")
        self.logger.info("=" * 60)
        self.logger.info(f"🎯 Estrategia: {resultados['estrategia']['modo']}")
        
        # Log configuración Excel usada
        config_excel = resultados.get('config_excel_usada', {})
        if config_excel:
            self.logger.info(f"⚙️ Configuración Excel:")
            self.logger.info(f"   Max diarios: {config_excel.get('MAX_CORREOS_DIARIOS')}")
            self.logger.info(f"   Correos por lote: {config_excel.get('CORREOS_POR_LOTE')}")
            self.logger.info(f"   Minutos entre lotes: {config_excel.get('MINUTOS_ENTRE_LOTES')}")
        
        self.logger.info(f"✅ Exitosos: {len(resultados['exitosos'])}")
        self.logger.info(f"❌ Fallidos: {len(resultados['fallidos'])}")
        self.logger.info(f"📊 Total: {resultados['total_procesados']}")
        self.logger.info(f"⏱️ Duración: {resultados['duracion']}")
        
        if resultados['fallidos']:
            self.logger.info("❌ ERRORES:")
            for fallo in resultados['fallidos'][:3]:
                self.logger.info(f"   • {fallo['email']}: {fallo['error']}")
            
            if len(resultados['fallidos']) > 3:
                self.logger.info(f"   ... y {len(resultados['fallidos']) - 3} más")
        
        self.logger.info("=" * 60)

    # ⭐ FUNCIONES DE REPORTE MEJORADAS CON CONFIGURACIÓN EXCEL
    
    def generar_reporte_detallado(self, resultados: Dict) -> str:
        """⭐ Generar reporte completo CON configuración Excel"""

        fecha_actual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        nombre_reporte = f"reporte_envio_excel_{fecha_actual}"

        # Estadísticas básicas
        total = len(resultados.get('exitosos', [])) + len(resultados.get('fallidos', []))
        exitosos = len(resultados.get('exitosos', []))
        fallidos = len(resultados.get('fallidos', []))
        porcentaje_exito = (exitosos / total * 100) if total > 0 else 0

        # 1. REPORTE DE TEXTO DETALLADO con configuración Excel
        reporte_texto = self._generar_reporte_texto_excel(resultados, exitosos, fallidos, porcentaje_exito)

        # 2. CSV COMPLETO (exitosos + fallidos)
        archivo_csv_completo = self._generar_csv_completo(resultados, nombre_reporte)

        # 3. CSV DE FALLIDOS PARA REINTENTO
        archivo_fallidos = self._generar_csv_fallidos_simple(resultados, nombre_reporte)

        # 4. CSV DE EXITOSOS 
        archivo_exitosos = self._generar_csv_exitosos(resultados, nombre_reporte)

        # 5. JSON COMPLETO con configuración Excel
        archivo_json = self._guardar_json_completo_excel(resultados, nombre_reporte)

        # Guardar reporte texto
        archivo_texto = os.path.join(self.reportes_folder, f"{nombre_reporte}.txt")
        with open(archivo_texto, 'w', encoding='utf-8') as f:
            f.write(reporte_texto)

        # Mostrar resumen rápido
        self.mostrar_resumen_rapido(resultados)

        print(f"\n📊 REPORTES GENERADOS CON CONFIGURACIÓN EXCEL:")
        print(f"   📄 Reporte completo: {archivo_texto}")
        print(f"   📊 CSV completo: {archivo_csv_completo}")
        print(f"   ✅ CSV exitosos: {archivo_exitosos}")
        print(f"   🔄 CSV fallidos: {archivo_fallidos}")
        print(f"   💾 JSON completo: {archivo_json}")

        return archivo_texto
    
    def _generar_reporte_texto_excel(self, resultados: Dict, exitosos: int, fallidos: int, porcentaje_exito: float) -> str:
        """⭐ Generar reporte en texto con configuración Excel"""

        reporte = "📧 REPORTE COMPLETO - ENVÍO INTELIGENTE CON EXCEL\n"
        reporte += "=" * 65 + "\n\n"

        # RESUMEN EJECUTIVO
        reporte += "📊 RESUMEN EJECUTIVO\n"
        reporte += "-" * 30 + "\n"
        reporte += f"📅 Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"

        # ⭐ CONFIGURACIÓN EXCEL USADA
        config_excel = resultados.get('config_excel_usada', {})
        if config_excel:
            reporte += f"\n⚙️ CONFIGURACIÓN EXCEL APLICADA:\n"
            reporte += f"   📈 Max correos diarios: {config_excel.get('MAX_CORREOS_DIARIOS', 'N/A')}\n"
            reporte += f"   ⏰ Horas de trabajo: {config_excel.get('HORAS_TRABAJO', 'N/A')}\n"
            reporte += f"   📦 Correos por lote: {config_excel.get('CORREOS_POR_LOTE', 'N/A')}\n"
            reporte += f"   ⏳ Minutos entre lotes: {config_excel.get('MINUTOS_ENTRE_LOTES', 'N/A')}\n"
            reporte += f"   🚀 Empezar inmediatamente: {config_excel.get('EMPEZAR_INMEDIATAMENTE', 'N/A')}\n"

        if 'estrategia' in resultados:
            estrategia = resultados['estrategia']
            reporte += f"\n🎯 ESTRATEGIA CALCULADA:\n"
            reporte += f"   🚀 Modo: {estrategia.get('modo', 'N/A')}\n"
            reporte += f"   📝 Descripción: {estrategia.get('descripcion', 'N/A')}\n"
            reporte += f"   ⏰ Tiempo estimado: {estrategia.get('tiempo_estimado', 'N/A')}\n"
            
            # Advertencias de la estrategia
            if 'advertencia' in estrategia:
                reporte += f"   ⚠️ Advertencia: {estrategia['advertencia']}\n"
            if 'advertencia_tiempo' in estrategia:
                reporte += f"   ⚠️ Tiempo: {estrategia['advertencia_tiempo']}\n"

        reporte += f"\n📊 RESULTADOS:\n"
        reporte += f"   ✅ Exitosos: {exitosos}\n"
        reporte += f"   ❌ Fallidos: {fallidos}\n"
        reporte += f"   📈 Total procesados: {exitosos + fallidos}\n"
        reporte += f"   📊 Porcentaje de éxito: {porcentaje_exito:.1f}%\n"

        if 'duracion' in resultados:
            reporte += f"   ⏱️ Duración total: {resultados['duracion']}\n"

        # ANÁLISIS CON CONFIGURACIÓN EXCEL
        reporte += f"\n📈 ANÁLISIS DE RESULTADOS\n"
        reporte += "-" * 30 + "\n"

        if porcentaje_exito >= 95:
            reporte += "🎉 EXCELENTE: Envío muy exitoso\n"
            reporte += "✅ La configuración Excel funcionó perfectamente\n"
        elif porcentaje_exito >= 85:
            reporte += "✅ BUENO: Envío exitoso con pocos problemas\n"
            reporte += "📊 La configuración Excel fue efectiva\n"
        elif porcentaje_exito >= 70:
            reporte += "⚠️ REGULAR: Envío con algunos problemas\n"
            reporte += "🔧 Considera ajustar la configuración Excel\n"
        else:
            reporte += "🚨 PROBLEMÁTICO: Muchos fallos\n"
            reporte += "⚙️ Revisar configuración Excel y datos de entrada\n"

        # EFECTIVIDAD DE LA CONFIGURACIÓN EXCEL
        if config_excel:
            total_procesado = exitosos + fallidos
            max_configurado = config_excel.get('MAX_CORREOS_DIARIOS', 0)
            
            reporte += f"\n⚙️ EFECTIVIDAD DE CONFIGURACIÓN EXCEL:\n"
            if total_procesado <= max_configurado:
                reporte += f"✅ Dentro del límite diario ({total_procesado}/{max_configurado})\n"
            else:
                reporte += f"⚠️ Excedió límite diario ({total_procesado}/{max_configurado})\n"
            
            lote_config = config_excel.get('CORREOS_POR_LOTE', 0)
            if lote_config > 0:
                lotes_usados = math.ceil(total_procesado / lote_config)
                reporte += f"📦 Lotes utilizados: {lotes_usados} (configuración: {lote_config} por lote)\n"

        # CORREOS EXITOSOS (primeros 10)
        if resultados.get('exitosos'):
            reporte += f"\n✅ CORREOS EXITOSOS ({len(resultados['exitosos'])} total)\n"
            reporte += "-" * 40 + "\n"

            for i, exitoso in enumerate(resultados['exitosos'][:10], 1):
                reporte += f"{i:2d}. {exitoso.get('email', 'N/A')} - {exitoso.get('nombre', 'Sin nombre')}"
                if 'timestamp' in exitoso:
                    reporte += f" ({exitoso['timestamp']})"
                if 'adjuntos_agregados' in exitoso and exitoso['adjuntos_agregados'] > 0:
                    reporte += f" - {exitoso['adjuntos_agregados']} adjuntos"
                reporte += "\n"

            if len(resultados['exitosos']) > 10:
                reporte += f"   ... y {len(resultados['exitosos']) - 10} más\n"

        # CORREOS FALLIDOS (todos)
        if resultados.get('fallidos'):
            reporte += f"\n❌ CORREOS FALLIDOS ({len(resultados['fallidos'])} total)\n"
            reporte += "-" * 40 + "\n"

            # Agrupar por tipo de error
            errores_agrupados = {}
            for fallido in resultados['fallidos']:
                error = fallido.get('error', 'Error desconocido')
                if error not in errores_agrupados:
                    errores_agrupados[error] = []
                errores_agrupados[error].append(fallido)

            for error, lista_fallidos in errores_agrupados.items():
                reporte += f"\n🔴 ERROR: {error} ({len(lista_fallidos)} casos)\n"
                for fallido in lista_fallidos:
                    reporte += f"   • {fallido.get('email', 'N/A')} - {fallido.get('nombre', 'Sin nombre')}\n"

        # RECOMENDACIONES CON CONFIGURACIÓN EXCEL
        reporte += f"\n💡 RECOMENDACIONES\n"
        reporte += "-" * 30 + "\n"

        if fallidos == 0:
            reporte += "🎯 ¡Perfecto! La configuración Excel funcionó óptimamente\n"
            reporte += "✅ No hay acciones requeridas\n"
        else:
            reporte += f"🔄 Reintentar {fallidos} correos fallidos usando el archivo CSV generado\n"
            reporte += f"📋 Abrir CSV fallidos en Excel para revisar y corregir\n"
            
            if porcentaje_exito < 85:
                reporte += f"\n🔧 AJUSTES SUGERIDOS EN CONFIGURACION.xlsx:\n"
                
                if config_excel:
                    lote_actual = config_excel.get('CORREOS_POR_LOTE', 50)
                    pausa_actual = config_excel.get('MINUTOS_ENTRE_LOTES', 6)
                    
                    if porcentaje_exito < 70:
                        nuevo_lote = max(10, lote_actual // 2)
                        nueva_pausa = min(15, pausa_actual * 2)
                        reporte += f"   • Reducir Correos_Por_Lote de {lote_actual} a {nuevo_lote}\n"
                        reporte += f"   • Aumentar Minutos_Entre_Lotes de {pausa_actual} a {nueva_pausa}\n"
                    else:
                        nuevo_lote = max(20, lote_actual - 10)
                        nueva_pausa = min(10, pausa_actual + 2)
                        reporte += f"   • Reducir Correos_Por_Lote de {lote_actual} a {nuevo_lote}\n"
                        reporte += f"   • Aumentar Minutos_Entre_Lotes de {pausa_actual} a {nueva_pausa}\n"

        # ARCHIVOS GENERADOS
        reporte += f"\n📁 ARCHIVOS GENERADOS\n"
        reporte += "-" * 30 + "\n"
        reporte += "📊 CSV completo con todos los datos y configuración Excel\n"
        reporte += "✅ CSV solo con exitosos\n"
        reporte += "🔄 CSV con correos fallidos para reintento\n"
        reporte += "💾 JSON con datos completos y configuración Excel\n"
        reporte += "📄 Este reporte en texto\n"

        reporte += f"\n🔄 CÓMO AJUSTAR CONFIGURACIÓN EXCEL:\n"
        reporte += "-" * 40 + "\n"
        reporte += "1. Abre CONFIGURACION.xlsx\n"
        reporte += "2. Ajusta valores según recomendaciones arriba\n"
        reporte += "3. Guarda el archivo\n"
        reporte += "4. Ejecuta el envío nuevamente\n"
        reporte += "5. El sistema usará automáticamente la nueva configuración\n"

        return reporte    