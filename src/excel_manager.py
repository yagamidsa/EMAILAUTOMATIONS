import pandas as pd
import os
from typing import Dict, List, Optional

class ExcelManager:
    """Gestor para leer y validar archivos Excel"""
    
    def __init__(self, data_folder: str = "data"):
        self.data_folder = data_folder
        self.campanas_file = os.path.join(data_folder, "CAMPAÑAS.xlsx")
        self.clientes_file = os.path.join(data_folder, "CLIENTES.xlsx")
        self.config_file = os.path.join(data_folder, "CONFIGURACION.xlsx")
    
    def verificar_archivos(self) -> Dict[str, bool]:
        """Verifica que todos los archivos Excel existan"""
        archivos = {
            'CAMPAÑAS.xlsx': os.path.exists(self.campanas_file),
            'CLIENTES.xlsx': os.path.exists(self.clientes_file),
            'CONFIGURACION.xlsx': os.path.exists(self.config_file)
        }
        return archivos
    
    def cargar_campanas(self) -> Dict:
        """Carga las campañas desde Excel"""
        try:
            df = pd.read_excel(self.campanas_file, sheet_name='Campañas')
            
            # Buscar campaña activa
            campana_activa = None
            for _, row in df.iterrows():
                if str(row['ACTIVA']).upper() == 'SÍ':
                    campana_activa = {
                        'id': row['ID'],
                        'nombre': row['Nombre_Campaña'],
                        'asunto': row['Asunto_Email'],
                        'contenido': row['Contenido_Email']
                    }
                    break
            
            return {
                'todas': df.to_dict('records'),
                'activa': campana_activa,
                'total': len(df)
            }
        except Exception as e:
            return {'error': f"Error cargando campañas: {str(e)}"}
    
    def cargar_clientes(self) -> Dict:
        """Carga la lista de clientes desde Excel"""
        try:
            df = pd.read_excel(self.clientes_file, sheet_name='Contactos')
            
            # Limpiar datos
            clientes = []
            for _, row in df.iterrows():
                cliente = {
                    'email': str(row['Email']).strip(),
                    'nombre': str(row['Nombre']).strip() if pd.notna(row['Nombre']) else '',
                    'empresa': str(row['Empresa']).strip() if pd.notna(row['Empresa']) else '',
                    'mensaje_personal': str(row['Mensaje_Personal']).strip() if pd.notna(row['Mensaje_Personal']) else ''
                }
                
                # Solo agregar si tiene email válido
                if '@' in cliente['email']:
                    clientes.append(cliente)
            
            return {
                'clientes': clientes,
                'total': len(clientes),
                'total_con_nombre': len([c for c in clientes if c['nombre']]),
                'total_sin_nombre': len([c for c in clientes if not c['nombre']])
            }
        except Exception as e:
            return {'error': f"Error cargando clientes: {str(e)}"}
    
    def cargar_configuracion(self) -> Dict:
        """Carga la configuración desde Excel"""
        try:
            df = pd.read_excel(self.config_file, sheet_name='Config')
            
            # Convertir a diccionario
            config = {}
            for _, row in df.iterrows():
                config[row['Configuración']] = row['Valor']
            
            return {
                'config': config,
                'valida': self._validar_configuracion(config)
            }
        except Exception as e:
            return {'error': f"Error cargando configuración: {str(e)}"}
    
    def _validar_configuracion(self, config: Dict) -> bool:
        """Valida que la configuración tenga los campos necesarios"""
        campos_requeridos = [
            'Tu_Email', 'Tu_Nombre', 'Total_Correos_Por_Dia',
            'Horas_Para_Enviar_Todo', 'Correos_Por_Lote', 'Minutos_Entre_Lotes'
        ]
        
        for campo in campos_requeridos:
            if campo not in config:
                return False
        return True
    
    def obtener_resumen(self) -> str:
        """Obtiene un resumen de todos los datos"""
        archivos = self.verificar_archivos()
        
        resumen = "📊 ESTADO DE ARCHIVOS EXCEL:\n"
        resumen += "=" * 40 + "\n"
        
        for archivo, existe in archivos.items():
            estado = "✅" if existe else "❌"
            resumen += f"{estado} {archivo}\n"
        
        if all(archivos.values()):
            # Cargar datos
            campanas = self.cargar_campanas()
            clientes = self.cargar_clientes()
            config = self.cargar_configuracion()
            
            resumen += "\n📋 RESUMEN DE DATOS:\n"
            resumen += "-" * 20 + "\n"
            
            if 'error' not in campanas:
                if campanas['activa']:
                    resumen += f"🎯 Campaña activa: {campanas['activa']['nombre']}\n"
                    resumen += f"📧 Asunto: {campanas['activa']['asunto'][:50]}...\n"
                else:
                    resumen += "⚠️  No hay campaña activa (marca 'SÍ' en alguna)\n"
                resumen += f"📊 Total campañas: {campanas['total']}\n"
            
            if 'error' not in clientes:
                resumen += f"👥 Total clientes: {clientes['total']}\n"
                resumen += f"   ├─ Con nombre: {clientes['total_con_nombre']}\n"
                resumen += f"   └─ Sin nombre: {clientes['total_sin_nombre']}\n"
            
            if 'error' not in config:
                cfg = config['config']
                resumen += f"⚙️  Configuración: {'✅ Válida' if config['valida'] else '❌ Inválida'}\n"
                resumen += f"   ├─ Email: {cfg.get('Tu_Email', 'No configurado')}\n"
                resumen += f"   ├─ Total correos: {cfg.get('Total_Correos_Por_Dia', 'No configurado')}\n"
                resumen += f"   └─ Duración: {cfg.get('Horas_Para_Enviar_Todo', 'No configurado')} horas\n"
        
        return resumen

# Función de prueba
if __name__ == "__main__":
    print("🧪 Probando ExcelManager...")
    
    excel_mgr = ExcelManager()
    print(excel_mgr.obtener_resumen())