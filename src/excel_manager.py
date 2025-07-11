import pandas as pd
import os
from typing import Dict, List, Optional

class ExcelManager:
    """Gestor para leer y validar archivos Excel - CORREGIDO"""
    
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
        """Carga las campañas desde Excel - CORREGIDO PARA LEER ACTIVA CORRECTAMENTE"""
        try:
            df = pd.read_excel(self.campanas_file, sheet_name='Campañas')
            
            print(f"🔍 DEBUG: Total campañas encontradas: {len(df)}")
            
            # ⭐ BUSCAR CAMPAÑA ACTIVA CORRECTAMENTE
            campana_activa = None
            
            for index, row in df.iterrows():
                print(f"🔍 DEBUG: Campaña {index + 1}:")
                print(f"   ID: {row['ID']}")
                print(f"   Nombre: {row['Nombre_Campaña']}")
                print(f"   ACTIVA: '{row['ACTIVA']}' (tipo: {type(row['ACTIVA'])})")
                
                # ⭐ VERIFICACIÓN CORRECTA DE "SÍ"
                activa_valor = str(row['ACTIVA']).strip().upper()
                
                # Verificar múltiples formas de decir "SÍ"
                if activa_valor in ['SÍ', 'SI', 'SÍ', 'YES', 'Y', '1', 'TRUE']:
                    print(f"   ✅ Esta campaña está ACTIVA")
                    
                    if campana_activa is not None:
                        print(f"   ⚠️ ADVERTENCIA: Ya hay otra campaña activa. Usando la primera encontrada.")
                    else:
                        campana_activa = {
                            'id': row['ID'],
                            'nombre': row['Nombre_Campaña'],
                            'asunto': row['Asunto_Email'],
                            'contenido': row['Contenido_Email']
                        }
                        print(f"   🎯 CAMPAÑA ACTIVA SELECCIONADA: {row['Nombre_Campaña']}")
                else:
                    print(f"   ❌ Esta campaña NO está activa")
            
            # Resultado final
            if campana_activa:
                print(f"\n✅ CAMPAÑA ACTIVA FINAL: {campana_activa['nombre']}")
            else:
                print(f"\n❌ NO HAY CAMPAÑA ACTIVA")
                print(f"💡 Asegúrate de que una campaña tenga 'SÍ' en la columna ACTIVA")
            
            return {
                'todas': df.to_dict('records'),
                'activa': campana_activa,
                'total': len(df)
            }
            
        except Exception as e:
            print(f"❌ ERROR cargando campañas: {str(e)}")
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

    def mostrar_debug_campanas(self):
        """Función de debug para mostrar todas las campañas"""
        try:
            df = pd.read_excel(self.campanas_file, sheet_name='Campañas')
            
            print(f"\n🔍 DEBUG COMPLETO DE CAMPAÑAS:")
            print(f"=" * 50)
            print(f"Total filas: {len(df)}")
            print(f"Columnas: {list(df.columns)}")
            print(f"")
            
            for index, row in df.iterrows():
                print(f"FILA {index + 1}:")
                print(f"  ID: {row['ID']}")
                print(f"  Nombre: {row['Nombre_Campaña']}")
                print(f"  ACTIVA: '{row['ACTIVA']}' (tipo: {type(row['ACTIVA'])})")
                print(f"  ACTIVA repr: {repr(row['ACTIVA'])}")
                print(f"  ACTIVA bytes: {row['ACTIVA'].encode('utf-8') if isinstance(row['ACTIVA'], str) else 'No es string'}")
                print(f"")
                
        except Exception as e:
            print(f"❌ Error en debug: {e}")

# Función de prueba
if __name__ == "__main__":
    print("🧪 Probando ExcelManager CORREGIDO...")
    
    excel_mgr = ExcelManager()
    
    # Mostrar debug completo
    excel_mgr.mostrar_debug_campanas()
    
    # Probar carga normal
    print(f"\n" + "="*50)
    print(excel_mgr.obtener_resumen())