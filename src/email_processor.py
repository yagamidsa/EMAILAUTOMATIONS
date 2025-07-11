import re
from typing import Dict, List, Optional

class EmailProcessor:
    """Procesa y personaliza el contenido de los correos"""
    
    def __init__(self):
        pass
    
    def extraer_nombre_de_email(self, email: str) -> str:
        """Extrae un nombre del email si no se proporciona nombre manual"""
        try:
            # Obtener la parte antes del @
            parte_local = email.split('@')[0]
            
            # Separar por puntos, guiones o números
            partes = re.split(r'[._\-0-9]+', parte_local)
            
            # Tomar las primeras 2 partes y capitalizar
            nombre_partes = []
            for parte in partes[:2]:
                if len(parte) > 1:  # Ignorar partes muy cortas
                    nombre_partes.append(parte.capitalize())
            
            if nombre_partes:
                return ' '.join(nombre_partes)
            else:
                # Si no se puede extraer, usar la parte completa
                return parte_local.capitalize()
                
        except Exception:
            return "Friend"  # Fallback seguro
    
    def personalizar_contenido(self, plantilla: str, cliente: Dict, config: Dict) -> str:
        """Personaliza el contenido del email con los datos del cliente"""
        
        # Obtener nombre (manual o extraído)
        if cliente.get('nombre') and cliente['nombre'].strip():
            nombre = cliente['nombre'].strip()
        else:
            nombre = self.extraer_nombre_de_email(cliente['email'])
        
        # Preparar variables de reemplazo
        variables = {
            'NOMBRE': nombre,
            'EMPRESA': cliente.get('empresa', ''),
            'MENSAJE_PERSONAL': cliente.get('mensaje_personal', ''),
            'REMITENTE_NOMBRE': config.get('Tu_Nombre', 'Admin'),
            'REMITENTE_EMAIL': config.get('Tu_Email', ''),
            'EMPRESA_REMITENTE': config.get('Tu_Empresa', '')
        }
        
        # Reemplazar variables en el contenido
        contenido_personalizado = plantilla
        for variable, valor in variables.items():
            marcador = '{' + variable + '}'
            contenido_personalizado = contenido_personalizado.replace(marcador, str(valor))
        
        return contenido_personalizado
    
    def validar_email(self, email: str) -> bool:
        """Valida que un email tenga formato correcto"""
        patron = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(patron, email.strip()) is not None
    
    def procesar_lista_clientes(self, clientes: List[Dict], campana: Dict, config: Dict) -> List[Dict]:
        """Procesa toda la lista de clientes y genera los correos personalizados"""
        
        correos_procesados = []
        emails_vistos = set()  # Para detectar duplicados
        
        for i, cliente in enumerate(clientes):
            try:
                email = cliente.get('email', '').strip().lower()
                
                # Validar email
                if not self.validar_email(email):
                    print(f"⚠️  Email inválido saltado: {email}")
                    continue
                
                # Detectar duplicados
                if email in emails_vistos:
                    print(f"⚠️  Email duplicado saltado: {email}")
                    continue
                emails_vistos.add(email)
                
                # Personalizar contenido
                contenido_personalizado = self.personalizar_contenido(
                    campana['contenido'], 
                    cliente, 
                    config
                )
                
                # Obtener nombre final
                if cliente.get('nombre') and cliente['nombre'].strip():
                    nombre_final = cliente['nombre'].strip()
                else:
                    nombre_final = self.extraer_nombre_de_email(email)
                
                correo_procesado = {
                    'indice': i + 1,
                    'email': email,
                    'nombre': nombre_final,
                    'empresa': cliente.get('empresa', ''),
                    'asunto': campana['asunto'],
                    'contenido': contenido_personalizado,
                    'estado': 'pendiente'
                }
                
                correos_procesados.append(correo_procesado)
                
            except Exception as e:
                print(f"❌ Error procesando cliente {i+1}: {str(e)}")
                continue
        
        return correos_procesados
    
    def obtener_vista_previa(self, clientes: List[Dict], campana: Dict, config: Dict, limite: int = 3) -> str:
        """Genera una vista previa de los primeros correos"""
        
        correos = self.procesar_lista_clientes(clientes, campana, config)
        
        if not correos:
            return "❌ No hay correos válidos para procesar"
        
        vista_previa = f"📧 VISTA PREVIA DE CORREOS (Primeros {min(limite, len(correos))} de {len(correos)}):\n"
        vista_previa += "=" * 60 + "\n"
        
        for i, correo in enumerate(correos[:limite]):
            vista_previa += f"\n📩 CORREO #{correo['indice']}:\n"
            vista_previa += f"Para: {correo['email']}\n"
            vista_previa += f"Nombre: {correo['nombre']}\n"
            vista_previa += f"Empresa: {correo['empresa']}\n"
            vista_previa += f"Asunto: {correo['asunto'][:50]}...\n"
            vista_previa += f"Contenido (primeras líneas):\n"
            
            # Mostrar solo las primeras 3 líneas del contenido
            lineas_contenido = correo['contenido'].split('\n')[:3]
            for linea in lineas_contenido:
                if linea.strip():
                    vista_previa += f"  {linea.strip()}\n"
            vista_previa += "  ...\n"
            
            if i < min(limite, len(correos)) - 1:
                vista_previa += "-" * 40 + "\n"
        
        vista_previa += f"\n📊 RESUMEN:\n"
        vista_previa += f"   ├─ Total correos válidos: {len(correos)}\n"
        vista_previa += f"   ├─ Emails únicos: {len(set(c['email'] for c in correos))}\n"
        vista_previa += f"   └─ Campaña: {campana['nombre']}\n"
        
        return vista_previa

# Función de prueba
if __name__ == "__main__":
    print("🧪 Probando EmailProcessor...")
    
    # Datos de prueba
    processor = EmailProcessor()
    
    # Prueba de extracción de nombres
    print("\n🔤 Pruebas de extracción de nombres:")
    emails_prueba = [
        "juan.carlos@empresa.com",
        "maria123@tienda.com", 
        "admin@prueba.com",
        "test@ejemplo.com"
    ]
    
    for email in emails_prueba:
        nombre = processor.extraer_nombre_de_email(email)
        print(f"   {email} → {nombre}")
    
    # Prueba de personalización
    print("\n📝 Prueba de personalización:")
    plantilla = "Hello {NOMBRE}, welcome to {EMPRESA}! {MENSAJE_PERSONAL}\n\nBest regards,\n{REMITENTE_NOMBRE}"
    
    cliente_prueba = {
        'email': 'test@ejemplo.com',
        'nombre': '',
        'empresa': 'Test Company',
        'mensaje_personal': 'Hope you are doing well!'
    }
    
    config_prueba = {
        'Tu_Nombre': 'Admin',
        'Tu_Email': 'admin@sage.com'
    }
    
    resultado = processor.personalizar_contenido(plantilla, cliente_prueba, config_prueba)
    print("Resultado:")
    print(resultado)