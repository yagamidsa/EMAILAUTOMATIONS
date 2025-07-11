# build_executable.py
# Script para crear ejecutable de Email Sender Pro

import os
import subprocess
import sys
from pathlib import Path

def crear_ejecutable():
    """Crear ejecutable completo de Email Sender Pro"""
    
    print("🚀 CREANDO EJECUTABLE EMAIL SENDER PRO")
    print("=" * 50)
    
    # Verificar que existe el archivo principal
    archivo_principal = "src/gui.py"
    if not os.path.exists(archivo_principal):
        # Buscar archivo alternativo
        posibles = ["gui.py", "src/gui_inteligente.py", "gui_inteligente.py"]
        archivo_principal = None
        for archivo in posibles:
            if os.path.exists(archivo):
                archivo_principal = archivo
                break
        
        if not archivo_principal:
            print("❌ No se encontró el archivo GUI principal")
            print("📁 Archivos disponibles:")
            for archivo in os.listdir("."):
                if archivo.endswith(".py"):
                    print(f"   • {archivo}")
            return False
    
    print(f"✅ Archivo principal encontrado: {archivo_principal}")
    
    # Crear directorio de distribución
    dist_dir = "dist"
    os.makedirs(dist_dir, exist_ok=True)
    
    # Parámetros de PyInstaller
    comando = [
        "pyinstaller",
        "--onefile",                    # Un solo archivo ejecutable
        "--windowed",                   # Sin consola (importante)
        "--name=EmailSenderPro",        # Nombre del ejecutable
        "--icon=email_icon.ico",        # Icono (opcional)
        "--distpath=dist",              # Carpeta de salida
        "--workpath=build",             # Carpeta temporal
        "--specpath=.",                 # Archivo .spec
        "--clean",                      # Limpiar archivos temporales
        
        # Incluir archivos necesarios
        "--add-data=data;data",         # Carpeta data
        "--add-data=adjuntos;adjuntos", # Carpeta adjuntos
        
        # Librerías específicas
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.scrolledtext",
        "--hidden-import=tkinter.messagebox",
        
        archivo_principal
    ]
    
    # Ejecutar PyInstaller
    print(f"\n🔨 Compilando ejecutable...")
    print(f"📝 Comando: {' '.join(comando)}")
    
    try:
        resultado = subprocess.run(comando, check=True, capture_output=True, text=True)
        
        print("\n✅ COMPILACIÓN EXITOSA!")
        print(f"📁 Ejecutable creado en: dist/EmailSenderPro.exe")
        
        # Verificar que el archivo existe
        ejecutable = "dist/EmailSenderPro.exe"
        if os.path.exists(ejecutable):
            tamaño = os.path.getsize(ejecutable) / (1024 * 1024)  # MB
            print(f"📊 Tamaño: {tamaño:.1f} MB")
            
            # Crear carpetas necesarias junto al ejecutable
            crear_estructura_ejecutable()
            
            print(f"\n🎉 ¡LISTO PARA USAR!")
            print(f"📁 Ubicación: {os.path.abspath(ejecutable)}")
            
            return True
        else:
            print("❌ El ejecutable no se creó correctamente")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Error durante compilación:")
        print(f"   {e.stderr}")
        return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def crear_estructura_ejecutable():
    """Crear estructura de carpetas junto al ejecutable"""
    
    print(f"\n📁 Creando estructura de carpetas...")
    
    # Carpetas necesarias junto al ejecutable
    carpetas = [
        "dist/data",
        "dist/adjuntos", 
        "dist/reportes"
    ]
    
    for carpeta in carpetas:
        os.makedirs(carpeta, exist_ok=True)
        print(f"   ✅ {carpeta}")
    
    # Copiar archivos Excel de ejemplo si existen
    archivos_excel = [
        ("data/CAMPAÑAS.xlsx", "dist/data/CAMPAÑAS.xlsx"),
        ("data/CLIENTES.xlsx", "dist/data/CLIENTES.xlsx"),
        ("data/CONFIGURACION.xlsx", "dist/data/CONFIGURACION.xlsx")
    ]
    
    for origen, destino in archivos_excel:
        if os.path.exists(origen):
            import shutil
            shutil.copy2(origen, destino)
            print(f"   📊 Copiado: {os.path.basename(destino)}")
    
    # Crear archivo README
    crear_readme_ejecutable()

def crear_readme_ejecutable():
    """Crear README para el usuario del ejecutable"""
    
    readme_content = """📧 EMAIL SENDER PRO - GUÍA RÁPIDA
==========================================

🚀 CÓMO USAR:
1. Ejecuta EmailSenderPro.exe
2. Haz clic en "Actualizar Datos"
3. Revisa la campaña activa y lista de clientes
4. Agrega archivos adjuntos a la carpeta "adjuntos/"
5. Haz clic en "ENVÍO INTELIGENTE"

📁 ESTRUCTURA DE CARPETAS:
├── EmailSenderPro.exe          (El programa principal)
├── data/                       (Archivos Excel)
│   ├── CAMPAÑAS.xlsx          (Tus campañas de email)
│   ├── CLIENTES.xlsx          (Lista de clientes)
│   └── CONFIGURACION.xlsx     (Configuración del envío)
├── adjuntos/                   (Archivos para adjuntar)
└── reportes/                   (Reportes de envío generados)

📊 ARCHIVOS EXCEL:
• CAMPAÑAS.xlsx: Define el contenido y asunto de tus emails
• CLIENTES.xlsx: Lista de destinatarios con nombres y empresas  
• CONFIGURACION.xlsx: Configuración de envío (cuántos por día, etc.)

📎 ADJUNTOS:
• Coloca archivos PDF, imágenes, documentos en la carpeta "adjuntos/"
• Se adjuntarán automáticamente a todos los correos

📋 REPORTES:
• Después de cada envío se generan reportes automáticos
• CSV con exitosos y fallidos para seguimiento
• Úsalos para reintentar correos que fallaron

⚠️ REQUISITOS:
• Microsoft Outlook instalado y configurado
• Cuenta de email configurada en Outlook
• Windows 10/11

🆘 SOPORTE:
Si tienes problemas, revisa:
1. Que Outlook esté abierto y funcionando
2. Que tengas permisos de administrador
3. Que los archivos Excel estén bien configurados

✅ ¡Listo para enviar emails profesionalmente!
"""
    
    with open("dist/README.txt", "w", encoding="utf-8") as f:
        f.write(readme_content)
    
    print(f"   📄 README.txt creado")

def crear_icono():
    """Crear icono simple para el ejecutable"""
    
    # Este es un icono básico en formato ICO (base64)
    # En un proyecto real, usarías un archivo .ico profesional
    icono_base64 = """
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAABILAAASCwAAAAAA
AAAAAAAA////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A
////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/
2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/////wD///8A////ANnZ2f/Z2dn/2dnZ/9nZ
2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/////AP///wDZ2dn/
2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ////
/wD///8A2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/
2dnZ/9nZ2f////8A////ANnZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ
2f/Z2dn/2dnZ/9nZ2f/Z2dn/////AP///wDZ2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/
2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/////wD///8A2dnZ/9nZ2f/Z2dn/2dnZ/9nZ
2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f////8A////ANnZ2f/Z2dn/
2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/////AP//
/wDZ2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/
2dnZ/////wD///8A2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ
2f/Z2dn/2dnZ/9nZ2f////8A////ANnZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/
2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/////AP///wD///8A2dnZ/9nZ2f/Z2dn/2dnZ/9nZ
2f/Z2dn/2dnZ/9nZ2f/Z2dn/2dnZ/9nZ2f/Z2dn/////AP///wD///8A////AP///wD///8A
////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A
////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAA==
"""
    
    try:
        import base64
        icono_data = base64.b64decode(icono_base64)
        with open("email_icon.ico", "wb") as f:
            f.write(icono_data)
        print(f"   🎨 Icono creado")
        return True
    except:
        print(f"   ⚠️ No se pudo crear icono (continuando sin icono)")
        return False

def limpiar_archivos_temporales():
    """Limpiar archivos temporales después de la compilación"""
    
    import shutil
    
    archivos_limpiar = [
        "build",
        "EmailSenderPro.spec",
        "email_icon.ico"
    ]
    
    print(f"\n🧹 Limpiando archivos temporales...")
    
    for archivo in archivos_limpiar:
        try:
            if os.path.isdir(archivo):
                shutil.rmtree(archivo)
                print(f"   🗑️ Eliminado directorio: {archivo}")
            elif os.path.isfile(archivo):
                os.remove(archivo)
                print(f"   🗑️ Eliminado archivo: {archivo}")
        except Exception as e:
            print(f"   ⚠️ No se pudo eliminar {archivo}: {e}")

def main():
    """Función principal"""
    
    print("🔧 PREPARANDO COMPILACIÓN...")
    
    # Verificar dependencias
    try:
        import pandas
        import openpyxl
        import win32com.client
        print("✅ Dependencias verificadas")
    except ImportError as e:
        print(f"❌ Falta dependencia: {e}")
        print("📥 Instala con: pip install pandas openpyxl pywin32")
        return False
    
    # Crear icono
    crear_icono()
    
    # Crear ejecutable
    if crear_ejecutable():
        # Limpiar archivos temporales
        limpiar_archivos_temporales()
        
        print(f"\n🎉 ¡EJECUTABLE CREADO EXITOSAMENTE!")
        print(f"📁 Ubicación: dist/EmailSenderPro.exe")
        print(f"📊 Incluye: Archivos Excel, carpetas, README")
        print(f"✅ Listo para distribuir")
        
        return True
    else:
        print(f"\n❌ Error creando ejecutable")
        return False

if __name__ == "__main__":
    main()
    input("\nPresiona Enter para cerrar...")