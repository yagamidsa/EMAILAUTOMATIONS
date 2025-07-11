import os
import pathlib
from typing import List, Dict
import logging

class FileManager:
    """Gestor de archivos adjuntos con validación de tamaño"""
    
    # Límites de tamaño en bytes
    MAX_TOTAL_SIZE = 20 * 1024 * 1024  # 20 MB total
    MAX_SINGLE_FILE = 10 * 1024 * 1024  # 10 MB por archivo
    OUTLOOK_LIMIT = 25 * 1024 * 1024   # 25 MB límite Outlook
    
    def __init__(self, adjuntos_folder: str = "adjuntos"):
        self.adjuntos_folder = adjuntos_folder
        
    def escanear_adjuntos(self) -> Dict:
        """Escanea la carpeta de adjuntos y valida archivos"""
        resultado = {
            'archivos': [],
            'total_archivos': 0,
            'total_tamaño': 0,
            'tamaño_ok': True,
            'advertencias': [],
            'errores': []
        }
        
        # Verificar si existe la carpeta
        if not os.path.exists(self.adjuntos_folder):
            os.makedirs(self.adjuntos_folder)
            resultado['advertencias'].append(f"Carpeta '{self.adjuntos_folder}' creada automáticamente")
            return resultado
            
        # Escanear archivos
        archivos_encontrados = []
        for archivo in os.listdir(self.adjuntos_folder):
            ruta_completa = os.path.join(self.adjuntos_folder, archivo)
            
            if os.path.isfile(ruta_completa):
                try:
                    tamaño = os.path.getsize(ruta_completa)
                    info_archivo = {
                        'nombre': archivo,
                        'ruta': ruta_completa,
                        'tamaño': tamaño,
                        'tamaño_texto': self._formatear_tamaño(tamaño),
                        'valido': True
                    }
                    
                    # Validar tamaño individual
                    if tamaño > self.MAX_SINGLE_FILE:
                        info_archivo['valido'] = False
                        resultado['errores'].append(
                            f"❌ '{archivo}' es muy grande ({self._formatear_tamaño(tamaño)}). "
                            f"Máximo: {self._formatear_tamaño(self.MAX_SINGLE_FILE)}"
                        )
                    elif tamaño > (self.MAX_SINGLE_FILE * 0.8):  # 80% del límite
                        resultado['advertencias'].append(
                            f"⚠️ '{archivo}' es grande ({self._formatear_tamaño(tamaño)})"
                        )
                    
                    archivos_encontrados.append(info_archivo)
                    resultado['total_tamaño'] += tamaño
                    
                except Exception as e:
                    resultado['errores'].append(f"Error leyendo '{archivo}': {str(e)}")
        
        resultado['archivos'] = archivos_encontrados
        resultado['total_archivos'] = len(archivos_encontrados)
        
        # Validar tamaño total
        if resultado['total_tamaño'] > self.OUTLOOK_LIMIT:
            resultado['tamaño_ok'] = False
            resultado['errores'].append(
                f"🚨 TAMAÑO TOTAL EXCESIVO: {self._formatear_tamaño(resultado['total_tamaño'])} "
                f"supera el límite de Outlook ({self._formatear_tamaño(self.OUTLOOK_LIMIT)}). "
                f"¡Los correos FALLARÁN!"
            )
        elif resultado['total_tamaño'] > self.MAX_TOTAL_SIZE:
            resultado['advertencias'].append(
                f"⚠️ TAMAÑO ALTO: {self._formatear_tamaño(resultado['total_tamaño'])} "
                f"supera lo recomendado ({self._formatear_tamaño(self.MAX_TOTAL_SIZE)})"
            )
        
        return resultado
    
    def _formatear_tamaño(self, tamaño_bytes: int) -> str:
        """Convierte bytes a formato legible"""
        if tamaño_bytes == 0:
            return "0 B"
        
        unidades = ['B', 'KB', 'MB', 'GB']
        tamaño = float(tamaño_bytes)
        indice = 0
        
        while tamaño >= 1024 and indice < len(unidades) - 1:
            tamaño /= 1024
            indice += 1
        
        return f"{tamaño:.1f} {unidades[indice]}"
    
    def obtener_resumen(self) -> str:
        """Obtiene un resumen legible de los adjuntos"""
        scan = self.escanear_adjuntos()
        
        resumen = "📎 ARCHIVOS ADJUNTOS:\n"
        resumen += "=" * 30 + "\n"
        
        if scan['total_archivos'] == 0:
            resumen += "📂 No hay archivos en la carpeta 'adjuntos/'\n"
            resumen += "💡 Puedes agregar PDFs, imágenes, documentos, etc.\n"
            return resumen
        
        # Información general
        estado = "✅" if scan['tamaño_ok'] else "❌"
        resumen += f"{estado} {scan['total_archivos']} archivo(s) encontrado(s)\n"
        resumen += f"📏 Tamaño total: {self._formatear_tamaño(scan['total_tamaño'])}\n"
        resumen += f"📊 Límite Outlook: {self._formatear_tamaño(self.OUTLOOK_LIMIT)}\n\n"
        
        # Lista de archivos
        resumen += "📋 Archivos:\n"
        for archivo in scan['archivos']:
            estado_archivo = "✅" if archivo['valido'] else "❌"
            resumen += f"   {estado_archivo} {archivo['nombre']} ({archivo['tamaño_texto']})\n"
        
        # Advertencias
        if scan['advertencias']:
            resumen += "\n⚠️ Advertencias:\n"
            for adv in scan['advertencias']:
                resumen += f"   {adv}\n"
        
        # Errores
        if scan['errores']:
            resumen += "\n❌ Errores:\n"
            for error in scan['errores']:
                resumen += f"   {error}\n"
        
        return resumen
    
    def obtener_archivos_validos(self) -> List[str]:
        """Retorna lista de rutas de archivos válidos para adjuntar"""
        scan = self.escanear_adjuntos()
        
        if not scan['tamaño_ok']:
            return []
        
        archivos_validos = []
        for archivo in scan['archivos']:
            if archivo['valido']:
                archivos_validos.append(archivo['ruta'])
        
        return archivos_validos

# Función de prueba
if __name__ == "__main__":
    print("🧪 Probando FileManager...")
    
    fm = FileManager()
    print(fm.obtener_resumen())
    
    archivos = fm.obtener_archivos_validos()
    print(f"\n🔗 Archivos válidos para adjuntar: {len(archivos)}")