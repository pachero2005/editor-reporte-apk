import os
import json
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# --- FUNCIONES DE RUTAS Y ARCHIVOS ---

def obtener_directorio_guardado():
    """Define la ruta donde se guardarán los archivos (compatible con Android/PC)"""
    # Si estamos en Android, intentamos usar Documents, si no, la carpeta actual
    try:
        from android.storage import primary_external_storage_path
        d = primary_external_storage_path()
        carpeta = os.path.join(d, 'Documents', 'Libro Diario')
    except ImportError:
        carpeta = os.path.join(os.getcwd(), 'Libro Diario')
    
    if not os.path.exists(carpeta):
        os.makedirs(carpeta, exist_ok=True)
    return carpeta

def obtener_ruta_json_actual():
    """Retorna la ruta fija del archivo JSON único de borrador"""
    carpeta = obtener_directorio_guardado()
    return os.path.join(carpeta, 'borrador_actual.json')

def guardar_datos_json(registros, ruta=None):
    """Guarda todos los registros en un único archivo JSON constante"""
    if not ruta:
        ruta = obtener_ruta_json_actual()
    try:
        with open(ruta, 'w', encoding='utf-8') as f:
            json.dump(registros, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Error al guardar JSON: {e}")

def cargar_datos_json(ruta=None):
    """Carga registros desde un archivo JSON"""
    try:
        if not ruta:
            ruta = obtener_ruta_json_actual()
        if os.path.exists(ruta):
            with open(ruta, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error al cargar JSON: {e}")
    return []

def listar_hojas_guardadas():
    """Lista todos los archivos de respaldo Excel guardados previamente"""
    carpeta = obtener_directorio_guardado()
    archivos = []
    if os.path.exists(carpeta):
        for f in os.listdir(carpeta):
            if f.startswith('Libro_Diario_') and f.endswith('.xlsx'):
                archivos.append(f)
    archivos.sort(reverse=True)
    return archivos

def exportar_a_excel(registros):
    """Crea el archivo Excel y evita generar archivos JSON basura"""
    if not registros:
        return None

    carpeta = obtener_directorio_guardado()
    nombre_base = f'Libro_Diario_{datetime.now().strftime("%Y-%m-%d_%H%M%S")}'
    ruta_excel = os.path.join(carpeta, f'{nombre_base}.xlsx')
    
    # Nota: Se eliminaron las líneas que creaban copias JSON dinámicas aquí.

    wb = Workbook()
    ws = wb.active
    ws.title = "Libro Diario"

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)

    # Puedes continuar agregando aquí la lógica de llenado de celdas de tu app principal...
    
    wb.save(ruta_excel)
    return ruta_excel
