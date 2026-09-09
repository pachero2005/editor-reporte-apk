import os
import json
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.graphics import Color, RoundedRectangle
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.utils import platform

# --- PERMISOS DE ANDROID ---
def solicitar_permisos_android():
    """Solicita permisos de almacenamiento en dispositivos Android"""
    if platform == 'android':
        try:
            from android.permissions import request_permissions, Permission
            request_permissions([
                Permission.READ_EXTERNAL_STORAGE, 
                Permission.WRITE_EXTERNAL_STORAGE
            ])
        except Exception as e:
            print(f"Error al solicitar permisos: {e}")

# --- GESTIÓN DE RUTAS Y ARCHIVOS JSON MAESTRO ---

def obtener_directorio_guardado():
    """Obtiene la carpeta 'Libro Diario' dentro de Documentos del usuario o dispositivo"""
    if platform == 'android':
        carpeta_documentos = '/storage/emulated/0/Documents'
        if not os.path.exists(carpeta_documentos):
            carpeta_documentos = '/storage/emulated/0/Documentos'
    else:
        user_dir = os.environ.get('USERPROFILE') or os.environ.get('HOME') or os.path.expanduser('~')
        carpeta_documentos = os.path.join(user_dir, 'Documents')
        if not os.path.exists(carpeta_documentos):
            carpeta_documentos = os.path.join(user_dir, 'Documentos')

    carpeta_destino = os.path.join(carpeta_documentos, 'Libro Diario')
    
    try:
        os.makedirs(carpeta_destino, exist_ok=True)
    except Exception as e:
        print(f"Error al crear carpeta pública, usando almacenamiento interno: {e}")
        app = App.get_running_app()
        if app and app.user_data_dir:
            carpeta_destino = os.path.join(app.user_data_dir, 'LibroDiario_Datos')
            os.makedirs(carpeta_destino, exist_ok=True)

    return carpeta_destino

def obtener_ruta_archivo_maestro():
    """Ruta del único archivo JSON general que contendrá todas las sesiones históricas"""
    return os.path.join(obtener_directorio_guardado(), 'LibroDiario_General.json')

def cargar_base_datos_maestra():
    """Carga el diccionario maestro filtrando hojas vacías"""
    ruta = obtener_ruta_archivo_maestro()
    if os.path.exists(ruta):
        try:
            with open(ruta, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict) and data:
                    data_filtrada = {k: v for k, v in data.items() if isinstance(v, list) and len(v) > 0}
                    if data_filtrada:
                        return data_filtrada
        except Exception as e:
            print(f"Error al cargar archivo maestro: {e}")
    
    timestamp_inicial = datetime.now().strftime("Libro_%Y-%m-%d_%H-%M-%S")
    return {timestamp_inicial: []}

def guardar_base_datos_maestra(data_dict):
    """Guarda el diccionario maestro en el único archivo JSON"""
    try:
        ruta = obtener_ruta_archivo_maestro()
        data_limpia = {k: v for k, v in data_dict.items() if isinstance(v, list) and len(v) > 0}
        
        if not data_limpia:
            if os.path.exists(ruta):
                os.remove(ruta)
            return

        with open(ruta, 'w', encoding='utf-8') as f:
            json.dump(data_limpia, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Error al guardar archivo maestro: {e}")

def exportar_a_excel_personalizado(registros, nombre_archivo_excel):
    """Crea o actualiza un archivo Excel independiente basado en el nombre de la sesión"""
    if not registros:
        return None

    carpeta = obtener_directorio_guardado()
    ruta_excel = os.path.join(carpeta, f"{nombre_archivo_excel}.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Libro Diario"

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)

    thin_side = Side(border_style="thin", color="D9D9D9")
    double_side = Side(border_style="double", color="000000")
    top_thin = Side(border_style="thin", color="000000")

    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(top=top_thin, bottom=double_side)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    headers = ["Fecha", "Detalle / Concepto", "Debe ($)", "Haber ($)"]
    ws.append(headers)

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = align_center

    for r in registros:
        ws.append([r.get("fecha", ""), r.get("detalle", ""), r.get("debe", 0.0), r.get("haber", 0.0)])

    total_debe = sum(r.get("debe", 0.0) for r in registros)
    total_haber = sum(r.get("haber", 0.0) for r in registros)
    diferencia = total_debe - total_haber

    ws.append([])
    ws.append(["TOTALES", "", total_debe, total_haber])
    ws.append(["DIFERENCIA", "", diferencia, ""])

    max_row = ws.max_row

    for row in range(2, max_row + 1):
        if row <= len(registros) + 1:
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.font = font_body
                cell.border = border_cell

                if col == 1:
                    cell.alignment = align_center
                elif col == 2:
                    cell.alignment = align_left
                elif col in [3, 4]:
                    cell.alignment = align_right
                    cell.number_format = '"$"#,##;("$"#,##);"-"'

        elif row in [max_row - 1, max_row]:
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.font = font_bold
                if col in [1, 2]:
                    cell.alignment = align_left
                elif col in [3, 4]:
                    cell.alignment = align_right
                    cell.number_format = '"$"#,##;("$"#,##);"-"'
                    if row == max_row - 1:
                        cell.border = border_total

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                val_str = f"${cell.value:,.2f}" if isinstance(cell.value, (int, float)) else str(cell.value)
                if len(val_str) > max_len:
                    max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    wb.save(ruta_excel)
    return ruta_excel

# --- COMPONENTES BASE ---

class CardContainer(BoxLayout):
    def __init__(self, bg_color=(0.14, 0.16, 0.20, 1), radius=12, **kwargs):
        super().__init__(**kwargs)
        self.bg_color = bg_color
        self.radius = radius
        self.bind(pos=self._update_canvas, size=self._update_canvas)

    def _update_canvas(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.bg_color)
            RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(self.radius)])

class ModernInput(TextInput):
    def __init__(self, is_numeric=False, **kwargs):
        super().__init__(**kwargs)
        self.background_active = ''
        self.background_normal = ''
        self.background_color = (0.08, 0.09, 0.12, 1)
        self.cursor_color = (0.3, 0.6, 1, 1)
        self.foreground_color = (1, 1, 1, 1)
        self.hint_text_color = (0.5, 0.55, 0.65, 1)
        self.multiline = False
        self.font_size = '15sp'
        self.padding = [dp(10), dp(10), dp(10), dp(10)]
        
        if is_numeric:
            self.input_type = 'number'
            self.input_filter = 'float'

class ModernButton(Button):
    def __init__(self, bg_color=(0.26, 0.53, 0.96, 1), radius=8, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ''
        self.background_down = ''
        self.background_color = (0, 0, 0, 0)
        self.bg_color = bg_color
        self.radius = radius
        self.bind(pos=self._update_canvas, size=self._update_canvas)

    def _update_canvas(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.bg_color)
            RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(self.radius)])

# --- APLICACIÓN PRINCIPAL ---

class LibroDiarioApp(App):
    def on_start(self):
        solicitar_permisos_android()

    def _guardar_estado_actual(self):
        """Guarda la sesión actual dentro del diccionario maestro y actualiza su propio Excel"""
        if hasattr(self, 'libros_data') and hasattr(self, 'nombre_hoja_actual'):
            if self.registros:
                self.libros_data[self.nombre_hoja_actual] = self.registros
                exportar_a_excel_personalizado(self.registros, self.nombre_hoja_actual)
            else:
                if self.nombre_hoja_actual in self.libros_data:
                    del self.libros_data[self.nombre_hoja_actual]
            
            guardar_base_datos_maestra(self.libros_data)

    def on_pause(self):
        self._guardar_estado_actual()
        return True

    def on_stop(self):
        self._guardar_estado_actual()

    def build(self):
        self.libros_data = cargar_base_datos_maestra()
        
        if not self.libros_data:
            timestamp_inicial = f"Libro_Diario_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
            self.libros_data = {timestamp_inicial: []}

        self.nombre_hoja_actual = list(self.libros_data.keys())[-1]
        self.registros = self.libros_data[self.nombre_hoja_actual]
        
        self.indice_edicion = None

        main_layout = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        with main_layout.canvas.before:
            Color(0.09, 0.10, 0.12, 1)
            RoundedRectangle(pos=(0, 0), size=(dp(2000), dp(2000)))

        lbl_titulo = Label(text="Libro Diario (Excel Independiente por Sesión)", font_size='20sp', bold=True, color=(1, 1, 1, 1), size_hint_y=None, height=dp(26), halign='left', valign='middle')
        lbl_subtitulo = Label(text="'+ Nuevo' crea un archivo Excel nuevo y cambia a él", font_size='12sp', color=(0.6, 0.65, 0.7, 1), size_hint_y=None, height=dp(18), halign='left', valign='middle')
        lbl_titulo.bind(size=lbl_titulo.setter('text_size'))
        lbl_subtitulo.bind(size=lbl_subtitulo.setter('text_size'))
        
        main_layout.add_widget(lbl_titulo)
        main_layout.add_widget(lbl_subtitulo)

        scroll_principal = ScrollView(bar_width=dp(4))
        contenedor_scroll = BoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None)
        contenedor_scroll.bind(minimum_height=contenedor_scroll.setter('height'))

        card_form = CardContainer(orientation='vertical', padding=dp(12), spacing=dp(8), size_hint_y=None)
        card_form.bind(minimum_height=card_form.setter('height'))

        box_fecha_cancelar = BoxLayout(orientation='horizontal', spacing=dp(6), size_hint_y=None, height=dp(62))
        self.txt_fecha = self._crear_campo("FECHA", datetime.now().strftime("%Y-%m-%d"))
        box_fecha_cancelar.add_widget(self.txt_fecha['container'])

        self.btn_cancelar = ModernButton(
            text="X Cancelar", 
            font_size='11sp', 
            bold=True, 
            bg_color=(0.75, 0.25, 0.25, 1), 
            size_hint=(None, None), 
            size=(dp(85), dp(38)), 
            pos_hint={'center_y': 0.35},
            opacity=0, 
            disabled=True
        )
        self.btn_cancelar.bind(on_release=self.cancelar_edicion)
        box_fecha_cancelar.add_widget(self.btn_cancelar)

        self.txt_detalle = self._crear_campo("DETALLE / CONCEPTO", "", hint="Ej. Ventas del día")

        box_montos = BoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(62))
        self.txt_debe = self._crear_campo("DEBE ($)", "", hint="0.00", is_numeric=True)
        self.txt_haber = self._crear_campo("HABER ($)", "", hint="0.00", is_numeric=True)
        box_montos.add_widget(self.txt_debe['container'])
        box_montos.add_widget(self.txt_haber['container'])

        box_botones_centro = BoxLayout(
            orientation='horizontal', 
            size_hint_y=None, 
            height=dp(42), 
            spacing=dp(6)
        )

        self.btn_abrir_hoja = ModernButton(
            text="Abrir", 
            font_size='12sp', 
            bold=True, 
            bg_color=(0.35, 0.30, 0.85, 1), 
            size_hint_x=0.25
        )
        self.btn_abrir_hoja.bind(on_release=self.mostrar_modal_abrir)

        self.btn_nueva_hoja = ModernButton(
            text="+ Nuevo", 
            font_size='12sp', 
            bold=True, 
            bg_color=(0.85, 0.50, 0.10, 1), 
            size_hint_x=0.25
        )
        self.btn_nueva_hoja.bind(on_release=self.nueva_hoja)

        self.btn_accion = ModernButton(
            text="+ Agregar", 
            font_size='12sp', 
            bold=True, 
            bg_color=(0.10, 0.60, 0.85, 1), 
            size_hint_x=0.25
        )
        self.btn_accion.bind(on_release=self.procesar_asiento)

        self.btn_guardar_excel = ModernButton(
            text="Excel", 
            font_size='12sp', 
            bold=True, 
            bg_color=(0.15, 0.65, 0.35, 1), 
            size_hint_x=0.25
        )
        self.btn_guardar_excel.bind(on_release=self.accion_guardar_excel)

        box_botones_centro.add_widget(self.btn_abrir_hoja)
        box_botones_centro.add_widget(self.btn_nueva_hoja)
        box_botones_centro.add_widget(self.btn_accion)
        box_botones_centro.add_widget(self.btn_guardar_excel)

        card_form.add_widget(box_fecha_cancelar)
        card_form.add_widget(self.txt_detalle['container'])
        card_form.add_widget(box_montos)
        card_form.add_widget(box_botones_centro)

        contenedor_scroll.add_widget(card_form)

        self.lbl_indicador_hoja = Label(text=f"SESIÓN ACTIVA: {self.nombre_hoja_actual}", font_size='11sp', bold=True, color=(0.4, 0.8, 1, 1), size_hint_y=None, height=dp(20), halign='left')
        self.lbl_indicador_hoja.bind(size=self.lbl_indicador_hoja.setter('text_size'))
        contenedor_scroll.add_widget(self.lbl_indicador_hoja)

        lbl_seccion = Label(text="MOVIMIENTOS REGISTRADOS", font_size='12sp', bold=True, color=(0.6, 0.65, 0.7, 1), size_hint_y=None, height=dp(20), halign='left')
        lbl_seccion.bind(size=lbl_seccion.setter('text_size'))
        contenedor_scroll.add_widget(lbl_seccion)

        self.lista_registros = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        self.lista_registros.bind(minimum_height=self.lista_registros.setter('height'))
        contenedor_scroll.add_widget(self.lista_registros)

        scroll_principal.add_widget(contenedor_scroll)
        main_layout.add_widget(scroll_principal)

        card_totales = CardContainer(orientation='horizontal', padding=[dp(10), dp(6), dp(10), dp(6)], spacing=dp(8), size_hint_y=None, height=dp(50), bg_color=(0.18, 0.21, 0.26, 1))
        
        box_totales_num = BoxLayout(orientation='vertical', spacing=dp(1), size_hint_x=0.5)
        self.lbl_total_debe = Label(text="Debe: $0.00", font_size='12sp', color=(0.7, 0.8, 0.7, 1), halign='left', valign='middle')
        self.lbl_total_haber = Label(text="Haber: $0.00", font_size='12sp', color=(0.8, 0.7, 0.7, 1), halign='left', valign='middle')
        self.lbl_total_debe.bind(size=self.lbl_total_debe.setter('text_size'))
        self.lbl_total_haber.bind(size=self.lbl_total_haber.setter('text_size'))
        box_totales_num.add_widget(self.lbl_total_debe)
        box_totales_num.add_widget(self.lbl_total_haber)

        self.lbl_diferencia = Label(text="Dif: $0.00", font_size='15sp', bold=True, color=(0.35, 0.85, 0.45, 1), size_hint_x=0.5, halign='right', valign='middle')
        self.lbl_diferencia.bind(size=self.lbl_diferencia.setter('text_size'))

        card_totales.add_widget(box_totales_num)
        card_totales.add_widget(self.lbl_diferencia)
        main_layout.add_widget(card_totales)

        self.actualizar_interfaz()

        return main_layout

    def _crear_campo(self, titulo, valor_inicial, hint="", is_numeric=False):
        box = BoxLayout(orientation='vertical', spacing=dp(2), size_hint_y=None, height=dp(62))
        lbl = Label(text=titulo, font_size='11sp', bold=True, color=(0.75, 0.8, 0.85, 1), size_hint_y=None, height=dp(16), halign='left', valign='middle')
        lbl.bind(size=lbl.setter('text_size'))
        input_field = ModernInput(text=valor_inicial, hint_text=hint, is_numeric=is_numeric, size_hint_y=None, height=dp(42))
        box.add_widget(lbl)
        box.add_widget(input_field)
        return {'container': box, 'input': input_field}

    def mostrar_modal_abrir(self, instance):
        self._guardar_estado_actual()
        nombres_hojas = [k for k, v in self.libros_data.items() if len(v) > 0]
        
        contenido = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        
        if not nombres_hojas:
            contenido.add_widget(Label(text="No hay sesiones guardadas.", color=(0.8, 0.8, 0.8, 1)))
        else:
            scroll = ScrollView()
            grid = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
            grid.bind(minimum_height=grid.setter('height'))

            for nombre in nombres_hojas:
                texto_btn = f"📁 {nombre}.xlsx"
                if nombre == self.nombre_hoja_actual:
                    texto_btn += " (Actual)"

                btn_hoja = ModernButton(
                    text=texto_btn, 
                    font_size='12sp', 
                    bg_color=(0.18, 0.22, 0.28, 1) if nombre != self.nombre_hoja_actual else (0.2, 0.45, 0.65, 1), 
                    size_hint_y=None, 
                    height=dp(45)
                )
                btn_hoja.bind(on_release=lambda b, nh=nombre: self.seleccionar_hoja(nh))
                grid.add_widget(btn_hoja)

            scroll.add_widget(grid)
            contenido.add_widget(scroll)

        btn_cerrar = ModernButton(text="Cerrar", font_size='13sp', bg_color=(0.4, 0.2, 0.2, 1), size_hint_y=None, height=dp(40))
        contenido.add_widget(btn_cerrar)

        popup = Popup(
            title="Seleccionar Archivo / Sesión", 
            content=contenido, 
            size_hint=(0.9, 0.7),
            background_color=(0.12, 0.14, 0.18, 1)
        )
        btn_cerrar.bind(on_release=popup.dismiss)
        self.popup_actual = popup
        popup.open()

    def seleccionar_hoja(self, nombre_hoja):
        self._guardar_estado_actual()
        self.nombre_hoja_actual = nombre_hoja
        self.registros = self.libros_data.get(nombre_hoja, [])
        self.cancelar_edicion()
        self.lbl_indicador_hoja.text = f"SESIÓN ACTIVA: {self.nombre_hoja_actual}"
        self.actualizar_interfaz()
        
        if hasattr(self, 'popup_actual'):
            self.popup_actual.dismiss()

    def nueva_hoja(self, instance=None):
        """Crea un nuevo archivo Excel independiente con marca de tiempo única"""
        self._guardar_estado_actual()
        
        nuevo_nombre = f"Libro_Diario_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
        self.libros_data[nuevo_nombre] = []
        self.nombre_hoja_actual = nuevo_nombre
        self.registros = self.libros_data[nuevo_nombre]

        self.cancelar_edicion()
        self._guardar_estado_actual()
        
        self.lbl_indicador_hoja.text = f"SESIÓN ACTIVA: {self.nombre_hoja_actual}"
        self.actualizar_interfaz()

        self.btn_nueva_hoja.text = "¡Nuevo Creado!"
        self.btn_nueva_hoja.bg_color = (0.2, 0.6, 0.35, 1)
        self.btn_nueva_hoja._update_canvas()

        def restaurar_btn(dt):
            self.btn_nueva_hoja.text = "+ Nuevo"
            self.btn_nueva_hoja.bg_color = (0.85, 0.50, 0.10, 1)
            self.btn_nueva_hoja._update_canvas()

        Clock.schedule_once(restaurar_btn, 1.5)

    def cargar_para_editar(self, index):
        def _inyeccion(dt):
            if index >= len(self.registros):
                return
            item = self.registros[index]
            self.txt_fecha['input'].text = str(item.get("fecha", ""))
            self.txt_detalle['input'].text = str(item.get("detalle", ""))
            self.txt_debe['input'].text = f"{item.get('debe', 0.0):.2f}"
            self.txt_haber['input'].text = f"{item.get('haber', 0.0):.2f}"

            self.indice_edicion = index
            self.btn_accion.text = "Guardar"
            self.btn_accion.bg_color = (0.2, 0.65, 0.32, 1)
            self.btn_accion._update_canvas()
            
            self.btn_cancelar.opacity = 1
            self.btn_cancelar.disabled = False

        Clock.schedule_once(_inyeccion, 0.05)

    def procesar_asiento(self, instance):
        fecha = str(self.txt_fecha['input'].text).strip()
        detalle = str(self.txt_detalle['input'].text).strip()
        
        try:
            debe = float(self.txt_debe['input'].text.strip() or 0.0)
            haber = float(self.txt_haber['input'].text.strip() or 0.0)
        except ValueError:
            return

        if not detalle:
            return

        registro = {"fecha": fecha, "detalle": detalle, "debe": debe, "haber": haber}

        if self.indice_edicion is not None:
            self.registros[self.indice_edicion] = registro
            self.resets_estado_boton()
        else:
            self.registros.append(registro)

        self._guardar_estado_actual()
        self.limpiar_campos()
        self.actualizar_interfaz()

    def accion_guardar_excel(self, instance):
        if not self.registros:
            return
            
        try:
            exportar_a_excel_personalizado(self.registros, self.nombre_hoja_actual)
            self.btn_guardar_excel.text = "¡Guardado!"
            self.btn_guardar_excel.bg_color = (0.1, 0.7, 0.3, 1)
            self.btn_guardar_excel._update_canvas()
        except Exception as e:
            self.btn_guardar_excel.text = "Error"
            self.btn_guardar_excel.bg_color = (0.8, 0.2, 0.2, 1)
            self.btn_guardar_excel._update_canvas()
            print(f"Error al guardar Excel: {e}")

        def restaurar_boton(dt):
            self.btn_guardar_excel.text = "Excel"
            self.btn_guardar_excel.bg_color = (0.15, 0.65, 0.35, 1)
            self.btn_guardar_excel._update_canvas()

        Clock.schedule_once(restaurar_boton, 2.0)

    def cancelar_edicion(self, instance=None):
        self.resets_estado_boton()
        self.limpiar_campos()

    def resets_estado_boton(self):
        self.indice_edicion = None
        self.btn_accion.text = "+ Agregar"
        self.btn_accion.bg_color = (0.10, 0.60, 0.85, 1)
        self.btn_accion._update_canvas()
        self.btn_cancelar.opacity = 0
        self.btn_cancelar.disabled = True

    def eliminar_asiento(self, index):
        if self.indice_edicion == index:
            self.cancelar_edicion()

        if index < len(self.registros):
            self.registros.pop(index)
            self._guardar_estado_actual()
            self.actualizar_interfaz()

    def actualizar_interfaz(self):
        def _render_tabla(dt):
            self.lista_registros.clear_widgets()
            total_debe = 0.0
            total_haber = 0.0

            for idx, item in enumerate(self.registros):
                total_debe += item.get("debe", 0.0)
                total_haber += item.get("haber", 0.0)

                card_item = CardContainer(orientation='horizontal', padding=[dp(12), dp(8), dp(8), dp(8)], spacing=dp(8), size_hint_y=None, height=dp(65), radius=8)
                info_box = BoxLayout(orientation='vertical', spacing=dp(2))
                
                lbl_det = Label(text=item.get('detalle', ''), font_size='15sp', bold=True, color=(1, 1, 1, 1), halign='left', valign='middle')
                lbl_sub = Label(text=f"{item.get('fecha', '')}  •  D: ${item.get('debe', 0.0):.2f}  |  H: ${item.get('haber', 0.0):.2f}", font_size='12sp', color=(0.7, 0.75, 0.8, 1), halign='left', valign='middle')
                lbl_det.bind(size=lbl_det.setter('text_size'))
                lbl_sub.bind(size=lbl_sub.setter('text_size'))
                
                info_box.add_widget(lbl_det)
                info_box.add_widget(lbl_sub)
                card_item.add_widget(info_box)

                btn_borrar = ModernButton(text="X", font_size='14sp', bold=True, bg_color=(0.45, 0.18, 0.22, 1), radius=6, size_hint=(None, None), size=(dp(36), dp(36)), pos_hint={'center_y': 0.5})
                btn_borrar.bind(on_release=lambda b, i=idx: self.eliminar_asiento(i))
                
                info_box.bind(on_touch_down=lambda instance, touch, i=idx: self.cargar_para_editar(i) if instance.collide_point(*touch.pos) else None)

                card_item.add_widget(btn_borrar)
                self.lista_registros.add_widget(card_item)

            diferencia = total_debe - total_haber

            self.lbl_total_debe.text = f"Debe: ${total_debe:.2f}"
            self.lbl_total_haber.text = f"Haber: ${total_haber:.2f}"
            self.lbl_diferencia.text = f"Dif: ${diferencia:.2f}"

            if diferencia < 0:
                self.lbl_diferencia.color = (0.95, 0.35, 0.35, 1)
            else:
                self.lbl_diferencia.color = (0.35, 0.85, 0.45, 1)

        Clock.schedule_once(_render_tabla, 0.01)

    def limpiar_campos(self):
        self.txt_fecha['input'].text = datetime.now().strftime("%Y-%m-%d")
        self.txt_detalle['input'].text = ""
        self.txt_debe['input'].text = ""
        self.txt_haber['input'].text = ""

if __name__ == "__main__":
    LibroDiarioApp().run()
