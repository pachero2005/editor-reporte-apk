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

# --- GESTIÓN DE RUTAS Y ARCHIVOS ---

def obtener_directorio_guardado():
    """Obtiene la carpeta donde se guardan los libros de diario"""
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
        print(f"Error al crear carpeta: {e}")
        carpeta_destino = carpeta_documentos

    return carpeta_destino

def obtener_ruta_json_actual():
    """Ruta para la hoja de trabajo actual (borrador en curso)"""
    return os.path.join(obtener_directorio_guardado(), 'borrador_actual.json')

def guardar_datos_json(registros, ruta=None):
    """Guarda registros en un archivo JSON"""
    try:
        if not ruta:
            ruta = obtener_ruta_json_actual()
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
                data = json.load(f)
                if isinstance(data, list):
                    return data
    except Exception as e:
        print(f"Error al cargar JSON: {e}")
    return []

def listar_hojas_guardadas():
    """Lista todos los archivos de respaldo Excel o JSON guardados previamente"""
    carpeta = obtener_directorio_guardado()
    archivos = []
    if os.path.exists(carpeta):
        for f in os.listdir(carpeta):
            if f.startswith('Libro_Diario_') and (f.endswith('.json') or f.endswith('.xlsx')):
                if f not in archivos:
                    archivos.append(f)
    archivos.sort(reverse=True)
    return archivos

def exportar_a_excel(registros):
    """Crea el archivo Excel limpio sin generar copias JSON basura"""
    if not registros:
        return None

    carpeta = obtener_directorio_guardado()
    nombre_base = f'Libro_Diario_{datetime.now().strftime("%Y-%m-%d_%H%M%S")}'
    ruta_excel = os.path.join(carpeta, f'{nombre_base}.xlsx')

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

    def on_pause(self):
        guardar_datos_json(self.registros)
        return True

    def on_stop(self):
        guardar_datos_json(self.registros)

    def build(self):
        self.registros = cargar_datos_json()
        self.indice_edicion = None

        main_layout = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        with main_layout.canvas.before:
            Color(0.09, 0.10, 0.12, 1)
            RoundedRectangle(pos=(0, 0), size=(dp(2000), dp(2000)))

        # ENCABEZADO SIMPLE (TÍTULO Y SUBTÍTULO)
        lbl_titulo = Label(text="Libro Diario", font_size='20sp', bold=True, color=(1, 1, 1, 1), size_hint_y=None, height=dp(26), halign='left', valign='middle')
        lbl_subtitulo = Label(text="Control de asientos contables", font_size='12sp', color=(0.6, 0.65, 0.7, 1), size_hint_y=None, height=dp(18), halign='left', valign='middle')
        lbl_titulo.bind(size=lbl_titulo.setter('text_size'))
        lbl_subtitulo.bind(size=lbl_subtitulo.setter('text_size'))
        
        main_layout.add_widget(lbl_titulo)
        main_layout.add_widget(lbl_subtitulo)

        scroll_principal = ScrollView(bar_width=dp(4))
        contenedor_scroll = BoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None)
        contenedor_scroll.bind(minimum_height=contenedor_scroll.setter('height'))

        # FORMULARIO DE INGRESO DE DATOS
        card_form = CardContainer(orientation='vertical', padding=dp(12), spacing=dp(8), size_hint_y=None)
        card_form.bind(minimum_height=card_form.setter('height'))

        # 1. Campo Fecha y Botón Cancelar edición
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

        # 2. Campo Detalle
        self.txt_detalle = self._crear_campo("DETALLE / CONCEPTO", "", hint="Ej. Ventas del día")

        # 3. Campos DEBE y HABER
        box_montos = BoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(62))
        self.txt_debe = self._crear_campo("DEBE ($)", "", hint="0.00", is_numeric=True)
        self.txt_haber = self._crear_campo("HABER ($)", "", hint="0.00", is_numeric=True)
        box_montos.add_widget(self.txt_debe['container'])
        box_montos.add_widget(self.txt_haber['container'])

        # 4. BARRA CON LOS 4 BOTONES DEBAJO DE DEBE Y HABER
        box_botones_centro = BoxLayout(
            orientation='horizontal', 
            size_hint_y=None, 
            height=dp(42), 
            spacing=dp(6)
        )

        # Botón 1: ABRIR (Morado)
        self.btn_abrir_hoja = ModernButton(
            text="Abrir", 
            font_size='12sp', 
            bold=True, 
            bg_color=(0.35, 0.30, 0.85, 1), 
            size_hint_x=0.25
        )
        self.btn_abrir_hoja.bind(on_release=self.mostrar_modal_abrir)

        # Botón 2: NUEVO (Naranja)
        self.btn_nueva_hoja = ModernButton(
            text="+ Nuevo", 
            font_size='12sp', 
            bold=True, 
            bg_color=(0.85, 0.50, 0.10, 1), 
            size_hint_x=0.25
        )
        self.btn_nueva_hoja.bind(on_release=self.nueva_hoja)

        # Botón 3: AGREGAR / GUARDAR (Azul Cian)
        self.btn_accion = ModernButton(
            text="+ Agregar", 
            font_size='12sp', 
            bold=True, 
            bg_color=(0.10, 0.60, 0.85, 1), 
            size_hint_x=0.25
        )
        self.btn_accion.bind(on_release=self.procesar_asiento)

        # Botón 4: EXPORTAR EXCEL (Verde)
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

        # Agregar elementos al formulario en orden vertical
        card_form.add_widget(box_fecha_cancelar)
        card_form.add_widget(self.txt_detalle['container'])
        card_form.add_widget(box_montos)
        card_form.add_widget(box_botones_centro)

        contenedor_scroll.add_widget(card_form)

        lbl_seccion = Label(text="MOVIMIENTOS REGISTRADOS", font_size='12sp', bold=True, color=(0.6, 0.65, 0.7, 1), size_hint_y=None, height=dp(20), halign='left')
        lbl_seccion.bind(size=lbl_seccion.setter('text_size'))
        contenedor_scroll.add_widget(lbl_seccion)

        self.lista_registros = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        self.lista_registros.bind(minimum_height=self.lista_registros.setter('height'))
        contenedor_scroll.add_widget(self.lista_registros)

        scroll_principal.add_widget(contenedor_scroll)
        main_layout.add_widget(scroll_principal)

        # BARRA INFERIOR DE TOTALES
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

    # --- MODAL ABRIR HOJA ---
    def mostrar_modal_abrir(self, instance):
        archivos = listar_hojas_guardadas()
        
        contenido = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        
        if not archivos:
            contenido.add_widget(Label(text="No hay hojas guardadas previamente.", color=(0.8, 0.8, 0.8, 1)))
        else:
            scroll = ScrollView()
            grid = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
            grid.bind(minimum_height=grid.setter('height'))

            for nombre_archivo in archivos:
                nombre_visible = nombre_archivo.replace('Libro_Diario_', 'Hoja: ').replace('.json', '').replace('.xlsx', '')
                
                btn_archivo = ModernButton(
                    text=nombre_visible, 
                    font_size='13sp', 
                    bg_color=(0.18, 0.22, 0.28, 1), 
                    size_hint_y=None, 
                    height=dp(45)
                )
                btn_archivo.bind(on_release=lambda b, fn=nombre_archivo: self.cargar_hoja_seleccionada(fn))
                grid.add_widget(btn_archivo)

            scroll.add_widget(grid)
            contenido.add_widget(scroll)

        btn_cerrar = ModernButton(text="Cerrar", font_size='13sp', bg_color=(0.4, 0.2, 0.2, 1), size_hint_y=None, height=dp(40))
        contenido.add_widget(btn_cerrar)

        popup = Popup(
            title="Selecciona una Hoja guardada", 
            content=contenido, 
            size_hint=(0.9, 0.7),
            background_color=(0.12, 0.14, 0.18, 1)
        )
        btn_cerrar.bind(on_release=popup.dismiss)
        self.popup_actual = popup
        popup.open()

    def cargar_hoja_seleccionada(self, nombre_archivo):
        ruta = os.path.join(obtener_directorio_guardado(), nombre_archivo)
        if nombre_archivo.endswith('.json'):
            registros_cargados = cargar_datos_json(ruta)
        else:
            registros_cargados = [] # Si seleccionan un excel antiguo de respaldo, se maneja de forma segura
            
        if registros_cargados:
            self.registros = registros_cargados
            guardar_datos_json(self.registros)
            self.cancelar_edicion()
            self.actualizar_interfaz()
            
        if hasattr(self, 'popup_actual'):
            self.popup_actual.dismiss()

    def nueva_hoja(self, instance=None):
        self.registros = []
        self.cancelar_edicion()
        guardar_datos_json(self.registros)
        self.actualizar_interfaz()

        self.btn_nueva_hoja.text = "¡Limpia!"
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

        guardar_datos_json(self.registros)
        self.limpiar_campos()
        self.actualizar_interfaz()

    def accion_guardar_excel(self, instance):
        if not self.registros:
            return
            
        try:
            exportar_a_excel(self.registros)
            self.btn_guardar_excel.text = "¡Listo!"
            self.btn_guardar_excel.bg_color = (0.1, 0.7, 0.3, 1)
            self.btn_guardar_excel._update_canvas()
        except Exception as e:
            self.btn_guardar_excel.text = "Error"
            self.btn_guardar_excel.bg_color = (0.8, 0.2, 0.2, 1)
            self.btn_guardar_excel._update_canvas()
            print(f"Error al guardar: {e}")

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
            guardar_datos_json(self.registros)
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
