"""
Esta es la aplicación de captación de SocioGraph 
Author: Dvni (DSML) 

Start time: 04/03/2026
Finish time: 15/04/2026?

10/03/2026
Por ahora solo Dios y yo sabemos como funciona esto (330 líneas), si lo ves y tiene más de 1000, posiblemente solo
Dios sepa como funciona. Lo voy a intentar dejar lo más comentado posible

18/03/2026
Añadida la interfaz para poder seleccionar el guión y los requisitos, 
también se ha añadido un botón para comprobar si el contacto ya ha sido llamado

22/03/2026
Preparado para instalar Twillio. Me voy de vacaciones 3 semanas, espero acordarme de algo cuando vuelva,
o por lo menos que me sirva con mis comentarios

07/04/2026
He vuelto de vacaciones y he estado mirando el código, ni idea de como funciona esto.
Voy a intentar dejarlo lo más comentado posible para que en un futuro pueda entenderlo mejor,
también he añadido algunas mejoras
Buena suerte para entender esto.

13/04/2026
Instalado lo de twillio, y hecho algunos ajustes a la interfaz durante la llamada. Ahora se hace desde el navegador.
"""

#pip install PyQt5 PyQtWebEngine requests openpyxl python-dotenv

import sys
import os
import json
import time
import socket
import webbrowser
import random
import requests
import unicodedata
import datetime as _dt
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, Qt, QThread, pyqtSignal
from dotenv import load_dotenv
# ─── Configuración Dinámica (para poder editar config.py en el USB) ───
def get_runtime_config():
    """
    Obtiene la configuración en tiempo de ejecución (SERVER_URL y AGENT_ID).
    
    Busca un archivo 'config.py' al lado del ejecutable (si el script está congelado/compilado)
    o en el mismo directorio del script, y carga sus valores de forma dinámica.
    
    Returns:
        tuple: (SERVER_URL, AGENT_ID)
    """
    # Si es un .exe, buscamos al lado del ejecutable. Si es .py, al lado del script.
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    path = os.path.join(base_dir, "config.py")
    # Valores por defecto por si el archivo no existe o está roto
    conf = {"SERVER_URL": "https://sociograph-3488.twil.io", "AGENT_ID": "agente_desconocido"}
    
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                # Ejecutamos el config.py como código Python en un entorno controlado
                ldict = {}
                exec(f.read(), {}, ldict)
                if "SERVER_URL" in ldict: conf["SERVER_URL"] = ldict["SERVER_URL"]
                if "AGENT_ID" in ldict: conf["AGENT_ID"] = ldict["AGENT_ID"]
        except Exception as e:
            print(f"Error cargando config.py externo: {e}")
            
    return conf["SERVER_URL"], conf["AGENT_ID"]

SERVER_URL, AGENT_ID = get_runtime_config()

# Cargar variables de entorno para Netelip (.env)
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(base_dir, ".env"))

# ─────────────────────────────────────────────────────────────────────
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QPushButton,
    QComboBox,
    QSpinBox,
    QTextEdit,
    QLineEdit,
    QFileDialog,
    QMessageBox,
    QHBoxLayout,
    QVBoxLayout,
    QFormLayout,
    QGroupBox,
    QListWidget,
    QListWidgetItem,
)
from openpyxl import load_workbook
from shared_excel_access import excel_file_lock, ExcelLockTimeout

COL_REQUERIDAS = [
    "NOMBRE",
    "APELLIDO1",
    "TELEFONO1",
    "TELEFONO2",
    "EMAIL",
    "Respuesta",
    "Observaciones",
    "DOM_LOCALIDAD",
    "DOM_CP_CD",
    "SEXO_ID",
    "EDAD",
    "SEG_NM",
]

# Columna que usamos para marcar "en llamada activa"
# Si no existe en el Excel se crea automáticamente al guardar por primera vez
COL_EN_LLAMADA = "_En_llamada"

# Respuesta que indica "no contestó": la fila va al final, no se descarta
RESPUESTA_NO_CONTESTA = "6.   No contesta / no disponible"

# Identificador único de este operador/equipo (hostname + PID)
OPERADOR_ID = f"{socket.gethostname()}_{os.getpid()}"

# Timeout en segundos para considerar un bloqueo como fantasma (5 minutos)
TIMEOUT_BLOQUEO = 5 * 60


class HistoryWorker(QThread):
    """
    Hilo secundario para buscar en el Excel Histórico sin bloquear la interfaz gráfica (UI).
    """
    finished = pyqtSignal(list, str)  # lista de coincidencias (dicts), error_msg

    def __init__(self, path, nombre_buscar, apellido_buscar):
        """
        Inicializa el worker con los parámetros de búsqueda.
        
        Args:
            path (str): Ruta al archivo Excel Histórico (.xlsx).
            nombre_buscar (str): Nombre normalizado a buscar.
            apellido_buscar (str): Apellido normalizado a buscar.
        """
        super().__init__()
        self.path = path
        self.nombre_buscar = nombre_buscar
        self.apellido_buscar = apellido_buscar

    def _norm(self, s):
        """
        Normaliza una cadena de texto para comparación: elimina acentos,
        convierte a mayúsculas y reduce espacios múltiples.
        
        Args:
            s (str/None): Cadena a normalizar.
            
        Returns:
            str: Cadena de texto normalizada.
        """
        if s is None: return ""
        s = unicodedata.normalize("NFD", str(s))
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        return " ".join(s.upper().split())

    def run(self):
        """Aquí ocurre la magia: buscamos en el historial sin despeinar a la UI."""
        try:
            # Abrimos el Excel en modo solo lectura y solo datos para volar bajito
            wb = load_workbook(self.path, data_only=True, read_only=True)
            sheet = wb.active
            
            # Mapeamos las cabeceras para no ir a ciegas
            headers_norm = {}
            header_row_idx = None
            
            # Buscamos cabeceras en las primeras 5 filas (priorizando la 1)
            rows_gen = sheet.iter_rows(min_row=1, max_row=5, values_only=True)
            for i, row_values in enumerate(rows_gen, 1):
                temp_headers = {}
                if not row_values: continue
                for col_idx, val in enumerate(row_values, 1):
                    if val is not None:
                        norm_val = self._norm(val)
                        temp_headers[norm_val] = col_idx
                
                # Buscamos Nombre, Apellido y Teléfono
                if "NOMBRE" in temp_headers and "TELEFONO" in temp_headers:
                    headers_norm = temp_headers
                    header_row_idx = i
                    break
            
            if header_row_idx is None:
                self.finished.emit(False, "", "No se encontraron las columnas NOMBRE y TELEFONO en el historial.")
                wb.close()
                return

            # Definir las columnas necesarias para la búsqueda
            col_nombre   = headers_norm.get("NOMBRE")
            col_apellido = headers_norm.get("APELLIDO") or headers_norm.get("APELLIDO1") or headers_norm.get("APELLIDO 1")
            coincidencias = []

            # Búsqueda total con iter_rows
            data_start = header_row_idx + 1
            for row_values in sheet.iter_rows(min_row=data_start, values_only=True):
                if not row_values or len(row_values) < max(headers_norm.values()):
                    continue

                # Normalizamos nombre y apellido del historial para una comparación justa
                h_nombre   = self._norm(row_values[col_nombre - 1])
                h_apellido = self._norm(row_values[col_apellido - 1]) if col_apellido else ""
                
                # ¡Bingo! Comprobamos si es la persona que buscamos (solo por nombre y apellido)
                nombre_ok = (h_nombre == self.nombre_buscar and h_nombre != "")
                apellido_ok = (not col_apellido) or (h_apellido == self.apellido_buscar)

                if nombre_ok and apellido_ok:
                    # Si hay match, recolectamos hasta el último dato de la fila
                    datos_fila = {"es_historial": True}
                    for norm_name, col_idx in headers_norm.items():
                        val = row_values[col_idx - 1]
                        # Las fechas mejor verlas en formato humano (DD/MM/YYYY)
                        if isinstance(val, (_dt.datetime, _dt.date)):
                            val = val.strftime("%d/%m/%Y")
                        datos_fila[norm_name] = str(val or "").strip() if val is not None else "-"
                    
                    coincidencias.append(datos_fila)

            wb.close()
            self.finished.emit(coincidencias, "")

        except Exception as e:
            self.finished.emit([], str(e))


class CallApp(QWidget):
    """
    Clase principal de la aplicación de captación telefónica (interfaz gráfica en PyQt5).
    
    Gestiona la carga de una hoja de cálculo Excel (.xlsx), la aplicación de filtros demográficos,
    la visualización del contacto seleccionado, la integración con el softphone web de Netelip
    mediante WebRTC/SIP, y el guardado seguro de respuestas en red compartida utilizando un
    sistema de bloqueos por archivos JSON y semáforos a nivel de archivo.
    """

    def __init__(self):
        """
        Inicializa la ventana principal, define las variables de control
        y construye la interfaz de usuario.
        """
        super().__init__()
        self.excel_path = None
        self.mtime_excel = None
        self.snapshot_disco = None
        self.workbook = None
        self.sheet = None
        self.cols = {}
        self.cache_filas = []
        self.fila_actual = None
        self.ventana_softphone = None
        self.script_docx_path = None
        self.requisitos_docx_path = None
        self.setWindowTitle("Gestor de Llamadas Excel")
        self.resize(900, 500)
        self._crear_interfaz()
        self._estado_inicial()
        self._poner_respuestas([])

    def _poner_respuestas(self, opciones):
        """
        Carga la lista de opciones de respuesta en el combobox de la ficha del contacto.
        
        Args:
            opciones (list): Lista de cadenas de texto con las respuestas válidas.
        """
        self.respuesta_combo.blockSignals(True)
        try:
            self.respuesta_combo.clear()
            for opt in opciones:
                self.respuesta_combo.addItem(str(opt))
        finally:
            self.respuesta_combo.blockSignals(False)

    # INTERFAZ / UI 
    def _crear_interfaz(self):
        """
        Crea y distribuye los componentes visuales (widgets) de la ventana,
        organizando el layout principal en dos paneles: panel izquierdo para filtros/guión,
        y panel derecho para la ficha del contacto actual.
        """
        main_layout = QHBoxLayout(self)
        # zona izquierda: filtros y acción
        left_box = QGroupBox("Filtros y acciones")
        left_layout = QVBoxLayout()

        # botón para selección de archivo Excel
        file_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Ruta al Excel compartido (.xlsx)")
        browse_btn = QPushButton("Examinar")
        browse_btn.clicked.connect(self.browse_excel)
        load_btn = QPushButton("Cargar Excel")
        load_btn.clicked.connect(self.load_excel_clicked)
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(browse_btn)
        file_layout.addWidget(load_btn)

        # Filtros
        filters_box = QGroupBox("Filtros")
        filters_layout = QFormLayout()
        self.localidad_combo = QComboBox()
        self.sexo_combo = QComboBox()
        self.respuesta_filter_combo = QComboBox()
        self.edad_min_spin = QSpinBox()
        self.edad_min_spin.setRange(18, 98)
        self.edad_min_spin.setValue(18)
        self.edad_max_spin = QSpinBox()
        self.edad_max_spin.setRange(19, 99)
        self.edad_max_spin.setValue(60)
        filters_layout.addRow("DOM_LOCALIDAD:", self.localidad_combo)
        filters_layout.addRow("SEXO_ID:", self.sexo_combo)
        self.respuesta_filter_combo.addItems(
            [
                "No activado",
                "3.   Reserva",
                "4.   Volver a llamar (indicar cuando)",
            ]
        )
        self.respuesta_filter_combo.setCurrentText("No activado")
        self.respuesta_filter_combo.currentTextChanged.connect(self._filtro_respuesta_cambiado)
        filters_layout.addRow("Respuesta:", self.respuesta_filter_combo)
        filters_layout.addRow("Edad mín:", self.edad_min_spin)
        filters_layout.addRow("Edad máx:", self.edad_max_spin)
        filters_box.setLayout(filters_layout)


        # Guión sepuede editar pero no se guarda, pon lo que quieras
        self.script_btn = QPushButton("Selección guión")
        self.script_btn.clicked.connect(self.select_script_docx)
        self.campo1_edit = QTextEdit()
        self.campo1_edit.setPlaceholderText(
            "Carga un Word (.docx) para mostrar aquí el guión"
        )
        self.campo1_edit.setMinimumHeight(120)

        # Boton de comprobar en el excel historial
        self.comprobar_btn = QPushButton("Comprobar")
        self.comprobar_btn.clicked.connect(self.handle_comprobar)

        # Requisitos (solo leer)
        self.requisitos_btn = QPushButton("Selección requisitos")
        self.requisitos_btn.clicked.connect(self.select_requisitos_docx)
        self.campo2_edit = QTextEdit()
        self.campo2_edit.setReadOnly(True)
        self.campo2_edit.setPlaceholderText(
            "Carga un Word (.docx) para mostrar aquí los requisitos (solo lectura)."
        )
        self.campo2_edit.setMinimumHeight(80)

        common_box_style = "background-color: #f0f0f0; padding: 6px; border-radius: 3px;"
        self.campo1_edit.setStyleSheet(common_box_style)
        self.campo2_edit.setStyleSheet(common_box_style)
        # Botón LLAMAR
        self.call_button = QPushButton("LLAMAR")
        self.call_button.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        self.call_button.clicked.connect(self.handle_call)
        self.load_contact_button = QPushButton("CARGAR CONTACTO")
        self.load_contact_button.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        self.load_contact_button.clicked.connect(self.handle_load_contact)
        left_layout.addLayout(file_layout)
        left_layout.addWidget(filters_box)
        left_layout.addWidget(self.script_btn)
        left_layout.addWidget(self.campo1_edit)
        left_layout.addWidget(self.requisitos_btn)
        left_layout.addWidget(self.campo2_edit)
        left_layout.addStretch()
        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(self.comprobar_btn)
        buttons_layout.addWidget(self.load_contact_button)
        buttons_layout.addWidget(self.call_button)
        left_layout.addLayout(buttons_layout)
        left_box.setLayout(left_layout)
        
        # Zona derecha: ficha de contacto
        right_box = QGroupBox("Ficha de contacto actual")
        right_layout = QVBoxLayout()
        form_layout = QFormLayout()
        self.nombre_label = QLabel("-")
        self.apellido_label = QLabel("-")
        self.telefono_label = QLabel("-")
        self.email_label = QLabel("-")
        form_layout.addRow("NOMBRE:", self.nombre_label)
        form_layout.addRow("APELLIDO1:", self.apellido_label)
        form_layout.addRow("TELEFONO:", self.telefono_label)
        form_layout.addRow("EMAIL:", self.email_label)

        self.respuesta_combo = QComboBox()
        self.respuesta_combo.setEditable(True)
        if self.respuesta_combo.lineEdit() is not None:
            self.respuesta_combo.lineEdit().setPlaceholderText("")
        self.observaciones_edit = QTextEdit()

        form_layout.addRow("Respuesta:", self.respuesta_combo)
        form_layout.addRow("Observaciones:", self.observaciones_edit)
        right_layout.addLayout(form_layout)

        self.lista_especial = QListWidget()
        self.lista_especial.itemDoubleClicked.connect(self._dobleclick_lista)
        right_layout.addWidget(self.lista_especial)

        self.save_button = QPushButton("GUARDAR")
        self.save_button.setStyleSheet("font-size: 14px; font-weight: bold; padding: 6px;")
        self.save_button.clicked.connect(self.handle_save)
        right_layout.addWidget(self.save_button, alignment=Qt.AlignRight)
        right_box.setLayout(right_layout)
        main_layout.addWidget(left_box, stretch=1)
        main_layout.addWidget(right_box, stretch=2)

    def _estado_inicial(self):
        """
        Deshabilita los controles de filtrado e interacción hasta que
        se realice la carga correcta del archivo Excel.
        """
        self.call_button.setEnabled(False)
        self.load_contact_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.localidad_combo.setEnabled(False)
        self.sexo_combo.setEnabled(False)
        self.respuesta_filter_combo.setEnabled(False)
        self.edad_min_spin.setEnabled(False)
        self.edad_max_spin.setEnabled(False)
        self.lista_especial.setEnabled(False)
        # Los Word se pueden seleccionar siempre, aun sin Excel cargado
        self.script_btn.setEnabled(True)
        self.requisitos_btn.setEnabled(True)

    # ------------- Manejo de Word (.docx) -------------
    def _leer_docx(self, path):
        """
        Lee y extrae el texto plano de un documento Word (.docx).
        
        Args:
            path (str): Ruta absoluta o relativa al archivo Word.
            
        Returns:
            str: Texto concatenado por saltos de línea extraído del archivo.
            
        Raises:
            RuntimeError: Si la librería python-docx no está instalada.
        """
        try:
            from docx import Document  # type: ignore
        except Exception:
            raise RuntimeError(
                "Falta la librería 'python-docx' para leer archivos Word (.docx). "
                "Instálala con: pip install python-docx"
            )

        doc = Document(path)
        partes = []
        for p in doc.paragraphs:
            txt = (p.text or "").rstrip()
            if txt:
                partes.append(txt)
        return "\n".join(partes).strip()

    def select_script_docx(self):
        """
        Muestra un diálogo de selección para cargar un archivo Word (.docx)
        con el guión de llamada y mostrarlo en la interfaz.
        """
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar guión (Word)",
            "",
            "Documentos Word (*.docx)",
        )
        if not path:
            return
        try:
            text = self._leer_docx(path)
            self.script_docx_path = path
            self.campo1_edit.setPlainText(text)
        except Exception as e:
            self._error(f"No se pudo cargar el guión desde Word:\n{e}")

    def select_requisitos_docx(self):
        """
        Muestra un diálogo de selección para cargar un archivo Word (.docx)
        con los requisitos de selección y mostrarlo en modo de solo lectura.
        """
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar requisitos (Word)",
            "",
            "Documentos Word (*.docx)",
        )
        if not path:
            return
        try:
            text = self._leer_docx(path)
            self.requisitos_docx_path = path
            self.campo2_edit.setPlainText(text)
        except Exception as e:
            self._error(f"No se pudieron cargar los requisitos desde Word:\n{e}")

    # ------------- Manejo de Excel -------------
    def browse_excel(self):
        """
        Muestra un diálogo de selección de archivos del sistema para
        elegir la ruta del archivo Excel (.xlsx).
        """
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo Excel",
            "",
            "Archivos Excel (*.xlsx)",
        )
        if path:
            self.file_path_edit.setText(path)

    def load_excel_clicked(self):
        """
        Manejador del evento del botón 'Cargar Excel'. Lee el path ingresado,
        valida su existencia e inicia el proceso de inicialización del archivo.
        """
        path = self.file_path_edit.text().strip()
        if not path:
            self._error("Debes seleccionar un Excel.")
            return
        if not os.path.exists(path):
            self._error("El archivo Excel especificado no existe.")
            return
        try:
            self.load_excel(path)
            self._info("Excel cargado correctamente.")
        except Exception as e:
            self._error(f"Error al cargar el Excel:\n{e}")

    def load_excel(self, path):
        """
        Carga e inicializa el libro de Excel, localizando la hoja que contenga
        todas las columnas obligatorias en su primera fila.
        
        Además, genera la columna interna de estado de llamada si no existe,
        carga las opciones de respuesta predefinidas desde la segunda hoja,
        llena la caché de filas con los filtros correspondientes y limpia bloqueos fantasmas.
        
        Args:
            path (str): Ruta al archivo Excel (.xlsx) que se desea cargar.
        """
        wb = load_workbook(path, data_only=True)

        # buscar la hoja que tenga todas las columnas requeridas en la fila 1
        sheet = None
        headers = {}
        for candidate in wb.worksheets:
            tmp_headers = {}
            for col in range(1, candidate.max_column + 1):
                value = candidate.cell(row=1, column=col).value
                if value is not None:
                    tmp_headers[str(value)] = col
            missing = [c for c in COL_REQUERIDAS if c not in tmp_headers]
            if not missing:
                sheet = candidate
                headers = tmp_headers
                break
        if sheet is None:
            raise ValueError(
                "Faltan columnas requeridas en el Excel, comprueba que estén estas: "
                + ", ".join(COL_REQUERIDAS)
            )
        self.excel_path = path
        try:
            self.mtime_excel = os.path.getmtime(path)
        except OSError:
            self.mtime_excel = None
        self.workbook = wb
        self.sheet = sheet
        self.cols = headers

        # si la columna de en llamada todavia no existe la metemos a tomar por culo a la derecha
        # max_column + 10 de margen pa no pisar nadaxd
        if COL_EN_LLAMADA not in self.cols:
            real_max_col = sheet.max_column if sheet.max_column is not None else 1
            safe_col = real_max_col + 10
            self.cols[COL_EN_LLAMADA] = safe_col

        self._cargar_filas_y_filtros()
        self._limpiar_bloqueos_fantasma()

        # respuestas de la segunda hoja IMPORTANTE QUE ESTÉ EN LA SEGUNDA HOJA, LA PRIMERA RESPUESTA EN LA LÍNEA 5
        respuestas = []
        try:
            if len(wb.worksheets) >= 2:
                hoja_respuestas = wb.worksheets[1]
                start_row = 5
                end_row = min(133, hoja_respuestas.max_row)
                for row in range(start_row, end_row + 1):
                    val = hoja_respuestas.cell(row=row, column=2).value
                    if val not in (None, ""):
                        texto = str(val).strip()
                        if texto:
                            respuestas.append(texto)
            respuestas_unicas = list(dict.fromkeys(respuestas))
            self._poner_respuestas(respuestas_unicas)
        except Exception as e:
            self._manejar_error(e, "Cargar respuestas (Hoja 2 B5:B133)")

        # Activar UI de filtros y llamadas
        self.localidad_combo.setEnabled(True)
        self.sexo_combo.setEnabled(True)
        self.respuesta_filter_combo.setEnabled(True)
        self.edad_min_spin.setEnabled(True)
        self.edad_max_spin.setEnabled(True)
        self.call_button.setEnabled(True)
        self.load_contact_button.setEnabled(True)
        self.save_button.setEnabled(True)
        self.lista_especial.setEnabled(True)

    # ─── Sistema de bloqueo por JSON (rápido, fiable en red) ───

    def _locks_path(self) -> str:
        """Ruta del fichero de bloqueos JSON junto al Excel."""
        return self.excel_path + ".locks.json"

    def _leer_locks(self) -> dict:
        """Lee el fichero de bloqueos con reintentos fuertes antibloqueo de SharePoint."""
        path = self._locks_path()
        if not os.path.exists(path):
            return {}
            
        for intento in range(15):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                    if not content:
                        return {}
                    return json.loads(content)
            except (json.JSONDecodeError, PermissionError):
                # SharePoint o el antivirus está tocando el archivo, esperamos y reintentamos
                time.sleep(0.2)
        
        # Si después de 3 segundos no podemos leerlo, devolvemos {} 
        # pero logueamos el fallo para saber qué pasa
        self._registrar_error("Fallo al leer JSON de bloqueos tras 15 reintentos (SharePoint lock)")
        return {}

    def _escribir_locks(self, locks: dict) -> None:
        """Escribe el fichero de bloqueos atómicamente, con reintentos pesados."""
        path = self._locks_path()
        tmp_path = path + ".tmp"
        
        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(locks, f)
        except Exception as e:
            self._registrar_error(f"No se pudo escribir el temporal del JSON: {e}")
            return
        
        # SharePoint bloquea el archivo mientras lo sube. 
        # Si da error, esperamos y reintentamos a lo cabezota (hasta 4 segundos).
        for intento in range(20):
            try:
                if os.path.exists(path):
                    os.replace(tmp_path, path)
                else:
                    os.rename(tmp_path, path)
                return 
            except PermissionError:
                time.sleep(0.2)
        
        try:
            os.replace(tmp_path, path)
        except Exception as e:
            self._registrar_error(f"Fallo épico guardando JSON tras 20 reintentos: {e}")

    def _es_mi_bloqueo(self, lock_val) -> bool:
        """Comprueba si un valor de bloqueo es nuestro (este proceso)."""
        if lock_val is None or str(lock_val).strip() == "":
            return False
        lock_str = str(lock_val).strip()
        return lock_str == OPERADOR_ID or lock_str.startswith(OPERADOR_ID + "|")

    def _valor_bloqueo(self) -> str:
        """Genera el valor de bloqueo con timestamp: OPERADOR_ID|unix_timestamp"""
        return f"{OPERADOR_ID}|{int(time.time())}"

    def _bloqueo_expirado(self, lock_val) -> bool:
        """Comprueba si un bloqueo ha superado TIMEOUT_BLOQUEO segundos."""
        if lock_val is None or str(lock_val).strip() == "":
            return False
        lock_str = str(lock_val).strip()
        if "|" not in lock_str:
            return False  # formato viejo sin timestamp, asumimos activo
        try:
            ts = int(lock_str.split("|")[1])
            return (time.time() - ts) > TIMEOUT_BLOQUEO
        except (ValueError, IndexError):
            return False

    def _limpiar_bloqueos_fantasma(self):
        """Limpia bloqueos expirados (>5 min) del JSON."""
        if self.excel_path is None:
            return
        try:
            lock_file = self._locks_path() + ".lock"
            # usamos un .lock propio para el json, asi no nos pisamos entre nosotros
            start = time.time()
            while True:
                try:
                    fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(fd)
                    break
                except FileExistsError:
                    if time.time() - start > 10:
                        return  # no pudimos bloquear, pues pasamos olímpicamente
                    time.sleep(0.3)

            try:
                locks = self._leer_locks()
                limpiados = 0
                filas_a_borrar = []
                for row_key, lock_val in locks.items():
                    if self._bloqueo_expirado(lock_val):
                        # pumba, a la calle por pesao
                        filas_a_borrar.append(row_key)
                        limpiados += 1
                for row_key in filas_a_borrar:
                    del locks[row_key]
                if limpiados > 0:
                    self._escribir_locks(locks)
                    self._registrar_error(f"Limpiados {limpiados} bloqueo(s) fantasma del JSON, puro clean up")
            finally:
                try:
                    os.remove(lock_file)
                except (FileNotFoundError, PermissionError):
                    pass
        except Exception as e:
            self._registrar_error(f"Error limpiando bloqueos fantasma: {e}")

    def _cargar_filas_y_filtros(self):
        assert self.sheet is not None
        self.cache_filas.clear()
        localidades = set()
        sexos = set()
        for row in range(2, self.sheet.max_row + 1):
            row_data = {"_row": row}
            for col_name, col_idx in self.cols.items():
                row_data[col_name] = self.sheet.cell(row=row, column=col_idx).value
            self.cache_filas.append(row_data)
            loc = row_data.get("DOM_LOCALIDAD")
            sexo = row_data.get("SEXO_ID")
            if loc not in (None, ""):
                localidades.add(str(loc))
            if sexo not in (None, "") and str(sexo).strip().upper() != "NB":
                sexos.add(str(sexo))
        # rellenar con los valores únicos
        self.localidad_combo.clear()
        self.localidad_combo.addItem("Todos")
        for val in sorted(localidades):
            self.localidad_combo.addItem(val)
        self.sexo_combo.clear()
        self.sexo_combo.addItem("Todos")
        for val in sorted(sexos):
            self.sexo_combo.addItem(val)

    # filtros y llamada 
    def _pasa_filtros(self, row_data):
        """
        Evalúa si un registro cumple con los filtros activos (localidad, sexo, respuesta y edad).
        
        Args:
            row_data (dict): Diccionario con los datos del contacto.
            
        Returns:
            bool: True si el registro pasa todos los filtros, False en caso contrario.
        """
        # Filtro localidad
        selected_loc = self.localidad_combo.currentText()
        if selected_loc != "Todos":
            if str(row_data.get("DOM_LOCALIDAD", "")).strip() != selected_loc:
                return False
        # Filtro sexo
        selected_sexo = self.sexo_combo.currentText()
        if selected_sexo != "Todos":
            if str(row_data.get("SEXO_ID", "")).strip() != selected_sexo:
                return False
        # Filtro respuesta
        selected_resp = self.respuesta_filter_combo.currentText()
        if selected_resp not in ("Todos", "No activado"):
            resp_val = row_data.get("Respuesta")
            if resp_val is None or str(resp_val).strip() == "":
                return False
            if str(resp_val).strip() != selected_resp:
                return False
                
        # Filtro edad
        edad_min = self.edad_min_spin.value()
        edad_max = self.edad_max_spin.value()
        edad_val = row_data.get("EDAD")
        if edad_val is None or str(edad_val).strip() == "":
            return False
        try:
            edad_int = int(edad_val)
        except Exception:
            return False
        if not (edad_min <= edad_int <= edad_max):
            return False
        return True

    def _filtro_respuesta_cambiado(self, selected):
        """
        Filtra la lista de contactos en caché según el estado de respuesta seleccionado
        ("Volver a llamar" o "Reserva") y llena la lista lateral.
        
        Args:
            selected (str): El valor seleccionado del combo de filtrado de respuestas.
        """
        if "Volver a llamar" in selected:
            keyword = "Volver a llamar"
        elif "Reserva" in selected:
            keyword = "Reserva"
        else:
            self.lista_especial.clear()
            return

        self.lista_especial.clear()
        for row_data in self.cache_filas:
            resp_val = row_data.get("Respuesta")
            if resp_val is None:
                continue
            resp_str = str(resp_val).strip()
            if keyword not in resp_str:
                continue
            nombre = str(row_data.get("NOMBRE") or "").strip()
            apellido = str(row_data.get("APELLIDO1") or "").strip()
            texto = (nombre + " " + apellido).strip() or "-"
            item = QListWidgetItem(texto)
            item.setData(Qt.UserRole, row_data)
            self.lista_especial.addItem(item)

    def _dobleclick_lista(self, item):
        """
        Manejador del evento doble clic en los elementos de la lista lateral.
        
        Si el elemento proviene de una búsqueda histórica, muestra un cuadro con las observaciones.
        Si es un contacto cargado para re-llamar, gestiona el bloqueo atómico y abre el softphone.
        
        Args:
            item (QListWidgetItem): Elemento de la lista sobre el cual se hizo doble clic.
        """
        row_data = item.data(Qt.UserRole)
        if not isinstance(row_data, dict):
            return

        # ¿Es un contacto del historial? Aquí no llamamos, solo chismeamos las notas
        if row_data.get("es_historial"):
            obs = row_data.get("OBSERVACIONES", "Sin observaciones.")
            # Ventanita limpia solo con lo que el agente necesita leer
            QMessageBox.information(self, "Observaciones Histórico", str(obs))
            return

        phone = self._sacar_telefono(row_data)
        if phone is None:
            self._error("El contacto seleccionado no tiene un teléfono válido.")
            return

        # mismo flujo que en handle_call tras elegir contacto
        if self.fila_actual is not None:
            self._liberar_fila(self.fila_actual["data"])
        
        # Intentamos bloqueo atómico
        if not self._bloquear_fila(row_data):
            self._error("No se pudo bloquear este contacto. Posiblemente otro operador ya lo ha ocupado.")
            # Refrescamos para quitarlo de la lista o marcarlo
            self._filtro_respuesta_cambiado(self.respuesta_filter_combo.currentText())
            return

        self.fila_actual = {"data": row_data, "phone": phone}
        self._mostrar_contacto(row_data)
        self._llamar_adb(phone)

    def _coincide_respuesta(self, row_data):
        """
        Determina si la respuesta asignada al contacto coincide con la actualmente seleccionada
        en el combobox.
        
        Args:
            row_data (dict): Diccionario con los datos del contacto.
            
        Returns:
            bool: True si coincide o si el combo está en "Todos", False de lo contrario.
        """
        selected_respuesta = self.respuesta_combo.currentText()
        if selected_respuesta != "Todos":
            if str(row_data.get("Respuesta", "")).strip() != selected_respuesta:
                return False
        return True

    def _ya_llamado(self, row_data):
        """
        Comprueba si el contacto ya ha sido llamado y gestionado.
        
        Los contactos sin respuesta registrada o que se marcaron como "No contesta" se consideran
        pendientes de llamar.
        
        Args:
            row_data (dict): Diccionario con los datos del contacto.
            
        Returns:
            bool: True si el contacto ya tiene una gestión finalizada, False si está pendiente.
        """
        respuesta = row_data.get("Respuesta")
        if respuesta is None or str(respuesta).strip() == "":
            return False  # sin respuesta: pendiente de llamar
        respuesta_str = str(respuesta).strip()
        if respuesta_str == RESPUESTA_NO_CONTESTA:
            return False  # "no contesta" pues pal final
        return True  # si tiene otra respuesta no le llammos otra vez

    def _ocupado_por_otro(self, row_data):
        """
        Verifica si otro operador tiene bloqueada la fila de este contacto actualmente.
        
        Args:
            row_data (dict): Diccionario con los datos del contacto.
            
        Returns:
            bool: True si el contacto está ocupado por otro operador y el bloqueo no ha expirado, False si está libre.
        """
        # devuelve true si otro compi esta llamando a este numero
        # usamos el valor de la cache que ya refrescamos al inicio de la operacion
        # PA NO REVENTAR EL SHAREPOINT A LECTURAS CADA MILISEGUNDO
        lock_val = row_data.get(COL_EN_LLAMADA)
        if lock_val is None or str(lock_val).strip() == "":
            return False  # via libre chavales
        if self._es_mi_bloqueo(lock_val):
            return False  # es mio, palante
        if self._bloqueo_expirado(lock_val):
            return False  # lleva mil años, me lo quedo
        return True  # no tocar, terreno ocupado

    def _sacar_telefono(self, row_data):
        """
        Extrae el primer número de teléfono móvil válido (TELEFONO1 o TELEFONO2) del contacto.
        
        Valida que sea un número móvil español (de 9 dígitos empezando por 6 o 7, o con prefijo 34).
        
        Args:
            row_data (dict): Diccionario con los datos del contacto.
            
        Returns:
            str: El número de teléfono normalizado de 9 u 11 dígitos, o None si no hay un móvil válido.
        """
        def telefono_valido(val):
            if val is None:
                return False
            # Dejar solo dígitos para validar longitud y contenido
            s = "".join(filter(str.isdigit, str(val)))
            if not s:
                return False
            
            # si tiene prefijo 
            if s.startswith("34"):
                # 11 dígitos en total (34 + 9 del móvil) y que sea móvil (6 o 7) jaja six seven
                return len(s) == 11 and s[2] in ("6", "7")
            if s[0] in ("6", "7"):
                return len(s) == 9
                
            return False

        tel1 = row_data.get("TELEFONO1")
        tel2 = row_data.get("TELEFONO2")
        if telefono_valido(tel1):
            return "".join(filter(str.isdigit, str(tel1)))
        if telefono_valido(tel2):
            return "".join(filter(str.isdigit, str(tel2)))
        return None

    def handle_call(self):
        """
        Manejador del evento del botón 'LLAMAR'.
        
        Refresca la caché local, busca el siguiente contacto elegible respetando
        los filtros, bloquea atómicamente la fila en el JSON para el operador actual,
        muestra la información en la ficha del contacto y abre el softphone para iniciar la llamada.
        """
        if not self.cache_filas:
            self._error("No hay filas en el Excel cargado.")
            return

        # refrescamos porsiaca antes de buscar a que este libre
        try:
            self._refrescar_cache()
        except Exception as e:
            self._error(f"No se pudo leer el estado actualizado del Excel:\n{e}")
            return

        # dos listas, los no_contesta los mandamos pal final pa q no molesten
        normales = []
        sin_contestar = []

        # contadores de diagnóstico
        _d_total = len(self.cache_filas)
        _d_filtro = 0
        _d_tel = 0
        _d_llamado = 0
        _d_ocupado = 0

        for row_data in self.cache_filas:
            if not self._pasa_filtros(row_data):
                _d_filtro += 1
                continue
            phone = self._sacar_telefono(row_data)
            if phone is None:
                _d_tel += 1
                continue
            if self._ya_llamado(row_data):
                _d_llamado += 1
                continue  # ya gestionada con otra respuesta, la saltamos para siempre
            if self._ocupado_por_otro(row_data):
                _d_ocupado += 1
                continue  # otro operador está en esa fila ahora mismo

            respuesta = row_data.get("Respuesta")
            respuesta_str = str(respuesta).strip() if respuesta is not None else ""
            if respuesta_str == RESPUESTA_NO_CONTESTA:
                sin_contestar.append(row_data)
            else:
                normales.append(row_data)

        # Candidatos en orden de prioridad
        candidatos = normales + sin_contestar

        if not candidatos:
            self._error(
                f"No se han encontrado contactos que cumplan los filtros.\n\n"
                f"--- Diagnóstico ---\n"
                f"Total filas en cache: {_d_total}\n"
                f"Descartadas por filtros (localidad/sexo/edad/resp): {_d_filtro}\n"
                f"Sin teléfono móvil válido: {_d_tel}\n"
                f"Ya llamados (tienen respuesta): {_d_llamado}\n"
                f"Ocupados por otro operador: {_d_ocupado}"
            )
            return

        # soltar lo anterior q teniamos trincado
        if self.fila_actual is not None:
            self._liberar_fila(self.fila_actual["data"])

        chosen = None
        for candidate in candidatos:
            # Intentamos bloquear de forma atómica en el Excel
            if self._bloquear_fila(candidate):
                chosen = candidate
                break
            else:
                # Si falló el bloqueo, alguien se adelantó. Al siguiente.
                continue

        if chosen is None:
            self._error("No se ha podido bloquear ningún contacto disponible.\n"
                        "Es posible que otros operadores los hayan ocupado justo ahora.")
            return

        phone = self._sacar_telefono(chosen)
        self.fila_actual = {"data": chosen, "phone": phone}
        self.snapshot_disco = {
            "Respuesta": chosen.get("Respuesta"),
            "Observaciones": chosen.get("Observaciones"),
        }
        self._mostrar_contacto(chosen)
        self._llamar_adb(phone)

    def handle_load_contact(self):
        """
        Manejador del evento del botón 'CARGAR CONTACTO'.
        
        Funciona de manera idéntica a `handle_call` (aplica filtros y bloquea
        un contacto elegible), pero carga los datos en la interfaz sin iniciar
        la llamada automáticamente (sin abrir el softphone).
        """
        if not self.cache_filas:
            self._error("No hay filas en el Excel cargado.")
            return

        # volver a leer q igual otro ya ha pillao el contacto este
        try:
            self._refrescar_cache()
        except Exception as e:
            self._error(f"No se pudo leer el estado actualizado del Excel:\n{e}")
            return

        normales = []
        sin_contestar = []

        # contadores de diagnóstico
        _d_total = len(self.cache_filas)
        _d_filtro = 0
        _d_llamado = 0
        _d_ocupado = 0

        # Identificar candidatos respetando bloqueos y respuestas
        for row_data in self.cache_filas:
            if not self._pasa_filtros(row_data):
                _d_filtro += 1
                continue
            if self._ya_llamado(row_data):
                _d_llamado += 1
                continue  # ya tiene respuesta, lo saltamos
            if self._ocupado_por_otro(row_data):
                _d_ocupado += 1
                continue
            
            respuesta = row_data.get("Respuesta")
            respuesta_str = str(respuesta).strip() if respuesta is not None else ""
            if respuesta_str == RESPUESTA_NO_CONTESTA:
                sin_contestar.append(row_data)
            else:
                normales.append(row_data)

        candidatos = normales + sin_contestar

        if not candidatos:
            self._error(
                f"No se han encontrado contactos que cumplan los filtros.\n\n"
                f"--- Diagnóstico ---\n"
                f"Total filas en cache: {_d_total}\n"
                f"Descartadas por filtros (localidad/sexo/edad/resp): {_d_filtro}\n"
                f"Ya llamados (tienen respuesta): {_d_llamado}\n"
                f"Ocupados por otro operador: {_d_ocupado}"
            )
            return

        # liberamos el anterior si lo hubiera
        if self.fila_actual is not None:
            self._liberar_fila(self.fila_actual["data"])

        chosen = None
        for candidate in candidatos:
            if self._bloquear_fila(candidate):
                chosen = candidate
                break
            else:
                continue

        if chosen is None:
            self._error("No se ha podido cargar ningún contacto (ocupados por otros operadores).")
            return

        phone = self._sacar_telefono(chosen) or ""

        self.fila_actual = {"data": chosen, "phone": phone}
        self.snapshot_disco = {
            "Respuesta": chosen.get("Respuesta"),
            "Observaciones": chosen.get("Observaciones"),
        }
        self._mostrar_contacto(chosen)

    def _refrescar_cache(self):
        """
        Refresca selectivamente la caché en memoria leyendo el estado actual
        de las respuestas en el Excel y el estado de los bloqueos activos en el JSON.
        """
        # pa no comerse medio giga de ram recargando el excel entero
        # leemos a pelo nomas lo q nos importa
        if self.excel_path is None or self.sheet is None:
            return
        wb_fresh = load_workbook(self.excel_path, data_only=True)
        sheet_fresh = wb_fresh[self.sheet.title] if self.sheet is not None else wb_fresh.active

        col_respuesta = self.cols.get("Respuesta")

        # las respuestas se leen del Excel, que remedio
        for row_data in self.cache_filas:
            row_idx = row_data["_row"]
            if col_respuesta is not None:
                row_data["Respuesta"] = sheet_fresh.cell(row=row_idx, column=col_respuesta).value

        # los bloqueos se leen del JSON (el turbo boost definitivo)
        locks = self._leer_locks()
        for row_data in self.cache_filas:
            row_key = str(row_data["_row"])
            row_data[COL_EN_LLAMADA] = locks.get(row_key)

    def _bloquear_fila(self, row_data) -> bool:
        """
        Intenta adquirir el bloqueo de una fila via JSON.
        Retorna True si se consiguió bloquear, False si ya estaba ocupada.
        """
        if self.excel_path is None:
            return False
        row_idx = row_data["_row"]
        row_key = str(row_idx)
        lock_file = self._locks_path() + ".lock"

        try:
            # Adquirir lock del fichero JSON a lo bruto pero seguro
            start = time.time()
            while True:
                try:
                    fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(fd)
                    break
                except FileExistsError:
                    if time.time() - start > 10:
                        self._registrar_error(f"Timeout bloqueando JSON para fila {row_idx}, se quedó pillado")
                        return False
                    time.sleep(0.2)

            try:
                locks = self._leer_locks()

                # Comprobar si alguien se nos ha adelantado justo ahora (Race Condition check)
                actual_val = locks.get(row_key)
                if actual_val is not None and actual_val.strip() != "":
                    if not self._es_mi_bloqueo(actual_val) and not self._bloqueo_expirado(actual_val):
                        return False  # F, nos lo quitaron de las manos

                # Si llegamos aquí, la fila es nuestra, boom
                lock_val = self._valor_bloqueo()
                locks[row_key] = lock_val
                self._escribir_locks(locks)
            finally:
                try:
                    os.remove(lock_file)
                except (FileNotFoundError, PermissionError):
                    pass

            row_data[COL_EN_LLAMADA] = lock_val
            return True

        except Exception as e:
            self._registrar_error(f"Error al bloquear fila {row_idx}: {e}")
            return False

    def _liberar_fila(self, row_data):
        """
        Libera el bloqueo de una fila en el archivo JSON.
        
        Remueve la marca del operador actual permitiendo que otros operadores puedan
        acceder y gestionar este registro.
        
        Args:
            row_data (dict): Diccionario con los datos del contacto a liberar.
        """
        # quitamos la marquita nuestra del JSON para que los compis puedan darle caña
        if self.excel_path is None:
            return
        en_llamada = row_data.get(COL_EN_LLAMADA)
        if en_llamada is None or not self._es_mi_bloqueo(en_llamada):
            return
        row_key = str(row_data["_row"])
        lock_file = self._locks_path() + ".lock"
        try:
            start = time.time()
            while True:
                try:
                    fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(fd)
                    break
                except FileExistsError:
                    if time.time() - start > 10:
                        return
                    time.sleep(0.2)
            try:
                locks = self._leer_locks()
                if row_key in locks:
                    del locks[row_key]
                    self._escribir_locks(locks)
            finally:
                try:
                    os.remove(lock_file)
                except (FileNotFoundError, PermissionError):
                    pass
            row_data[COL_EN_LLAMADA] = None
        except Exception as e:
            self._registrar_error(f"Error al liberar bloqueo de fila: {e}")

    def _mostrar_contacto(self, row_data):
        """
        Muestra los detalles del contacto (nombre, apellido, teléfonos, correo,
        respuesta anterior y observaciones) en los componentes de la interfaz.
        
        Args:
            row_data (dict): Diccionario con la información del contacto.
        """
        nombre = row_data.get("NOMBRE", "")
        apellido = row_data.get("APELLIDO1", "")
        tel1 = row_data.get("TELEFONO1", "")
        tel2 = row_data.get("TELEFONO2", "")
        email = row_data.get("EMAIL", "")
        respuesta = row_data.get("Respuesta", "")
        observaciones = row_data.get("Observaciones", "")

        # Mostramos los datos en sus cajitas correspondientes
        self.nombre_label.setText(str(nombre) if nombre is not None else "-")
        self.apellido_label.setText(str(apellido) if apellido is not None else "-")
        
        # El teléfono bien clarito, si hay dos los ponemos con su separador
        tels = []
        if tel1: tels.append(str(tel1))
        if tel2: tels.append(str(tel2))
        self.telefono_label.setText(" / ".join(tels) if tels else "-")

        self.email_label.setText(str(email) if email is not None else "-")
        self.respuesta_combo.setEditText(str(respuesta) if respuesta is not None else "")
        self.observaciones_edit.setPlainText(str(observaciones) if observaciones is not None else "")

    def _llamar_adb(self, phone):
        """
        Lanza la llamada directamente en un dispositivo Android conectado por USB
        usando comandos ADB (adb shell am start -a android.intent.action.CALL -d tel:+NUMERO).
        
        Args:
            phone (str): Número de teléfono a marcar.
        """
        phone = str(phone).strip()
        if not phone.startswith("+"):
            if phone.startswith("34"):
                phone = "+" + phone
            else:
                phone = "+34" + phone

        # Buscamos la ruta de adb.exe
        import shutil
        adb_path = "adb"
        if not shutil.which("adb"):
            # Si no está en el PATH del proceso actual, buscamos en la carpeta de instalación de Winget
            local_appdata = os.getenv("LOCALAPPDATA", "")
            if local_appdata:
                winget_packages_dir = os.path.join(local_appdata, "Microsoft", "WinGet", "Packages")
                if os.path.exists(winget_packages_dir):
                    for root, dirs, files in os.walk(winget_packages_dir):
                        if "adb.exe" in files:
                            adb_path = os.path.join(root, "adb.exe")
                            break

        import subprocess
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0  # Ocultar la ventana de consola en Windows

            # Ejecutamos el comando de llamada ADB
            cmd = [adb_path, "shell", "am", "start", "-a", "android.intent.action.CALL", "-d", f"tel:{phone}"]
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                startupinfo=startupinfo,
                timeout=15
            )

            if result.returncode == 0 and "Error" not in result.stderr:
                self._info(f"Llamada iniciada en Android vía ADB al número: {phone}")
            else:
                error_msg = result.stderr or result.stdout
                self._error(
                    f"No se pudo iniciar la llamada por ADB.\n\n"
                    f"Detalles:\n{error_msg.strip()}\n\n"
                    f"Asegúrate de tener un móvil Android conectado por USB con la depuración USB activada."
                )
        except subprocess.TimeoutExpired:
            self._error(
                "La llamada por ADB superó el tiempo de espera.\n\n"
                "1. Comprueba la conexión del cable USB.\n"
                "2. Si es la primera vez que lo conectas al PC, mira la pantalla de tu móvil y acepta la ventana de autorización '¿Permitir depuración USB?'."
            )
        except FileNotFoundError:
            self._error(
                "No se encontró el comando 'adb'.\n\n"
                "Para usar este método necesitas instalar las herramientas de plataforma de Android (platform-tools) "
                "y añadir 'adb' al PATH de variables de entorno del sistema."
            )
        except Exception as e:
            self._error(f"Error inesperado al intentar llamar por ADB:\n{e}")

    def _abrir_softphone(self, phone):
        """
        Inicia el softphone de Netelip (SIP/WebRTC).
        
        Genera dinámicamente un archivo HTML/JS en un servidor HTTP local
        temporal y abre el navegador por defecto del sistema apuntando a él.
        
        Args:
            phone (str): Número de teléfono a marcar.
        """
        phone = str(phone).strip()
        if not phone.startswith("+"):
            if phone.startswith("34"):
                phone = "+" + phone
            else:
                phone = "+34" + phone

        # Cargar las credenciales de Netelip desde .env
        sip_user = os.getenv("SIP_USER")
        sip_password = os.getenv("SIP_PASSWORD")
        sip_server = os.getenv("SIP_SERVER", "wss://sip-eu.netelip.com")
        sip_domain = os.getenv("SIP_DOMAIN", "sip.netelip.com")

        if not sip_user or not sip_password:
            self._error(
                "Faltan credenciales de Netelip en el archivo .env\n\n"
                "Por favor, configura SIP_USER y SIP_PASSWORD en el archivo .env de la carpeta de la aplicación."
            )
            return

        try:
            import threading
            import http.server
            import socketserver

            html_template = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>SocioGraph Softphone V5 (Netelip)</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jssip/3.10.0/jssip.min.js"></script>
    <style>
        body { 
            font-family: 'Outfit', 'Inter', sans-serif; 
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); 
            color: #f8fafc; 
            display: flex; 
            flex-direction: column; 
            align-items: center; 
            justify-content: center; 
            height: 100vh; 
            margin: 0; 
            gap: 24px;
        }
        .container {
            background: rgba(30, 41, 59, 0.7);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            padding: 40px;
            border-radius: 20px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3), 0 8px 10px -6px rgba(0, 0, 0, 0.3);
            display: flex;
            flex-direction: column;
            align-items: center;
            width: 320px;
            text-align: center;
        }
        #status { 
            font-size: 16px; 
            color: #38bdf8; 
            font-weight: 500;
            margin-top: 10px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .status-dot {
            width: 8px;
            height: 8px;
            background-color: #38bdf8;
            border-radius: 50%;
            display: inline-block;
        }
        .status-dot.pulsing {
            animation: pulse 1.5s infinite alternate;
        }
        @keyframes pulse {
            0% { transform: scale(0.8); opacity: 0.5; }
            100% { transform: scale(1.3); opacity: 1; box-shadow: 0 0 10px #38bdf8; }
        }
        #phone { 
            font-size: 28px; 
            font-weight: 700; 
            color: #ffffff; 
            margin: 15px 0 25px 0;
            letter-spacing: 0.02em;
        }
        #btn-close-manual { 
            background: #475569; 
            color: #f1f5f9; 
            border: none; 
            padding: 14px 28px; 
            border-radius: 12px; 
            font-weight: 600; 
            cursor: not-allowed; 
            opacity: 0.5;
            font-size: 14px;
            transition: all 0.3s ease;
            margin-bottom: 20px;
            width: 100%;
        }
        button.action-btn { 
            padding: 16px 32px; 
            border-radius: 50px; 
            border: none; 
            font-size: 16px; 
            font-weight: 600;
            cursor: pointer; 
            transition: all 0.2s ease;
            width: 80%;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -2px rgba(0, 0, 0, 0.1);
        }
        button.action-btn:active {
            transform: scale(0.97);
        }
        #btn-call { 
            background: linear-gradient(135deg, #10b981 0%, #059669 100%); 
            color: white; 
            display: none;
        }
        #btn-call:hover {
            background: linear-gradient(135deg, #34d399 0%, #059669 100%);
            box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.3);
        }
        #btn-hang { 
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); 
            color: white; 
            display: inline-block;
        }
        #btn-hang:hover {
            background: linear-gradient(135deg, #f87171 0%, #dc2626 100%);
            box-shadow: 0 10px 15px -3px rgba(239, 68, 68, 0.3);
        }
    </style>
</head>
<body>
    <audio id="remoteAudio" autoplay></audio>
    <div class="container">
        <button id="btn-close-manual" onclick="window.close()" disabled>✖ ESPERANDO FIN DE LLAMADA</button>
        <div id="status"><span class="status-dot pulsing"></span>Iniciando...</div>
        <div id="phone">[[PHONE]]</div>
        <button id="btn-call" class="action-btn" onclick="makeCall()">Llamar</button>
        <button id="btn-hang" class="action-btn" onclick="hangUp()">Colgar</button>
    </div>

    <script>
        var socket = new JsSIP.WebSocketInterface('[[SIP_SERVER]]');
        
        var configuration = {
            sockets  : [ socket ],
            uri      : 'sip:[[SIP_USER]]@[[SIP_DOMAIN]]',
            password : '[[SIP_PASSWORD]]',
            register : true
        };

        var ua = new JsSIP.UA(configuration);
        var session = null;

        function updateStatus(text, dotClass = 'pulsing', color = '#38bdf8') {
            var statusEl = document.getElementById('status');
            statusEl.innerHTML = '<span class="status-dot ' + dotClass + '" style="background-color: ' + color + '"></span>' + text;
            statusEl.style.color = color;
        }

        function finishSoftphone() {
            updateStatus('Llamada finalizada', '', '#94a3b8');
            document.getElementById('btn-call').style.display = 'none';
            document.getElementById('btn-hang').style.display = 'none';
            
            var closeBtn = document.getElementById('btn-close-manual');
            closeBtn.disabled = false;
            closeBtn.style.opacity = '1';
            closeBtn.style.background = 'linear-gradient(135deg, #10b981 0%, #059669 100%)';
            closeBtn.style.cursor = 'pointer';
            closeBtn.textContent = ' X CERRAR PESTAÑA ';

            setTimeout(function() {
                window.open('', '_self', ''); 
                window.close();
            }, 1500);
        }

        ua.on('connected', function() {
            updateStatus('Conectado al servidor SIP');
        });

        ua.on('disconnected', function() {
            updateStatus('Desconectado', '', '#ef4444');
        });

        ua.on('registered', function() {
            updateStatus('Listo (Registrado)', '', '#10b981');
            makeCall();
        });

        ua.on('registrationFailed', function(data) {
            updateStatus('Fallo de Registro: ' + data.cause, '', '#ef4444');
            console.error('Registration failed:', data);
        });

        ua.on('newRTCSession', function(data) {
            session = data.session;
            
            if (session.direction === 'outgoing') {
                session.on('peerconnection', function(pcData) {
                    var pc = pcData.peerconnection;
                    
                    var attachStream = function(stream) {
                        console.log('Attaching stream...');
                        var remoteAudio = document.getElementById('remoteAudio');
                        if (remoteAudio) {
                            remoteAudio.srcObject = stream;
                        }
                    };

                    pc.addEventListener('track', function(e) {
                        if (e.streams && e.streams[0]) {
                            attachStream(e.streams[0]);
                        }
                    });

                    // Fallback para addstream
                    pc.onaddstream = function(e) {
                        if (e.stream) {
                            attachStream(e.stream);
                        }
                    };
                });

                session.on('connecting', function() {
                    updateStatus('Conectando...', 'pulsing', '#38bdf8');
                });

                session.on('progress', function() {
                    updateStatus('Llamando...', 'pulsing', '#38bdf8');
                });

                session.on('accepted', function() {
                    updateStatus('En llamada ✓', 'pulsing', '#10b981');
                    document.getElementById('btn-call').style.display = 'none';
                    document.getElementById('btn-hang').style.display = 'inline-block';
                });

                session.on('failed', function(e) {
                    console.warn('Call failed:', e);
                    updateStatus('Llamada fallida: ' + e.cause, '', '#ef4444');
                    finishSoftphone();
                });

                session.on('ended', function(e) {
                    updateStatus('Llamada terminada', '', '#94a3b8');
                    finishSoftphone();
                });
            }
        });

        function makeCall() {
            if (session) return;
            
            updateStatus('Llamando...', 'pulsing', '#38bdf8');
            document.getElementById('btn-call').style.display = 'none';
            document.getElementById('btn-hang').style.display = 'inline-block';
            
            var eventHandlers = {
                // Event handlers are handled globally on newRTCSession
            };

            var options = {
                'eventHandlers'    : eventHandlers,
                'mediaConstraints' : { 'audio': true, 'video': false }
            };

            try {
                ua.call('sip:[[PHONE]]@[[SIP_DOMAIN]]', options);
            } catch (err) {
                console.error('Error calling:', err);
                updateStatus('Error al llamar: ' + err.message, '', '#ef4444');
                finishSoftphone();
            }
        }

        function hangUp() {
            if (session) {
                session.terminate();
            }
        }

        ua.start();
    </script>
</body>
</html>"""

            html_content = html_template.replace("[[SIP_SERVER]]", sip_server)
            html_content = html_content.replace("[[SIP_USER]]", sip_user)
            html_content = html_content.replace("[[SIP_PASSWORD]]", sip_password)
            html_content = html_content.replace("[[SIP_DOMAIN]]", sip_domain)
            html_content = html_content.replace("[[PHONE]]", phone)

            html_bytes = html_content.encode("utf-8")

            class SoftphoneHandler(http.server.BaseHTTPRequestHandler):
                def do_GET(self):
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.send_header("Content-Length", str(len(html_bytes)))
                    self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                    self.send_header("Pragma", "no-cache")
                    self.send_header("Expires", "0")
                    self.end_headers()
                    self.wfile.write(html_bytes)
                def log_message(self, format, *args):
                    pass

            server = socketserver.TCPServer(("127.0.0.1", 0), SoftphoneHandler)
            port = server.server_address[1]
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            webbrowser.open(f"http://127.0.0.1:{port}")
        except Exception as e:
            self._error(f"Error al iniciar la llamada:\n{e}")

    def handle_comprobar(self):
        """
        Manejador del evento del botón 'Comprobar'.
        
        Permite seleccionar un archivo Excel de histórico de llamadas, normaliza
        los datos de nombre y apellido del contacto activo e inicia la búsqueda
        asíncrona en segundo plano usando un `HistoryWorker`.
        """
        if self.fila_actual is None:
            self._error("Primero carga un contacto para poder comprobar.")
            return

        historico_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Excel Histórico",
            "",
            "Archivos Excel (*.xlsx)",
        )
        if not historico_path:
            return
        if not os.path.exists(historico_path):
            self._error("El archivo Histórico especificado no existe.")
            return

        # Preparar datos a buscar
        row_data = self.fila_actual["data"]
        # Usamos la misma lógica de normalización que el worker para comparar manzanas con manzanas
        import unicodedata
        def _norm_local(s):
            s = unicodedata.normalize("NFD", str(s or ""))
            s = "".join(c for c in s if unicodedata.category(c) != "Mn")
            return " ".join(s.upper().split())

        nombre_buscar   = _norm_local(row_data.get("NOMBRE"))
        apellido_buscar = _norm_local(row_data.get("APELLIDO1"))

        # UI: Desactivar botón y cambiar cursor
        self.comprobar_btn.setEnabled(False)
        self.comprobar_btn.setText("Comprobando...")
        QApplication.setOverrideCursor(Qt.WaitCursor)

        # Lanzar hilo de búsqueda (ahora solo nombre y apellido)
        self.worker = HistoryWorker(historico_path, nombre_buscar, apellido_buscar)
        self.worker.finished.connect(self._on_comprobar_finished)
        self.worker.start()

    def _on_comprobar_finished(self, coincidencias, error_msg):
        """
        Manejador del evento de finalización de la búsqueda del historial.
        
        Restaura los controles visuales y carga los resultados encontrados en la
        lista lateral. Si coincide el teléfono, los resalta visualmente en color rojo.
        
        Args:
            coincidencias (list): Lista de diccionarios con contactos coincidentes.
            error_msg (str): Mensaje de error, si ocurrió alguno durante la búsqueda.
        """
        # UI: Restaurar botón y cursor
        self.comprobar_btn.setEnabled(True)
        self.comprobar_btn.setText("Comprobar")
        QApplication.restoreOverrideCursor()

        if error_msg:
            self._error(f"Error al comprobar el Histórico:\n{error_msg}")
            return

        # Limpiar y rellenar la lista lateral con los resultados del historial
        self.lista_especial.clear()
        
        if not coincidencias:
            QMessageBox.information(self, "Historial", "No se han encontrado coincidencias por nombre y apellidos.")
            return

        # Añadir un separador visual o título
        header_item = QListWidgetItem("─── COINCIDENCIAS HISTÓRICO ───")
        header_item.setFlags(Qt.NoItemFlags) # no seleccionable
        header_item.setForeground(Qt.blue)
        self.lista_especial.addItem(header_item)

        for c in coincidencias:
            # Construimos la línea definitiva según los deseos del jefe
            campos_orden = [
                c.get("EMPRESA CAPTACION", "-"),
                c.get("NOMBRE", "-"),
                c.get("APELLIDO") or c.get("APELLIDO1") or c.get("APELLIDO 1") or "-",
                c.get("DNI", "-"),
                c.get("FECHA EN LA QUE PARTICIPO", "-"),
                c.get("ESTUDIO", "-"),
                c.get("MARCA", "-"),
                c.get("PARTICIPANTE", "-"),
                c.get("TELEFONO", "-")
            ]
            
            # Todo en una línea separada por barritas para que no abulte
            texto_linea = " || ".join(campos_orden)
            item = QListWidgetItem(texto_linea)
            item.setData(Qt.UserRole, c)
            
            # Alerta roja: si el teléfono del historial es igual al de ahora, ¡CUIDADO!
            tel_hist = campos_orden[-1]
            row_data_actual = self.fila_actual["data"]
            tel1_act = str(row_data_actual.get("TELEFONO1") or "").strip()
            tel2_act = str(row_data_actual.get("TELEFONO2") or "").strip()
            
            if tel_hist != "-" and (tel_hist in (tel1_act, tel2_act)):
                item.setForeground(Qt.red)
                font = item.font()
                font.setBold(True)
                item.setFont(font)
            
            self.lista_especial.addItem(item)

        QMessageBox.information(
            self, 
            "Historial", 
            f"Se han encontrado {len(coincidencias)} coincidencia(s).\n\n"
            "Los detalles se han cargado en la lista lateral."
        )

    def handle_save(self):
        """
        Manejador del evento del botón 'GUARDAR'.
        
        Verifica que el contacto y el libro de Excel sigan cargados y bloqueados
        por este operador en el JSON. Adquiere un semáforo de red de forma segura
        para modificar el archivo Excel escribiendo la respuesta y observaciones del operador,
        guarda el archivo Excel, y finalmente libera el bloqueo atómico.
        """
        if self.fila_actual is None:
            self._error("Selecciona un contacto")
            return
        if self.workbook is None or self.sheet is None or self.excel_path is None:
            self._error("No hay libro de Excel cargado")
            return

        row_data = self.fila_actual["data"]
        row_idx = row_data["_row"]
        try:
            respuesta_text = self.respuesta_combo.currentText()
            observaciones_text = self.observaciones_edit.toPlainText()
            col_resp = self.cols["Respuesta"]
            col_obs = self.cols["Observaciones"]
            # comprobar que nuestro bloqueo sigue en el JSON vivito y coleando
            locks = self._leer_locks()
            row_key = str(row_idx)
            current_lock = locks.get(row_key)
            if current_lock is None or not self._es_mi_bloqueo(current_lock):
                self._error(
                    "Ojo que alguien nos ha pisado la fila, no se puede guardar esto!!"
                )
                return

            with excel_file_lock(self.excel_path):
                wb = load_workbook(self.excel_path, data_only=False)
                sheet = wb[self.sheet.title] if self.sheet is not None else wb.active

                # si nadie nos jodio somo los reyes, guardamos el textooooo
                sheet.cell(row=row_idx, column=col_resp, value=respuesta_text)
                sheet.cell(row=row_idx, column=col_obs, value=observaciones_text)
                wb.save(self.excel_path)

            # liberar bloqueo del JSON
            self._liberar_fila(row_data)

            # Actualizar referencias internas y cache
            self.workbook = wb
            self.sheet = sheet
            try:
                self.mtime_excel = os.path.getmtime(self.excel_path)
            except OSError:
                self.mtime_excel = None

            row_data["Respuesta"] = respuesta_text
            row_data["Observaciones"] = observaciones_text
            row_data[COL_EN_LLAMADA] = None  # liberado en cache local también
            self.snapshot_disco = None

            # Limpiar fila_actual pa que no se pueda guardar dos veces 
            self.fila_actual = None

            self._info("Cambios guardados en el Excel")
        except ExcelLockTimeout as e:
            self._error(str(e))
        except Exception as e:
            self._error(f"Error al guardar en el Excel:\n{e}")

    def closeEvent(self, event):
        """
        Manejador del evento de cierre de la ventana de PyQt5.
        
        Asegura la liberación del bloqueo de la fila activa en el JSON si se
        cierra la ventana abruptamente sin guardar cambios.
        
        Args:
            event (QCloseEvent): Evento de cierre de ventana.
        """
        
        if self.fila_actual is not None:
            self._liberar_fila(self.fila_actual["data"])
        super().closeEvent(event)

    def _manejar_error(self, error, contexto=""):
        """
        Maneja y propaga una excepción, registrándola en el log y mostrando
        un mensaje crítico de error en la interfaz.
        
        Args:
            error (Exception): La excepción capturada.
            contexto (str, opcional): Contexto o etapa en que ocurrió el error.
        """
        mensaje = f"{type(error).__name__}: {error}"
        if contexto:
            mensaje = f"[{contexto}] {mensaje}"
        self._registrar_error(mensaje)
        self._error(mensaje)

    def _registrar_error(self, mensaje):
        """
        Escribe un mensaje de error en el archivo de log 'errores_app.log'
        situado junto a la aplicación.
        
        Args:
            mensaje (str): Contenido a escribir en el log.
        """
        try:
            log_path = os.path.join(os.path.dirname(sys.argv[0]), "errores_app.log")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(mensaje + "\n")
        except Exception:
            pass

    def _error(self, message):
        """
        Muestra un diálogo modal de error crítico en la interfaz.
        
        Args:
            message (str): Mensaje de error a mostrar.
        """
        QMessageBox.critical(self, "Error", message)

    def _info(self, message):
        """
        Muestra un diálogo modal de información (Aceptado) en la interfaz.
        
        Args:
            message (str): Mensaje de información a mostrar.
        """
        QMessageBox.information(self, "Aceptado", message)



class SoftphoneWindow(QWidget):
    """
    Ventana alternativa integrada en PyQt5 para el softphone web.
    
    Utiliza QWebEngineView para renderizar la interfaz web de Netelip (SIP/WebRTC)
    dentro de un widget de la propia aplicación en lugar de abrir el navegador externo.
    """

    def __init__(self, token, phone, parent=None):
        """
        Inicializa la ventana del softphone y carga el contenido HTML del cliente SIP.
        
        Args:
            token (str): Token de autenticación (no utilizado actualmente).
            phone (str): Número de teléfono a marcar.
            parent (QWidget, opcional): Widget padre en la jerarquía de Qt.
        """
        super().__init__(parent)
        self.setWindowTitle("Softphone")
        self.resize(340, 460)
        from PyQt5.QtWebEngineWidgets import QWebEngineSettings
        self.web_view = QWebEngineView()
        self.web_view.settings().setAttribute(
            QWebEngineSettings.ScreenCaptureEnabled, True
        )
        layout = QVBoxLayout(self)
        layout.addWidget(self.web_view)
        self.web_view.setHtml(self._construir_html(phone), QUrl("about:blank"))

    def _construir_html(self, phone):
        """
        Genera el código HTML y JavaScript (JsSIP) necesario para conectar con Netelip
        y realizar la llamada WebRTC.
        
        Args:
            phone (str): Número de teléfono a marcar.
            
        Returns:
            str: Cadena de texto con el código HTML/JS completo.
        """
        sip_user = os.getenv("SIP_USER", "")
        sip_password = os.getenv("SIP_PASSWORD", "")
        sip_server = os.getenv("SIP_SERVER", "wss://sip-eu.netelip.com")
        sip_domain = os.getenv("SIP_DOMAIN", "sip.netelip.com")
        
        return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>SocioGraph Softphone V5 (Netelip)</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jssip/3.10.0/jssip.min.js"></script>
  <style>
    body {{ font-family: Arial, sans-serif; background: #1a1a2e; color: #eee;
           display: flex; flex-direction: column; align-items: center;
           justify-content: center; height: 100vh; margin: 0; gap: 16px; }}
    #status {{ font-size: 13px; color: #a0aec0; }}
    #phone  {{ font-size: 22px; font-weight: bold; color: #90cdf4; }}
    button  {{ padding: 12px 28px; border: none; border-radius: 50px;
               font-size: 15px; font-weight: bold; cursor: pointer; }}
    #btn-hang {{ background: #fc8181; color: white; display: none; }}
    #btn-call {{ background: #48bb78; color: white; }}
  </style>
</head>
<body>
  <audio id="remoteAudio" autoplay></audio>
  <div id="status">Iniciando...</div>
  <div id="phone">{phone}</div>
  <button id="btn-call" onclick="makeCall()"> Llamar</button>
  <button id="btn-hang" onclick="hangUp()"> Colgar</button>
 <script>
    var socket = new JsSIP.WebSocketInterface('{sip_server}');
    
    var configuration = {{
        sockets  : [ socket ],
        uri      : 'sip:{sip_user}@{sip_domain}',
        password : '{sip_password}',
        register : true
    }};

    var ua = new JsSIP.UA(configuration);
    var session = null;

    function finishSoftphone() {{
        document.getElementById('status').textContent = 'Llamada finalizada';
        setTimeout(function() {{
            window.open('', '_self', ''); 
            window.close();
        }}, 1500);
    }}

    ua.on('registered', function() {{
      document.getElementById("status").textContent = "Listo";
      makeCall();
    }});

    ua.on('registrationFailed', function(data) {{
      document.getElementById("status").textContent = "Fallo de Registro: " + data.cause;
    }});

    ua.on('newRTCSession', function(data) {{
      session = data.session;
      if (session.direction === 'outgoing') {{
        session.on('peerconnection', function(pcData) {{
          var pc = pcData.peerconnection;
          pc.addEventListener('track', function(e) {{
            var remoteAudio = document.getElementById('remoteAudio');
            if (remoteAudio && e.streams && e.streams[0]) {{
              remoteAudio.srcObject = e.streams[0];
            }}
          }});
          pc.onaddstream = function(e) {{
            var remoteAudio = document.getElementById('remoteAudio');
            if (remoteAudio && e.stream) {{
              remoteAudio.srcObject = e.stream;
            }}
          }};
        }});

        session.on('connecting', function() {{
          document.getElementById("status").textContent = "Conectando...";
        }});
        session.on('progress', function() {{
          document.getElementById("status").textContent = "Llamando...";
        }});
        session.on('accepted', function() {{
          document.getElementById("status").textContent = "En llamada";
          document.getElementById("btn-call").style.display = "none";
          document.getElementById("btn-hang").style.display = "inline-block";
        }});
        session.on('failed', function(e) {{
          document.getElementById("status").textContent = "Llamada fallida: " + e.cause;
          finishSoftphone();
        }});
        session.on('ended', function() {{
          document.getElementById("status").textContent = "Llamada terminada";
          finishSoftphone();
        }});
      }}
    }});

    function makeCall() {{
      if (session) return;
      var options = {{
        'mediaConstraints' : {{ 'audio': true, 'video': false }}
      }};
      ua.call('sip:{phone}@{sip_domain}', options);
    }}

    function hangUp() {{
      if (session) {{
        session.terminate();
      }}
    }}

    ua.start();
  </script>
</body>
</html>"""


def main():
    """
    Punto de entrada principal de la aplicación.
    
    Instancia la aplicación QApplication, inicializa y muestra la ventana principal
    CallApp, e inicia el bucle de eventos de Qt.
    """
    app = QApplication(sys.argv)
    window = CallApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()