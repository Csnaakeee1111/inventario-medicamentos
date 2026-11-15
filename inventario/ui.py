"""Ventana principal del inventario (PySide6)."""
from __future__ import annotations

import sys
from datetime import date, timedelta
from pathlib import Path

from PySide6.QtCore import QDate, QObject, QStandardPaths, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont, QKeySequence
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QCompleter, QDateEdit, QDialog,
    QDialogButtonBox, QFileDialog, QFormLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPlainTextEdit, QProgressDialog, QPushButton, QSpinBox,
    QStyledItemDelegate, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)

from . import actualizador, excel_vivo
from .db import (
    DIAS, DIAS_CORTOS, MESES, SIN_LOTE, ErrorInventario, Inventario, carpeta_datos, dias_habiles,
    es_habil, lunes_de, texto_fecha, ultimo_habil_del_mes,
)
from .exportar import Hoja, a_excel, a_pdf
from .version import NOMBRE_APP, VERSION

VERDE = "E2F0D9"
ROJO = "FBE2E2"
NARANJA = "FCE4CC"
AMARILLO = "FFF2CC"
GRIS = "EEEEEE"
HOY_COLOR = "D9E8FF"
MORADO = excel_vivo.MORADO

ESTILO = """
QWidget { color: #1B1F24; }
QMainWindow, QDialog { background: #F4F6F9; }
QTabWidget::pane { border: 1px solid #C9D3DF; background: white; border-radius: 6px; }
QTabBar::tab { padding: 8px 18px; font-size: 11pt; background: #E3E9F0; border-top-left-radius: 6px;
               border-top-right-radius: 6px; margin-right: 2px; }
QTabBar::tab:selected { background: white; color: #1F4E79; font-weight: bold; }
QPushButton { padding: 6px 14px; border-radius: 5px; background: #FFFFFF; border: 1px solid #B8C4D2; }
QPushButton:hover { background: #E8F0FA; }
QPushButton[principal="true"] { background: #1F4E79; color: white; border: none; font-weight: bold; }
QPushButton[principal="true"]:hover { background: #2B6399; }
QTableWidget { gridline-color: #D5DCE4; selection-background-color: #BBD3EE; selection-color: black; }
QHeaderView::section { background: #1F4E79; color: white; padding: 5px; border: none;
                       border-right: 1px solid #3A6A97; font-weight: bold; }
QLineEdit, QComboBox, QDateEdit, QSpinBox, QPlainTextEdit { padding: 4px; border: 1px solid #B8C4D2;
                       border-radius: 4px; background: white; }
QLabel#titulo { font-size: 14pt; font-weight: bold; color: #1F4E79; }
QLabel#ayuda { color: #5A6B7D; }
"""


# ── Utilidades ──────────────────────────────────────────────────────────────

def a_qdate(d: date) -> QDate:
    return QDate(d.year, d.month, d.day)


def de_qdate(q: QDate) -> date:
    return date(q.year(), q.month(), q.day())


def celda(texto="", editable=False, alinear=None, color: str | None = None, dato=None,
          negrita=False) -> QTableWidgetItem:
    it = QTableWidgetItem("" if texto is None else str(texto))
    flags = Qt.ItemIsSelectable | Qt.ItemIsEnabled
    if editable:
        flags |= Qt.ItemIsEditable
    it.setFlags(flags)
    if isinstance(texto, int) and alinear is None:
        alinear = Qt.AlignRight | Qt.AlignVCenter
    if alinear is not None:
        it.setTextAlignment(alinear)
    if color:
        it.setBackground(QColor("#" + color))
    if dato is not None:
        it.setData(Qt.UserRole, dato)
    if negrita:
        f = it.font()
        f.setBold(True)
        it.setFont(f)
    return it


def tabla_basica(encabezados: list[str]) -> QTableWidget:
    t = QTableWidget(0, len(encabezados))
    t.setHorizontalHeaderLabels(encabezados)
    t.verticalHeader().setVisible(False)
    t.setAlternatingRowColors(True)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    t.horizontalHeader().setStretchLastSection(True)
    t.setWordWrap(False)
    return t


def boton(texto: str, accion, principal=False) -> QPushButton:
    b = QPushButton(texto)
    if principal:
        b.setProperty("principal", True)
    b.clicked.connect(accion)
    return b


def color_vencimiento(iso: str | None) -> str | None:
    if not iso:
        return None
    dias = (date.fromisoformat(iso) - date.today()).days
    if dias < 0:
        return ROJO
    if dias <= 30:
        return NARANJA
    if dias <= 90:
        return AMARILLO
    return None


def carpeta_documentos() -> Path:
    base = Path(QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation) or Path.home())
    carpeta = base / "Inventario Medicamentos"
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def hoja_de_tabla(t: QTableWidget, titulo: str, subtitulo="", nombre_hoja="Datos") -> Hoja:
    """Convierte lo que se ve en una tabla en una Hoja exportable (con colores)."""
    enc = [t.horizontalHeaderItem(c).text().replace("\n", " ") for c in range(t.columnCount())
           if not t.isColumnHidden(c)]
    cols = [c for c in range(t.columnCount()) if not t.isColumnHidden(c)]
    filas, colores = [], {}
    for r in range(t.rowCount()):
        fila = []
        for i, c in enumerate(cols):
            it = t.item(r, c)
            texto = it.text() if it else ""
            try:
                valor = int(texto) if texto.lstrip("-").isdigit() else texto
            except ValueError:
                valor = texto
            fila.append(valor)
            if it and it.background().style() != Qt.NoBrush:
                colores[(r, i)] = it.background().color().name()[1:].upper()
        filas.append(fila)
    return Hoja(titulo, enc, filas, subtitulo=subtitulo, nombre_hoja=nombre_hoja, colores=colores)


class DelegadoNumero(QStyledItemDelegate):
    """Edita celdas con un cuadro de números (sin letras ni negativos)."""

    def createEditor(self, parent, option, index):
        s = QSpinBox(parent)
        s.setRange(0, 9_999_999)
        s.setButtonSymbols(QSpinBox.NoButtons)
        s.setAlignment(Qt.AlignRight)
        return s

    def setEditorData(self, editor, index):
        texto = index.data() or "0"
        editor.setValue(int(texto) if texto.isdigit() else 0)
        editor.selectAll()

    def setModelData(self, editor, model, index):
        editor.interpretText()
        model.setData(index, str(editor.value()) if editor.value() else "")


class BuscadorMedicamento(QLineEdit):
    """Campo con autocompletado: se escribe una parte del nombre y aparecen los que coinciden."""

    def __init__(self, inv: Inventario, parent=None):
        super().__init__(parent)
        self.inv = inv
        self.setPlaceholderText("Escribe para buscar el medicamento…")
        self.completador = QCompleter(self)
        self.completador.setCaseSensitivity(Qt.CaseInsensitive)
        self.completador.setFilterMode(Qt.MatchContains)
        self.completador.setMaxVisibleItems(12)
        self.setCompleter(self.completador)
        self.recargar()

    def recargar(self):
        from PySide6.QtCore import QStringListModel
        self.completador.setModel(QStringListModel(self.inv.nombres_medicamentos(), self.completador))

    def medicamento(self):
        return self.inv.medicamento_por_nombre(self.text())


# ── Diálogos ────────────────────────────────────────────────────────────────

class DialogoMedicamento(QDialog):
    def __init__(self, inv: Inventario, medicamento_id: int | None = None, nombre="", parent=None):
        super().__init__(parent)
        self.inv, self.medicamento_id = inv, medicamento_id
        self.setWindowTitle("Editar medicamento" if medicamento_id else "Nuevo medicamento")
        self.setMinimumWidth(460)
        self.nombre = QLineEdit(nombre)
        self.nombre.setPlaceholderText("Ej.: Paracetamol 500 mg")
        self.descripcion = QPlainTextEdit()
        self.descripcion.setPlaceholderText("Para qué sirve, concentración, indicaciones…")
        self.descripcion.setFixedHeight(80)
        self.presentacion = QLineEdit()
        self.presentacion.setPlaceholderText("Tableta, jarabe 120 ml, ampolla…")
        self.minimo = QSpinBox()
        self.minimo.setRange(0, 999_999)
        self.minimo.setToolTip("Si la existencia baja de este número, se marca en rojo.")
        if medicamento_id:
            m = inv.medicamento(medicamento_id)
            self.nombre.setText(m["nombre"])
            self.descripcion.setPlainText(m["descripcion"])
            self.presentacion.setText(m["presentacion"])
            self.minimo.setValue(m["minimo"])
        f = QFormLayout(self)
        f.addRow("Medicamento:", self.nombre)
        f.addRow("Descripción:", self.descripcion)
        f.addRow("Presentación:", self.presentacion)
        f.addRow("Existencia mínima:", self.minimo)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Save).setText("Guardar")
        bb.button(QDialogButtonBox.Cancel).setText("Cancelar")
        bb.accepted.connect(self.guardar)
        bb.rejected.connect(self.reject)
        f.addRow(bb)

    def guardar(self):
        try:
            self.medicamento_id = self.inv.guardar_medicamento(
                self.nombre.text(), self.descripcion.toPlainText(), self.presentacion.text(),
                self.minimo.value(), self.medicamento_id)
        except ErrorInventario as e:
            QMessageBox.warning(self, "No se pudo guardar", str(e))
            return
        self.accept()


class DialogoMovimiento(QDialog):
    """Registrar una entrada o salida rápida (crea el lote si es nuevo)."""

    def __init__(self, inv: Inventario, fecha: date, tipo="ENTRADA", medicamento="", parent=None):
        super().__init__(parent)
        self.inv = inv
        self.setWindowTitle("Registrar entrada / salida")
        self.setMinimumWidth(480)
        self.tipo = QComboBox()
        self.tipo.addItem("Entrada (llegó medicamento)", "ENTRADA")
        self.tipo.addItem("Salida (se entregó / usó)", "SALIDA")
        self.tipo.setCurrentIndex(0 if tipo == "ENTRADA" else 1)
        self.medicamento = BuscadorMedicamento(inv)
        self.medicamento.setText(medicamento)
        nuevo = QPushButton("+ Nuevo")
        nuevo.setToolTip("Agregar un medicamento que todavía no está en el catálogo")
        nuevo.clicked.connect(self.nuevo_medicamento)
        fila_med = QHBoxLayout()
        fila_med.addWidget(self.medicamento, 1)
        fila_med.addWidget(nuevo)
        self.lote = QComboBox()
        self.lote.setEditable(True)
        self.lote.lineEdit().setPlaceholderText("Número de lote (o déjalo vacío)")
        self.vence = QDateEdit(a_qdate(date.today() + timedelta(days=365)))
        self.vence.setCalendarPopup(True)
        self.vence.setDisplayFormat("dd/MM/yyyy")
        self.sin_vence = QCheckBox("Sin fecha de vencimiento")
        self.sin_vence.toggled.connect(lambda v: self.vence.setEnabled(not v))
        fila_vence = QHBoxLayout()
        fila_vence.addWidget(self.vence)
        fila_vence.addWidget(self.sin_vence)
        self.fecha = QDateEdit(a_qdate(fecha))
        self.fecha.setCalendarPopup(True)
        self.fecha.setDisplayFormat("dddd dd/MM/yyyy")
        self.cantidad = QSpinBox()
        self.cantidad.setRange(1, 9_999_999)
        self.nota = QLineEdit()
        self.nota.setPlaceholderText("Opcional: proveedor, paciente, servicio…")
        self.info = QLabel()
        self.info.setObjectName("ayuda")

        f = QFormLayout(self)
        f.addRow("Tipo:", self.tipo)
        f.addRow("Medicamento:", fila_med)
        f.addRow("Lote:", self.lote)
        f.addRow("Vencimiento:", fila_vence)
        f.addRow("Fecha:", self.fecha)
        f.addRow("Cantidad:", self.cantidad)
        f.addRow("Nota:", self.nota)
        f.addRow(self.info)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Save).setText("Registrar")
        bb.button(QDialogButtonBox.Cancel).setText("Cancelar")
        bb.accepted.connect(self.guardar)
        bb.rejected.connect(self.reject)
        f.addRow(bb)

        self.medicamento.textChanged.connect(self.cargar_lotes)
        self.medicamento.completador.activated.connect(self.cargar_lotes)
        self.lote.currentIndexChanged.connect(self.lote_elegido)
        self.tipo.currentIndexChanged.connect(self.actualizar_info)
        self.cargar_lotes()

    def nuevo_medicamento(self):
        d = DialogoMedicamento(self.inv, nombre=self.medicamento.text(), parent=self)
        if d.exec():
            self.medicamento.recargar()
            self.medicamento.setText(self.inv.medicamento(d.medicamento_id)["nombre"])
            self.cargar_lotes()

    def cargar_lotes(self, *_):
        m = self.medicamento.medicamento()
        texto = self.lote.currentText()
        self.lote.blockSignals(True)
        self.lote.clear()
        if m:
            for l in self.inv.lotes(m["id"]):
                vence = texto_fecha(l["vencimiento"]) or "sin vencimiento"
                self.lote.addItem(f"{l['lote']}   (hay {l['existencia']}, vence {vence})",
                                  {"lote": l["lote"], "vence": l["vencimiento"], "hay": l["existencia"]})
        self.lote.setEditText(texto if texto and not m else (self.lote.itemText(0) if self.lote.count() else ""))
        self.lote.blockSignals(False)
        self.lote_elegido()

    def lote_elegido(self, *_):
        d = self.lote.currentData()
        if d and self.lote.currentText() == self.lote.itemText(self.lote.currentIndex()):
            if d["vence"]:
                self.sin_vence.setChecked(False)
                self.vence.setDate(QDate.fromString(d["vence"], "yyyy-MM-dd"))
            else:
                self.sin_vence.setChecked(True)
        self.actualizar_info()

    def actualizar_info(self, *_):
        m = self.medicamento.medicamento()
        if not self.medicamento.text().strip():
            self.info.setText("")
        elif not m:
            self.info.setText("⚠ No está en el catálogo. Usa «+ Nuevo» para agregarlo una sola vez.")
        else:
            self.info.setText(f"📝 {m['descripcion']}" if m["descripcion"] else "")

    def nombre_lote(self) -> str:
        texto = self.lote.currentText()
        d = self.lote.currentData()
        if d and texto == self.lote.itemText(self.lote.currentIndex()):
            return d["lote"]
        return texto.split("   (")[0].strip()

    def guardar(self):
        m = self.medicamento.medicamento()
        if not m:
            QMessageBox.warning(self, "Falta el medicamento",
                                "Elige un medicamento de la lista o agrégalo con «+ Nuevo».")
            return
        fecha = de_qdate(self.fecha.date())
        if not es_habil(fecha):
            QMessageBox.warning(self, "Fecha", "Solo se registran movimientos de lunes a viernes.")
            return
        tipo = self.tipo.currentData()
        lote = self.nombre_lote() or SIN_LOTE
        existe = any(l["lote"].upper() == lote.upper() for l in self.inv.lotes(m["id"]))
        if tipo == "SALIDA" and not existe:
            QMessageBox.warning(self, "Lote", f"El lote «{lote}» no existe para este medicamento. "
                                              "Para una salida elige un lote de la lista.")
            return
        vence = None if self.sin_vence.isChecked() else de_qdate(self.vence.date()).isoformat()
        try:
            lote_id = self.inv.obtener_lote(m["id"], lote, vence if tipo == "ENTRADA" or not existe else None)
            self.inv.sumar_dia(lote_id, fecha.isoformat(), tipo, self.cantidad.value(), self.nota.text())
        except ErrorInventario as e:
            QMessageBox.warning(self, "No se pudo registrar", str(e))
            return
        self.accept()


class DialogoLotes(QDialog):
    """Corregir número de lote o vencimiento de un medicamento."""

    def __init__(self, inv: Inventario, medicamento_id: int, parent=None):
        super().__init__(parent)
        self.inv, self.med = inv, medicamento_id
        self.setWindowTitle(f"Lotes de {inv.medicamento(medicamento_id)['nombre']}")
        self.resize(560, 360)
        self.t = tabla_basica(["Lote", "Vencimiento (dd/mm/aaaa)", "Existencia"])
        self.t.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        v = QVBoxLayout(self)
        ayuda = QLabel("Doble clic para corregir el lote o la fecha. Deja la fecha vacía si no vence.")
        ayuda.setObjectName("ayuda")
        v.addWidget(ayuda)
        v.addWidget(self.t)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Save).setText("Guardar cambios")
        bb.button(QDialogButtonBox.Cancel).setText("Cerrar")
        bb.accepted.connect(self.guardar)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        for l in inv.lotes(medicamento_id):
            r = self.t.rowCount()
            self.t.insertRow(r)
            self.t.setItem(r, 0, celda(l["lote"], editable=True, dato=l["id"]))
            self.t.setItem(r, 1, celda(texto_fecha(l["vencimiento"]), editable=True))
            self.t.setItem(r, 2, celda(l["existencia"]))

    def guardar(self):
        from datetime import datetime
        try:
            for r in range(self.t.rowCount()):
                texto = self.t.item(r, 1).text().strip()
                vence = None
                if texto:
                    try:
                        vence = datetime.strptime(texto, "%d/%m/%Y").date().isoformat()
                    except ValueError:
                        raise ErrorInventario(f"Fecha no válida en la fila {r + 1}: «{texto}». "
                                              "Usa dd/mm/aaaa.") from None
                self.inv.editar_lote(self.t.item(r, 0).data(Qt.UserRole), self.t.item(r, 0).text(), vence)
        except ErrorInventario as e:
            QMessageBox.warning(self, "No se pudo guardar", str(e))
            return
        self.accept()


def texto_fecha_hora(cuando: str) -> str:
    from datetime import datetime
    try:
        return datetime.strptime(cuando, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y a las %H:%M")
    except ValueError:
        return cuando


class DialogoCambios(QDialog):
    """Lista de cambios (cuándo, quién, antes, después) con botón para marcarlos revisados."""

    def __init__(self, inv: Inventario, cambios, titulo: str, parent=None, solo_ver=False):
        super().__init__(parent)
        self.inv = inv
        self.ids = [c["id"] for c in cambios if not c["revisado"]]
        self.setWindowTitle(titulo)
        self.resize(900, 420)
        t = tabla_basica(["Cuándo", "Quién", "Desde", "Medicamento", "Lote", "Qué cambió", "Antes", "Ahora",
                          "Revisado"])
        t.setRowCount(len(cambios))
        for r, c in enumerate(cambios):
            campo = c["campo"]
            valores = [texto_fecha_hora(c["cuando"]), c["usuario"] or "¿?", c["origen"].capitalize(),
                       c["nombre"] or "", c["nombre_lote"] or "", excel_vivo.NOMBRES_CAMPO.get(campo, campo),
                       excel_vivo.mostrar(campo, c["antes"]), excel_vivo.mostrar(campo, c["despues"]),
                       "Sí" if c["revisado"] else "No"]
            for i, v in enumerate(valores):
                t.setItem(r, i, celda(v, color=None if c["revisado"] else MORADO if i == 7 else None))
        v = QVBoxLayout(self)
        if cambios:
            v.addWidget(QLabel(f"<b>{len(cambios)}</b> cambio(s). El color morado se quita al marcarlos "
                               "como revisados; el historial se queda guardado."))
        else:
            v.addWidget(QLabel("No hay cambios."))
        v.addWidget(t)
        b = QHBoxLayout()
        b.addStretch()
        if self.ids and not solo_ver:
            b.addWidget(boton("✔ Marcar como revisado", self.revisar, principal=True))
        b.addWidget(boton("Cerrar", self.reject))
        v.addLayout(b)

    def revisar(self):
        self.inv.marcar_revisados(self.ids)
        self.accept()


# ── Pestañas ───────────────────────────────────────────────────────────────

class Pestana(QWidget):
    cambio = Signal()  # algo cambió en la base: las demás pestañas se recargan

    def __init__(self, inv: Inventario):
        super().__init__()
        self.inv = inv

    def recargar(self):
        pass

    def hojas(self) -> list[Hoja]:
        return []

    def nombre_archivo(self) -> str:
        return "inventario"

    def exportar(self, formato: str):
        hojas = self.hojas()
        if not hojas or not any(h.filas for h in hojas):
            QMessageBox.information(self, "Exportar", "No hay datos para exportar.")
            return
        sugerido = carpeta_documentos() / f"{self.nombre_archivo()}.{formato}"
        filtro = "Excel (*.xlsx)" if formato == "xlsx" else "PDF (*.pdf)"
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar como", str(sugerido), filtro)
        if not ruta:
            return
        ruta = Path(ruta)
        if ruta.suffix.lower() != f".{formato}":
            ruta = ruta.with_suffix(f".{formato}")
        try:
            (a_excel if formato == "xlsx" else a_pdf)(hojas, ruta)
        except PermissionError:
            QMessageBox.warning(self, "Archivo abierto",
                                f"No se pudo escribir «{ruta.name}». Ciérralo en Excel o en el lector de PDF "
                                "e inténtalo otra vez.")
            return
        r = QMessageBox.question(self, "Listo", f"Guardado en:\n{ruta}\n\n¿Abrirlo ahora?")
        if r == QMessageBox.Yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(ruta)))

    def botones_exportar(self, layout: QHBoxLayout):
        layout.addWidget(boton("📗 Excel", lambda: self.exportar("xlsx")))
        layout.addWidget(boton("📕 PDF", lambda: self.exportar("pdf")))


class PestanaSemana(Pestana):
    """Tabla de lunes a viernes: entradas y salidas de cada lote, día por día."""

    COL_DIA0 = 4

    def __init__(self, inv: Inventario):
        super().__init__(inv)
        self.lunes = lunes_de(date.today())
        self.filas: list[dict] = []
        self.cargando = False

        self.titulo = QLabel()
        self.titulo.setObjectName("titulo")
        arriba = QHBoxLayout()
        arriba.addWidget(boton("◀ Semana anterior", lambda: self.mover(-7)))
        arriba.addWidget(boton("Esta semana", self.hoy))
        arriba.addWidget(boton("Semana siguiente ▶", lambda: self.mover(7)))
        arriba.addSpacing(16)
        arriba.addWidget(self.titulo, 1)

        acciones = QHBoxLayout()
        acciones.addWidget(boton("➕ Registrar entrada / salida", self.nuevo_movimiento, principal=True))
        self.ver_ceros = QCheckBox("Mostrar lotes sin existencia")
        self.ver_ceros.toggled.connect(self.recargar)
        acciones.addWidget(self.ver_ceros)
        acciones.addStretch()
        self.botones_exportar(acciones)

        ayuda = QLabel("Escribe directo en las casillas verdes (entradas) y rojas (salidas) del día. "
                       "El saldo final se calcula solo. Sábado y domingo no se registran.")
        ayuda.setObjectName("ayuda")
        ayuda.setWordWrap(True)

        self.t = QTableWidget()
        self.t.verticalHeader().setVisible(False)
        self.t.setItemDelegate(DelegadoNumero(self.t))
        self.t.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed)
        self.t.setSelectionMode(QAbstractItemView.SingleSelection)
        self.t.itemChanged.connect(self.celda_cambiada)

        v = QVBoxLayout(self)
        v.addLayout(arriba)
        v.addLayout(acciones)
        v.addWidget(ayuda)
        v.addWidget(self.t)
        self.recargar()

    def mover(self, dias: int):
        self.lunes += timedelta(days=dias)
        self.recargar()

    def hoy(self):
        self.lunes = lunes_de(date.today())
        self.recargar()

    def texto_semana(self) -> str:
        vie = self.lunes + timedelta(days=4)
        if self.lunes.month == vie.month:
            return f"Semana del {self.lunes.day} al {vie.day} de {MESES[vie.month - 1]} de {vie.year}"
        return (f"Semana del {self.lunes.day} de {MESES[self.lunes.month - 1]} al "
                f"{vie.day} de {MESES[vie.month - 1]} de {vie.year}")

    def recargar(self):
        self.cargando = True
        self.titulo.setText(self.texto_semana())
        dias = dias_habiles(self.lunes)
        enc = ["Medicamento", "Lote", "Vence", "Inicial"]
        for d, corto in zip(dias, DIAS_CORTOS):
            enc += [f"{corto} {d:%d/%m}\nEntra", f"{corto} {d:%d/%m}\nSale"]
        enc.append("Final")
        self.t.clear()
        self.t.setColumnCount(len(enc))
        self.t.setHorizontalHeaderLabels(enc)
        self.filas = self.inv.filas_semana(self.lunes, solo_con_existencia=not self.ver_ceros.isChecked())
        self.t.setRowCount(len(self.filas))
        hoy = date.today().isoformat()
        for r, f in enumerate(self.filas):
            self.t.setItem(r, 0, celda(f["nombre"], dato=f["lote_id"]))
            self.t.item(r, 0).setToolTip(f["descripcion"] or "")
            self.t.setItem(r, 1, celda(f["lote"]))
            self.t.setItem(r, 2, celda(texto_fecha(f["vencimiento"]), color=color_vencimiento(f["vencimiento"])))
            self.t.setItem(r, 3, celda(f["inicial"], color=GRIS))
            for i, d in enumerate(f["dias"]):
                es_hoy = d["fecha"] == hoy
                ce = celda(d["entrada"] or "", editable=True, alinear=Qt.AlignRight | Qt.AlignVCenter,
                           color=HOY_COLOR if es_hoy else VERDE)
                cs = celda(d["salida"] or "", editable=True, alinear=Qt.AlignRight | Qt.AlignVCenter,
                           color=HOY_COLOR if es_hoy else ROJO)
                if d["ajuste"]:
                    for c in (ce, cs):
                        c.setToolTip(f"Ese día hubo un ajuste por cuadre de {d['ajuste']:+d}")
                self.t.setItem(r, self.COL_DIA0 + 2 * i, ce)
                self.t.setItem(r, self.COL_DIA0 + 2 * i + 1, cs)
            self.t.setItem(r, len(enc) - 1, celda(f["final"], negrita=True))
        self.t.resizeColumnsToContents()
        self.t.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.cargando = False

    def celda_cambiada(self, it: QTableWidgetItem):
        if self.cargando:
            return
        c = it.column()
        if not (self.COL_DIA0 <= c < self.COL_DIA0 + 10):
            return
        r = it.row()
        dia = (c - self.COL_DIA0) // 2
        tipo = "ENTRADA" if (c - self.COL_DIA0) % 2 == 0 else "SALIDA"
        fila = self.filas[r]
        fecha = fila["dias"][dia]["fecha"]
        texto = it.text().strip()
        cantidad = int(texto) if texto.isdigit() else 0
        try:
            self.inv.fijar_dia(fila["lote_id"], fecha, tipo, cantidad)
        except ErrorInventario as e:
            QMessageBox.warning(self, "No se pudo guardar", str(e))
        fila_sel, col_sel = r, c
        self.recargar()
        self.cambio.emit()
        if fila_sel < self.t.rowCount():
            self.t.setCurrentCell(fila_sel, col_sel)

    def nuevo_movimiento(self):
        hoy = date.today()
        fecha = hoy if self.lunes <= hoy <= self.lunes + timedelta(days=4) else self.lunes
        if not es_habil(fecha):
            fecha = self.lunes + timedelta(days=4)
        actual = self.t.currentRow()
        med = self.filas[actual]["nombre"] if 0 <= actual < len(self.filas) else ""
        if DialogoMovimiento(self.inv, fecha, medicamento=med, parent=self).exec():
            self.recargar()
            self.cambio.emit()

    def nombre_archivo(self):
        return f"Semana_{self.lunes.isoformat()}"

    def hojas(self):
        dias = dias_habiles(self.lunes)
        enc = ["Medicamento", "Descripción", "Lote", "Vence", "Inicial"]
        grupos = []
        for i, (d, nombre) in enumerate(zip(dias, DIAS)):
            enc += ["Entra", "Sale"]
            grupos.append((f"{nombre} {d:%d/%m}", 5 + 2 * i, 6 + 2 * i))
        enc.append("Final")
        filas, colores = [], {}
        totales = ["TOTAL", "", "", "", 0] + [0] * 10 + [0]
        for r, f in enumerate(self.filas):
            fila = [f["nombre"], f["descripcion"], f["lote"], texto_fecha(f["vencimiento"]), f["inicial"]]
            for i, d in enumerate(f["dias"]):
                fila += [d["entrada"] or "", d["salida"] or ""]
                totales[5 + 2 * i] += d["entrada"]
                totales[6 + 2 * i] += d["salida"]
            fila.append(f["final"])
            totales[4] += f["inicial"]
            totales[-1] += f["final"]
            filas.append(fila)
            cv = color_vencimiento(f["vencimiento"])
            if cv:
                colores[(r, 3)] = cv
        return [Hoja(f"Control semanal de medicamentos", enc, filas, subtitulo=self.texto_semana(),
                     nombre_hoja="Semana", colores=colores, grupos=grupos, totales=totales)]


class PestanaInventario(Pestana):
    """Tabla principal: medicamento, descripción, existencias físicas, kardex, lote, vencimiento."""

    def __init__(self, inv: Inventario):
        super().__init__(inv)
        self.buscar = QLineEdit()
        self.buscar.setPlaceholderText("🔍 Buscar por medicamento, descripción o lote…")
        self.buscar.textChanged.connect(self.recargar)
        self.inactivos = QCheckBox("Ver desactivados")
        self.inactivos.toggled.connect(self.recargar)
        self.vacios = QCheckBox("Ver lotes agotados")
        self.vacios.toggled.connect(self.recargar)
        arriba = QHBoxLayout()
        arriba.addWidget(self.buscar, 1)
        arriba.addWidget(self.vacios)
        arriba.addWidget(self.inactivos)

        acciones = QHBoxLayout()
        acciones.addWidget(boton("➕ Nuevo medicamento", self.nuevo, principal=True))
        acciones.addWidget(boton("✏ Editar", self.editar))
        acciones.addWidget(boton("🏷 Lotes y vencimientos", self.editar_lotes))
        self.b_activo = boton("🚫 Desactivar", self.alternar_activo)
        acciones.addWidget(self.b_activo)
        acciones.addStretch()
        self.botones_exportar(acciones)

        self.resumen = QLabel()
        self.resumen.setObjectName("ayuda")
        self.t = tabla_basica(["Medicamento", "Descripción", "Existencias físicas", "Kardex",
                               "Diferencia", "Lote", "Fecha de vencimiento", "Presentación", "Mínimo"])
        self.t.setWordWrap(True)
        self.t.doubleClicked.connect(self.editar)
        self.t.itemSelectionChanged.connect(self.seleccion)
        self.t.cellClicked.connect(self.clic_celda)
        self.aviso = QPushButton()
        self.aviso.setStyleSheet("QPushButton { background: #E4D7F5; border: 1px solid #A889D6; "
                                 "text-align: left; padding: 6px 12px; font-weight: bold; }")
        self.aviso.clicked.connect(lambda: self.mostrar_cambios(None))
        v = QVBoxLayout(self)
        v.addLayout(arriba)
        v.addWidget(self.aviso)
        v.addLayout(acciones)
        v.addWidget(self.t)
        v.addWidget(self.resumen)
        self.recargar()

    def recargar(self):
        filas = self.inv.inventario_por_lote(self.buscar.text(), self.inactivos.isChecked(),
                                             self.vacios.isChecked())
        self.filas = filas
        pendientes: dict[tuple, list] = {}
        todos = self.inv.cambios_pendientes()
        for c in todos:
            if c["lote_id"] and c["campo"] in self.CAMPOS_LOTE:
                pendientes.setdefault(("lote", c["lote_id"], c["campo"]), []).append(c)
            elif c["medicamento_id"]:
                campo = "nombre" if c["campo"] == "nuevo" else c["campo"]
                pendientes.setdefault(("med", c["medicamento_id"], campo), []).append(c)
        self.pendientes = pendientes
        self.aviso.setVisible(bool(todos))
        quienes = sorted({c["usuario"] or "¿?" for c in todos})
        self.aviso.setText(f"🟪 {len(todos)} cambio(s) hechos desde el Excel sin revisar "
                           f"(por {', '.join(quienes)}). Clic en una casilla morada para ver detalle, "
                           "o aquí para verlos todos.")
        totales: dict[int, int] = {}
        for f in filas:
            totales[f["med_id"]] = totales.get(f["med_id"], 0) + (f["kardex"] or 0)
        self.t.setRowCount(len(filas))
        for r, f in enumerate(filas):
            gris = None if f["activo"] else GRIS
            self.t.setItem(r, 0, celda(f["nombre"] + ("" if f["activo"] else " (desactivado)"),
                                       dato=f["med_id"], color=gris, negrita=True))
            self.t.setItem(r, 1, celda(f["descripcion"], color=gris))
            fisico = f["fisico"]
            ft = celda("" if fisico is None else fisico, alinear=Qt.AlignRight | Qt.AlignVCenter)
            if f["fecha_conteo"]:
                ft.setToolTip(f"Contado el {texto_fecha(f['fecha_conteo'])}")
            self.t.setItem(r, 2, ft)
            bajo = f["minimo"] and totales[f["med_id"]] < f["minimo"]
            kt = celda(f["kardex"] if f["lote"] else 0, color=ROJO if bajo else None)
            if bajo:
                kt.setToolTip(f"Por debajo del mínimo ({f['minimo']}). Total de todos los lotes: "
                              f"{totales[f['med_id']]}")
            self.t.setItem(r, 3, kt)
            if fisico is not None and f["lote"]:
                dif = fisico - f["kardex"]
                self.t.setItem(r, 4, celda(dif, color=None if dif == 0 else NARANJA))
            else:
                self.t.setItem(r, 4, celda(""))
            self.t.setItem(r, 5, celda(f["lote"] or "—"))
            self.t.setItem(r, 6, celda(texto_fecha(f["vencimiento"]), color=color_vencimiento(f["vencimiento"])))
            self.t.setItem(r, 7, celda(f["presentacion"]))
            self.t.setItem(r, 8, celda(f["minimo"] or ""))
            self.pintar_cambios(r, f, pendientes)
        self.t.resizeColumnsToContents()
        self.t.horizontalHeader().setStretchLastSection(False)
        self.t.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.t.resizeRowsToContents()
        meds = len({f["med_id"] for f in filas})
        self.resumen.setText(f"{meds} medicamentos · {len([f for f in filas if f['lote']])} lotes · "
                             "Existencias físicas = último cuadre. Kardex = lo que dicen las entradas y salidas. "
                             "Rojo: vencido o bajo el mínimo · naranja: vence en 30 días · amarillo: en 90.")
        self.seleccion()

    # Columna de la tabla para cada campo que puede cambiar desde el Excel.
    COLUMNA = {"nombre": 0, "descripcion": 1, "fisico": 2, "kardex": 3, "lote": 5, "vencimiento": 6,
               "presentacion": 7, "minimo": 8}
    CAMPOS_LOTE = ("lote", "vencimiento", "fisico", "kardex")

    def cambios_de_celda(self, r: int, c: int) -> list:
        if not (0 <= r < len(self.filas)):
            return []
        f = self.filas[r]
        campo = next((k for k, v in self.COLUMNA.items() if v == c), None)
        if campo is None:
            return []
        if campo in self.CAMPOS_LOTE:
            return self.pendientes.get(("lote", f["lote_id"], campo), []) if f["lote_id"] else []
        return self.pendientes.get(("med", f["med_id"], campo), [])

    def pintar_cambios(self, r: int, f: dict, pendientes: dict):
        for campo, c in self.COLUMNA.items():
            lista = self.cambios_de_celda(r, c)
            if lista:
                it = self.t.item(r, c)
                it.setBackground(QColor("#" + MORADO))
                ultimo = lista[0]
                it.setToolTip(f"Cambiado desde Excel por {ultimo['usuario'] or '¿?'} el "
                              f"{texto_fecha_hora(ultimo['cuando'])}\n"
                              f"Antes: {excel_vivo.mostrar(ultimo['campo'], ultimo['antes'])} · "
                              f"Ahora: {excel_vivo.mostrar(ultimo['campo'], ultimo['despues'])}\n"
                              "Clic para ver el detalle.")

    def clic_celda(self, r: int, c: int):
        lista = self.cambios_de_celda(r, c)
        if lista:
            self.mostrar_cambios(lista)

    def mostrar_cambios(self, lista):
        titulo = "Cambios hechos desde el Excel"
        if lista is None:
            lista = self.inv.cambios_pendientes()
        if DialogoCambios(self.inv, lista, titulo, self).exec():
            self.recargar()
            self.cambio.emit()

    def med_actual(self) -> int | None:
        r = self.t.currentRow()
        return self.t.item(r, 0).data(Qt.UserRole) if r >= 0 and self.t.item(r, 0) else None

    def seleccion(self):
        mid = self.med_actual()
        if mid:
            activo = self.inv.medicamento(mid)["activo"]
            self.b_activo.setText("🚫 Desactivar" if activo else "✅ Activar")

    def nuevo(self):
        d = DialogoMedicamento(self.inv, nombre=self.buscar.text(), parent=self)
        if d.exec():
            self.buscar.clear()
            self.recargar()
            self.cambio.emit()
            r = QMessageBox.question(self, "Medicamento agregado",
                                     "¿Registrar ahora la primera entrada (lote, vencimiento y cantidad)?")
            if r == QMessageBox.Yes:
                hoy = date.today()
                fecha = hoy if es_habil(hoy) else hoy - timedelta(days=hoy.weekday() - 4)
                nombre = self.inv.medicamento(d.medicamento_id)["nombre"]
                if DialogoMovimiento(self.inv, fecha, medicamento=nombre, parent=self).exec():
                    self.recargar()
                    self.cambio.emit()

    def editar(self, *_):
        mid = self.med_actual()
        if not mid:
            return
        if DialogoMedicamento(self.inv, mid, parent=self).exec():
            self.recargar()
            self.cambio.emit()

    def editar_lotes(self):
        mid = self.med_actual()
        if not mid:
            QMessageBox.information(self, "Lotes", "Primero selecciona un medicamento en la tabla.")
            return
        if DialogoLotes(self.inv, mid, self).exec():
            self.recargar()
            self.cambio.emit()

    def alternar_activo(self):
        mid = self.med_actual()
        if not mid:
            return
        m = self.inv.medicamento(mid)
        if m["activo"]:
            r = QMessageBox.question(
                self, "Desactivar",
                f"¿Desactivar «{m['nombre']}»?\n\nNo se borra nada: deja de salir en las listas y en la "
                "semana. Puedes activarlo otra vez con «Ver desactivados».")
            if r != QMessageBox.Yes:
                return
        self.inv.activar_medicamento(mid, not m["activo"])
        self.recargar()
        self.cambio.emit()

    def nombre_archivo(self):
        return f"Inventario_{date.today().isoformat()}"

    def hojas(self):
        return [hoja_de_tabla(self.t, "Inventario de medicamentos",
                              f"Al {date.today():%d/%m/%Y}", nombre_hoja="Inventario")]


class PestanaKardex(Pestana):
    def __init__(self, inv: Inventario):
        super().__init__(inv)
        self.medicamento = BuscadorMedicamento(inv)
        self.medicamento.textChanged.connect(self.medicamento_cambiado)
        self.medicamento.completador.activated.connect(self.medicamento_cambiado)
        self.lote = QComboBox()
        self.lote.currentIndexChanged.connect(self.recargar)
        self.usar_fechas = QCheckBox("Entre fechas:")
        self.desde = QDateEdit(a_qdate(date.today().replace(day=1)))
        self.hasta = QDateEdit(a_qdate(date.today()))
        for d in (self.desde, self.hasta):
            d.setCalendarPopup(True)
            d.setDisplayFormat("dd/MM/yyyy")
            d.setEnabled(False)
            d.dateChanged.connect(self.recargar)
        self.usar_fechas.toggled.connect(lambda v: (self.desde.setEnabled(v), self.hasta.setEnabled(v),
                                                    self.recargar()))
        arriba = QHBoxLayout()
        arriba.addWidget(QLabel("Medicamento:"))
        arriba.addWidget(self.medicamento, 2)
        arriba.addWidget(QLabel("Lote:"))
        arriba.addWidget(self.lote, 1)
        arriba.addWidget(self.usar_fechas)
        arriba.addWidget(self.desde)
        arriba.addWidget(QLabel("a"))
        arriba.addWidget(self.hasta)
        acciones = QHBoxLayout()
        self.info = QLabel("Elige un medicamento para ver su kardex (todas sus entradas y salidas con el saldo).")
        self.info.setObjectName("ayuda")
        self.info.setWordWrap(True)
        acciones.addWidget(self.info, 1)
        self.botones_exportar(acciones)
        acciones.addWidget(boton("📗 Kardex de TODOS (Excel)", self.exportar_todo))
        self.t = tabla_basica(["Fecha", "Día", "Lote", "Concepto", "Entrada", "Salida", "Saldo", "Nota"])
        v = QVBoxLayout(self)
        v.addLayout(arriba)
        v.addLayout(acciones)
        v.addWidget(self.t)

    def medicamento_cambiado(self, *_):
        m = self.medicamento.medicamento()
        self.lote.blockSignals(True)
        self.lote.clear()
        self.lote.addItem("Todos los lotes", None)
        if m:
            for l in self.inv.lotes(m["id"]):
                self.lote.addItem(l["lote"], l["id"])
        self.lote.blockSignals(False)
        self.recargar()

    def recargar(self):
        m = self.medicamento.medicamento()
        self.t.setRowCount(0)
        if not m:
            return
        desde = de_qdate(self.desde.date()).isoformat() if self.usar_fechas.isChecked() else None
        hasta = de_qdate(self.hasta.date()).isoformat() if self.usar_fechas.isChecked() else None
        filas = self.inv.kardex(m["id"], self.lote.currentData(), desde, hasta)
        self.t.setRowCount(len(filas))
        for r, k in enumerate(filas):
            d = date.fromisoformat(k.fecha)
            self.t.setItem(r, 0, celda(texto_fecha(k.fecha)))
            self.t.setItem(r, 1, celda(DIAS[d.weekday()] if d.weekday() < 5 else ""))
            self.t.setItem(r, 2, celda(k.lote))
            color = (MORADO if "Excel" in k.concepto else NARANJA) if k.concepto.startswith("Ajuste") else None
            self.t.setItem(r, 3, celda(k.concepto, color=color, negrita=k.concepto == "Saldo anterior"))
            self.t.setItem(r, 4, celda(k.entrada or "", alinear=Qt.AlignRight | Qt.AlignVCenter,
                                       color=VERDE if k.entrada else None))
            self.t.setItem(r, 5, celda(k.salida or "", alinear=Qt.AlignRight | Qt.AlignVCenter,
                                       color=ROJO if k.salida else None))
            self.t.setItem(r, 6, celda(k.saldo, negrita=True))
            self.t.setItem(r, 7, celda(k.nota))
        total = sum(l["existencia"] for l in self.inv.lotes(m["id"]))
        self.info.setText(f"<b>{m['nombre']}</b> — {m['descripcion'] or 'sin descripción'} · "
                          f"Existencia según kardex: <b>{total}</b>")
        self.t.scrollToBottom()

    def nombre_archivo(self):
        m = self.medicamento.medicamento()
        return "Kardex_" + ("".join(ch if ch.isalnum() else "_" for ch in m["nombre"]) if m else "")

    def hojas(self):
        m = self.medicamento.medicamento()
        if not m:
            return []
        sub = f"{m['descripcion']} · Lote: {self.lote.currentText()}"
        if self.usar_fechas.isChecked():
            sub += f" · Del {self.desde.date().toString('dd/MM/yyyy')} al {self.hasta.date().toString('dd/MM/yyyy')}"
        return [hoja_de_tabla(self.t, f"Kardex — {m['nombre']}", sub, nombre_hoja="Kardex")]

    def exportar_todo(self):
        filas = self.inv.todos_los_movimientos()
        if not filas:
            QMessageBox.information(self, "Exportar", "Todavía no hay movimientos.")
            return
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar como", str(carpeta_documentos() / f"Kardex_completo_{date.today()}.xlsx"),
            "Excel (*.xlsx)")
        if not ruta:
            return
        h = Hoja("Kardex de todos los medicamentos", ["Medicamento", "Fecha", "Lote", "Concepto",
                                                      "Entrada", "Salida", "Saldo", "Nota"],
                 [[f["nombre"], texto_fecha(f["fecha"]), f["lote"], f["concepto"], f["entrada"] or "",
                   f["salida"] or "", f["saldo"], f["nota"]] for f in filas], nombre_hoja="Kardex")
        try:
            a_excel([h], Path(ruta))
        except PermissionError:
            QMessageBox.warning(self, "Archivo abierto", "Cierra el archivo en Excel e inténtalo otra vez.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(ruta))


class PestanaCuadre(Pestana):
    """Conteo físico semanal o mensual contra lo que dice el kardex."""

    def __init__(self, inv: Inventario):
        super().__init__(inv)
        self.filas: list[dict] = []
        self.cargando = False
        self.tipo = QComboBox()
        self.tipo.addItem("Cuadre semanal", "SEMANAL")
        self.tipo.addItem("Cuadre mensual (fin de mes)", "MENSUAL")
        self.tipo.currentIndexChanged.connect(self.fecha_sugerida)
        self.fecha = QDateEdit()
        self.fecha.setCalendarPopup(True)
        self.fecha.setDisplayFormat("dddd dd/MM/yyyy")
        self.fecha.dateChanged.connect(self.recargar)
        arriba = QHBoxLayout()
        arriba.addWidget(self.tipo)
        arriba.addWidget(QLabel("Fecha del conteo:"))
        arriba.addWidget(self.fecha)
        arriba.addStretch()
        arriba.addWidget(boton("🖨 Hoja para contar (PDF)", self.hoja_conteo))
        self.botones_exportar(arriba)

        ayuda = QLabel("1) Imprime la hoja para contar. 2) Cuenta lo que hay en el estante. "
                       "3) Escribe el número en «Físico». 4) Guarda: el programa muestra las diferencias "
                       "y, si quieres, corrige el kardex.")
        ayuda.setObjectName("ayuda")
        ayuda.setWordWrap(True)
        self.t = tabla_basica(["Medicamento", "Lote", "Vence", "Sistema (kardex)", "Físico (contado)",
                               "Diferencia"])
        self.t.setItemDelegateForColumn(4, DelegadoNumero(self.t))
        self.t.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed)
        self.t.itemChanged.connect(self.celda_cambiada)
        self.resumen = QLabel()
        guardar = QHBoxLayout()
        guardar.addWidget(self.resumen, 1)
        guardar.addWidget(boton("💾 Guardar cuadre", self.guardar, principal=True))

        historial_titulo = QLabel("Cuadres anteriores (doble clic para ver)")
        historial_titulo.setObjectName("titulo")
        self.historial = tabla_basica(["Fecha", "Tipo", "Líneas", "Con diferencia", "¿Se ajustó el kardex?",
                                       "Nota"])
        self.historial.setMaximumHeight(180)
        self.historial.doubleClicked.connect(self.ver_cuadre)

        v = QVBoxLayout(self)
        v.addLayout(arriba)
        v.addWidget(ayuda)
        v.addWidget(self.t, 3)
        v.addLayout(guardar)
        v.addWidget(historial_titulo)
        v.addWidget(self.historial, 1)
        self.fecha_sugerida()

    def fecha_sugerida(self):
        hoy = date.today()
        if self.tipo.currentData() == "MENSUAL":
            f = ultimo_habil_del_mes(hoy)
            if f > hoy:
                f = hoy if es_habil(hoy) else hoy - timedelta(days=hoy.weekday() - 4)
        else:
            f = hoy if es_habil(hoy) else hoy - timedelta(days=hoy.weekday() - 4)
        self.fecha.setDate(a_qdate(f))
        self.recargar()

    def recargar(self):
        self.cargando = True
        fecha = de_qdate(self.fecha.date()).isoformat()
        anteriores = {f["lote_id"]: self.t.item(r, 4).text() for r, f in enumerate(self.filas)
                      if self.t.item(r, 4)} if self.filas else {}
        self.filas = self.inv.filas_para_cuadre(fecha)
        self.t.setRowCount(len(self.filas))
        for r, f in enumerate(self.filas):
            self.t.setItem(r, 0, celda(f["nombre"]))
            self.t.setItem(r, 1, celda(f["lote"]))
            self.t.setItem(r, 2, celda(texto_fecha(f["vencimiento"]), color=color_vencimiento(f["vencimiento"])))
            self.t.setItem(r, 3, celda(f["sistema"]))
            self.t.setItem(r, 4, celda(anteriores.get(f["lote_id"], ""), editable=True,
                                       alinear=Qt.AlignRight | Qt.AlignVCenter, color=AMARILLO))
            self.t.setItem(r, 5, celda(""))
        self.cargando = False
        for r in range(len(self.filas)):
            self.diferencia(r)
        self.t.resizeColumnsToContents()
        self.t.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.cargar_historial()

    def diferencia(self, r: int):
        texto = self.t.item(r, 4).text().strip()
        if not texto.isdigit():
            self.t.item(r, 5).setText("")
            self.t.item(r, 5).setBackground(QColor(Qt.transparent))
        else:
            dif = int(texto) - self.filas[r]["sistema"]
            it = self.t.item(r, 5)
            it.setText(f"{dif:+d}" if dif else "0 ✔")
            it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            it.setBackground(QColor("#" + (VERDE if dif == 0 else ROJO)))
        contados = [r for r in range(self.t.rowCount()) if self.t.item(r, 4).text().strip().isdigit()]
        malos = [r for r in contados if self.t.item(r, 5).text() != "0 ✔"]
        self.resumen.setText(f"Contados {len(contados)} de {len(self.filas)} lotes · "
                             f"{len(malos)} con diferencia")

    def celda_cambiada(self, it: QTableWidgetItem):
        if not self.cargando and it.column() == 4:
            self.cargando = True
            self.diferencia(it.row())
            self.cargando = False

    def guardar(self):
        conteos = {}
        for r, f in enumerate(self.filas):
            texto = self.t.item(r, 4).text().strip()
            if texto.isdigit():
                conteos[f["lote_id"]] = int(texto)
        if not conteos:
            QMessageBox.information(self, "Cuadre", "Escribe al menos un conteo en la columna «Físico».")
            return
        faltan = len(self.filas) - len(conteos)
        difs = sum(1 for r, f in enumerate(self.filas)
                   if f["lote_id"] in conteos and conteos[f["lote_id"]] != f["sistema"])
        caja = QMessageBox(self)
        caja.setWindowTitle("Guardar cuadre")
        caja.setIcon(QMessageBox.Question)
        texto = f"Se contaron {len(conteos)} lotes"
        if faltan:
            texto += f" (faltan {faltan} sin contar; esos no se tocan)"
        texto += f".\n{difs} tienen diferencia con el kardex."
        caja.setText(texto)
        caja.setInformativeText("¿Corregir el kardex para que quede igual a lo contado?\n"
                                "Se anota un «Ajuste por cuadre» con la diferencia.")
        si = caja.addButton("Guardar y ajustar kardex", QMessageBox.YesRole)
        no = caja.addButton("Solo guardar el conteo", QMessageBox.NoRole)
        caja.addButton("Cancelar", QMessageBox.RejectRole)
        caja.exec()
        if caja.clickedButton() not in (si, no):
            return
        try:
            self.inv.guardar_cuadre(de_qdate(self.fecha.date()).isoformat(), self.tipo.currentData(),
                                    conteos, ajustar=caja.clickedButton() is si)
        except ErrorInventario as e:
            QMessageBox.warning(self, "No se pudo guardar", str(e))
            return
        self.filas = []
        self.recargar()
        self.cambio.emit()
        QMessageBox.information(self, "Cuadre", "Cuadre guardado. Aparece abajo en «Cuadres anteriores».")

    def cargar_historial(self):
        filas = self.inv.cuadres()
        self.historial.setRowCount(len(filas))
        for r, c in enumerate(filas):
            self.historial.setItem(r, 0, celda(texto_fecha(c["fecha"]), dato=c["id"]))
            self.historial.setItem(r, 1, celda(c["tipo"].capitalize()))
            self.historial.setItem(r, 2, celda(c["lineas"]))
            self.historial.setItem(r, 3, celda(c["con_diferencia"] or 0,
                                               color=NARANJA if c["con_diferencia"] else None))
            self.historial.setItem(r, 4, celda("Sí" if c["ajustado"] else "No"))
            self.historial.setItem(r, 5, celda(c["nota"]))

    def ver_cuadre(self, *_):
        r = self.historial.currentRow()
        if r < 0:
            return
        cid = self.historial.item(r, 0).data(Qt.UserRole)
        fecha = self.historial.item(r, 0).text()
        tipo = self.historial.item(r, 1).text()
        d = QDialog(self)
        d.setWindowTitle(f"Cuadre {tipo.lower()} del {fecha}")
        d.resize(760, 480)
        t = tabla_basica(["Medicamento", "Lote", "Vence", "Sistema", "Físico", "Diferencia"])
        det = self.inv.detalle_cuadre(cid)
        t.setRowCount(len(det))
        for i, x in enumerate(det):
            dif = x["fisico"] - x["sistema"]
            for c, v in enumerate([x["nombre"], x["lote"], texto_fecha(x["vencimiento"]), x["sistema"],
                                   x["fisico"], dif]):
                t.setItem(i, c, celda(v, color=(ROJO if dif else VERDE) if c == 5 else None))
        v = QVBoxLayout(d)
        v.addWidget(t)
        b = QHBoxLayout()
        b.addStretch()
        titulo = f"Cuadre {tipo.lower()} — {fecha}"

        def exportar(fmt):
            ruta, _ = QFileDialog.getSaveFileName(
                d, "Guardar como", str(carpeta_documentos() / f"Cuadre_{tipo}_{fecha.replace('/', '-')}.{fmt}"),
                "Excel (*.xlsx)" if fmt == "xlsx" else "PDF (*.pdf)")
            if ruta:
                (a_excel if fmt == "xlsx" else a_pdf)([hoja_de_tabla(t, titulo, nombre_hoja="Cuadre")], Path(ruta))
                QDesktopServices.openUrl(QUrl.fromLocalFile(ruta))

        b.addWidget(boton("📗 Excel", lambda: exportar("xlsx")))
        b.addWidget(boton("📕 PDF", lambda: exportar("pdf")))
        b.addWidget(boton("Cerrar", d.accept))
        v.addLayout(b)
        d.exec()

    def hoja_conteo(self):
        if not self.filas:
            QMessageBox.information(self, "Hoja para contar", "No hay lotes con existencia a esa fecha.")
            return
        fecha = self.fecha.date().toString("dd/MM/yyyy")
        h = Hoja(f"Hoja de conteo físico — {self.tipo.currentText()}",
                 ["Medicamento", "Lote", "Vence", "Contado", "Observaciones"],
                 [[f["nombre"], f["lote"], texto_fecha(f["vencimiento"]), "", " " * 30] for f in self.filas],
                 subtitulo=f"Fecha: {fecha}      Contó: ____________________      Revisó: ____________________")
        ruta = carpeta_documentos() / f"Hoja_conteo_{de_qdate(self.fecha.date()).isoformat()}.pdf"
        try:
            a_pdf([h], ruta)
        except PermissionError:
            QMessageBox.warning(self, "Archivo abierto", "Cierra el PDF anterior e inténtalo otra vez.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(ruta)))

    def nombre_archivo(self):
        return f"Cuadre_{self.tipo.currentData().lower()}_{de_qdate(self.fecha.date()).isoformat()}"

    def hojas(self):
        return [hoja_de_tabla(self.t, self.tipo.currentText(),
                              f"Fecha: {self.fecha.date().toString('dd/MM/yyyy')}", nombre_hoja="Cuadre")]


class PestanaVencimientos(Pestana):
    def __init__(self, inv: Inventario):
        super().__init__(inv)
        self.dias = QSpinBox()
        self.dias.setRange(0, 3650)
        self.dias.setValue(90)
        self.dias.setSuffix(" días")
        self.dias.valueChanged.connect(self.recargar)
        arriba = QHBoxLayout()
        arriba.addWidget(QLabel("Mostrar lotes con existencia que vencen en los próximos"))
        arriba.addWidget(self.dias)
        arriba.addStretch()
        self.botones_exportar(arriba)
        self.t = tabla_basica(["Medicamento", "Lote", "Vence", "Faltan", "Existencia"])
        v = QVBoxLayout(self)
        v.addLayout(arriba)
        v.addWidget(self.t)
        self.recargar()

    def recargar(self):
        filas = self.inv.vencimientos(self.dias.value())
        self.t.setRowCount(len(filas))
        for r, f in enumerate(filas):
            color = color_vencimiento(f["vencimiento"])
            faltan = "VENCIDO" if f["dias"] < 0 else ("hoy" if f["dias"] == 0 else f"{f['dias']} días")
            for c, v in enumerate([f["nombre"], f["lote"], texto_fecha(f["vencimiento"]), faltan, f["existencia"]]):
                self.t.setItem(r, c, celda(v, color=color))
        self.t.resizeColumnsToContents()

    def nombre_archivo(self):
        return f"Vencimientos_{date.today().isoformat()}"

    def hojas(self):
        return [hoja_de_tabla(self.t, "Medicamentos por vencer",
                              f"Próximos {self.dias.value()} días desde {date.today():%d/%m/%Y}",
                              nombre_hoja="Vencimientos")]


# ── Actualizaciones en segundo plano ─────────────────────────────────────────

class BuscadorActualizacion(QObject):
    listo = Signal(object, str)  # (Novedad | None, error)

    def correr(self):
        try:
            self.listo.emit(actualizador.buscar(), "")
        except Exception as e:  # sin internet, GitHub caído, etc.
            self.listo.emit(None, str(e))


# ── Ventana principal ───────────────────────────────────────────────────────

class VentanaPrincipal(QMainWindow):
    def __init__(self, inv: Inventario):
        super().__init__()
        self.inv = inv
        self.setWindowTitle(f"{NOMBRE_APP} — v{VERSION}")
        self.resize(1300, 760)

        self.tabs = QTabWidget()
        self.semana = PestanaSemana(inv)
        self.inventario = PestanaInventario(inv)
        self.kardex = PestanaKardex(inv)
        self.cuadre = PestanaCuadre(inv)
        self.vencimientos = PestanaVencimientos(inv)
        for p, nombre in [(self.semana, "📅 Semana (Lun–Vie)"), (self.inventario, "💊 Inventario"),
                          (self.kardex, "📒 Kardex"), (self.cuadre, "⚖ Cuadre / conteo físico"),
                          (self.vencimientos, "⏰ Vencimientos")]:
            self.tabs.addTab(p, nombre)
            p.cambio.connect(self.algo_cambio)
        self.tabs.currentChanged.connect(lambda i: self.tabs.widget(i).recargar())
        self.setCentralWidget(self.tabs)
        self.menus()
        self.statusBar().showMessage(f"Datos en: {inv.ruta}")

        # Cambio de semana automático si el programa queda abierto de un día para otro.
        self.dia_actual = date.today()
        self.reloj = QTimer(self)
        self.reloj.timeout.connect(self.revisar_dia)
        self.reloj.start(60_000)

        # Inventario_actual.xlsx: lo que la gente cambie ahí entra al programa.
        self.ruta_excel = excel_vivo.ruta_por_defecto(carpeta_documentos())
        self.excel_pendiente = True  # hay que volver a escribir el Excel cuando no esté abierto
        self.ultimo_mtime = None
        self.vigia_excel = QTimer(self)
        self.vigia_excel.timeout.connect(self.revisar_excel)
        self.vigia_excel.start(3000)
        QTimer.singleShot(800, self.revisar_excel)

        if actualizador.configurado() and actualizador.es_exe():
            QTimer.singleShot(3000, lambda: self.buscar_actualizacion(silencioso=True))

    def menus(self):
        archivo = self.menuBar().addMenu("&Archivo")
        a = QAction("📗 Exportar TODO a Excel", self)
        a.setShortcut(QKeySequence("Ctrl+E"))
        a.triggered.connect(self.exportar_todo)
        archivo.addAction(a)
        archivo.addSeparator()
        a = QAction("💾 Hacer respaldo ahora", self)
        a.triggered.connect(self.respaldar)
        archivo.addAction(a)
        a = QAction("♻ Restaurar un respaldo…", self)
        a.triggered.connect(self.restaurar)
        archivo.addAction(a)
        a = QAction("📂 Abrir carpeta de datos", self)
        a.triggered.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.inv.ruta.parent))))
        archivo.addAction(a)
        a = QAction("📊 Abrir Inventario_actual.xlsx", self)
        a.triggered.connect(self.abrir_excel_vivo)
        archivo.addAction(a)
        a = QAction("🕘 Historial de cambios desde Excel", self)
        a.triggered.connect(lambda: DialogoCambios(self.inv, self.inv.historial_cambios(),
                                                   "Historial de cambios", self, solo_ver=True).exec())
        archivo.addAction(a)
        a = QAction("📂 Abrir carpeta de exportaciones", self)
        a.triggered.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(carpeta_documentos()))))
        archivo.addAction(a)
        archivo.addSeparator()
        a = QAction("Salir", self)
        a.triggered.connect(self.close)
        archivo.addAction(a)

        ayuda = self.menuBar().addMenu("A&yuda")
        a = QAction("🔄 Buscar actualizaciones", self)
        a.triggered.connect(lambda: self.buscar_actualizacion(silencioso=False))
        ayuda.addAction(a)
        a = QAction("❓ ¿Qué es cada cosa?", self)
        a.triggered.connect(self.explicar)
        ayuda.addAction(a)
        a = QAction("Acerca de", self)
        a.triggered.connect(lambda: QMessageBox.about(
            self, NOMBRE_APP, f"<b>{NOMBRE_APP}</b><br>Versión {VERSION}<br><br>Datos: {self.inv.ruta}"))
        ayuda.addAction(a)

    def algo_cambio(self):
        actual = self.tabs.currentWidget()
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if w is not actual:
                w.recargar()
        if hasattr(self.kardex, "medicamento"):
            self.kardex.medicamento.recargar()
        self.excel_pendiente = True

    def abrir_excel_vivo(self):
        if not self.ruta_excel.exists():
            excel_vivo.exportar(self.inv, self.ruta_excel)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.ruta_excel)))

    def revisar_excel(self):
        """Cada 3 s: si guardaron el Excel, tomar los cambios; si el programa cambió algo, reescribirlo."""
        if excel_vivo.hay_cambios(self.inv, self.ruta_excel):
            try:
                mtime = self.ruta_excel.stat().st_mtime_ns
            except OSError:
                return
            # Esperar a que el archivo deje de cambiar (Excel guarda en varios pasos).
            if mtime != self.ultimo_mtime:
                self.ultimo_mtime = mtime
                return
            res = excel_vivo.importar(self.inv, self.ruta_excel)
            if res.aplicados or res.choques:
                for i in range(self.tabs.count()):
                    self.tabs.widget(i).recargar()
                self.kardex.medicamento.recargar()
            self.excel_pendiente = True
            if res.hay_algo:
                self.avisar_importacion(res)
            return
        if self.excel_pendiente and not excel_vivo.abierto_en_otro_programa(self.ruta_excel):
            try:
                if excel_vivo.exportar(self.inv, self.ruta_excel):
                    self.excel_pendiente = False
                    self.ultimo_mtime = self.ruta_excel.stat().st_mtime_ns
            except Exception as e:
                self.statusBar().showMessage(f"No se pudo actualizar Inventario_actual.xlsx: {e}", 10000)

    def avisar_importacion(self, res):
        partes = []
        if res.aplicados:
            partes.append(f"<b>Se aplicaron {len(res.aplicados)} cambio(s) hechos por {res.usuario}:</b><br>"
                          + "<br>".join(f"• {x}" for x in res.aplicados[:15])
                          + ("<br>…" if len(res.aplicados) > 15 else ""))
        if res.choques:
            partes.append("<b>⚠ No se aplicaron (el dato también cambió en el programa):</b><br>"
                          + "<br>".join(f"• {x}" for x in res.choques[:10]))
        if res.errores:
            partes.append("<b>✗ Con error (se ignoraron):</b><br>" + "<br>".join(f"• {x}" for x in res.errores[:10]))
        caja = QMessageBox(QMessageBox.Information, "Cambios desde el Excel", "<br><br>".join(partes),
                           QMessageBox.Ok, self)
        caja.setWindowModality(Qt.NonModal)
        caja.setAttribute(Qt.WA_DeleteOnClose)
        caja.show()
        self.statusBar().showMessage(f"Excel: {len(res.aplicados)} cambio(s) aplicados", 15000)

    def revisar_dia(self):
        if date.today() != self.dia_actual:
            semana_vieja = lunes_de(self.dia_actual)
            self.dia_actual = date.today()
            if self.semana.lunes == semana_vieja:
                self.semana.hoy()
            self.vencimientos.recargar()
            self.inv.respaldar()

    def explicar(self):
        QMessageBox.information(self, "¿Qué es cada cosa?", (
            "<b>Semana (Lun–Vie):</b> cada día se anota lo que entró y lo que salió. "
            "Al empezar otra semana, la tabla cambia sola de fechas.<br><br>"
            "<b>Kardex:</b> la historia de cada medicamento: cada entrada y salida con el saldo que "
            "quedó. Es lo que <i>debería</i> haber según lo registrado.<br><br>"
            "<b>Existencias físicas:</b> lo que se contó a mano en el estante en el último cuadre.<br><br>"
            "<b>Cuadre:</b> comparar lo físico con el kardex. Si no coinciden, hay un error o una pérdida. "
            "Se puede hacer cada semana y a fin de mes.<br><br>"
            "<b>Lote y vencimiento:</b> cada caja trae su número de lote y fecha. Un mismo medicamento "
            "puede tener varios lotes; se recomienda sacar primero el que vence antes."))

    def exportar_todo(self):
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar como", str(carpeta_documentos() / f"Inventario_completo_{date.today()}.xlsx"),
            "Excel (*.xlsx)")
        if not ruta:
            return
        hojas = (self.inventario.hojas() + self.semana.hojas() + self.vencimientos.hojas())
        movs = self.inv.todos_los_movimientos()
        hojas.append(Hoja("Kardex de todos los medicamentos",
                          ["Medicamento", "Fecha", "Lote", "Concepto", "Entrada", "Salida", "Saldo", "Nota"],
                          [[f["nombre"], texto_fecha(f["fecha"]), f["lote"], f["concepto"], f["entrada"] or "",
                            f["salida"] or "", f["saldo"], f["nota"]] for f in movs], nombre_hoja="Kardex"))
        try:
            a_excel(hojas, Path(ruta))
        except PermissionError:
            QMessageBox.warning(self, "Archivo abierto", "Cierra el archivo en Excel e inténtalo otra vez.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(ruta))

    def respaldar(self):
        destino = self.inv.respaldar()
        QMessageBox.information(self, "Respaldo", f"Respaldo guardado en:\n{destino}")

    def restaurar(self):
        ruta, _ = QFileDialog.getOpenFileName(self, "Elegir respaldo", str(self.inv.ruta.parent / "respaldos"),
                                              "Respaldo (*.db)")
        if not ruta:
            return
        r = QMessageBox.warning(self, "Restaurar",
                                "Esto reemplaza TODOS los datos actuales por los del respaldo.\n"
                                "Antes se guarda un respaldo de lo actual. ¿Seguir?",
                                QMessageBox.Yes | QMessageBox.No)
        if r != QMessageBox.Yes:
            return
        self.inv.respaldar()
        self.inv.restaurar(Path(ruta))
        for i in range(self.tabs.count()):
            self.tabs.widget(i).recargar()
        QMessageBox.information(self, "Restaurar", "Datos restaurados.")

    # Actualizaciones

    def buscar_actualizacion(self, silencioso: bool):
        if not actualizador.configurado():
            if not silencioso:
                QMessageBox.information(self, "Actualizaciones",
                                        "Las actualizaciones automáticas todavía no están configuradas.")
            return
        self._silencioso = silencioso
        self._hilo = QThread(self)
        self._buscador = BuscadorActualizacion()
        self._buscador.moveToThread(self._hilo)
        self._hilo.started.connect(self._buscador.correr)
        self._buscador.listo.connect(self.resultado_actualizacion)
        self._buscador.listo.connect(self._hilo.quit)
        self._hilo.start()

    def resultado_actualizacion(self, novedad, error: str):
        if error:
            if not self._silencioso:
                QMessageBox.warning(self, "Actualizaciones",
                                    f"No se pudo revisar (¿hay internet?).\n\n{error}")
            return
        if not novedad:
            if not self._silencioso:
                QMessageBox.information(self, "Actualizaciones", f"Ya tienes la última versión ({VERSION}).")
            return
        r = QMessageBox.question(
            self, "Hay una versión nueva",
            f"Versión nueva: {novedad.version} (tienes {VERSION}).\n\n{novedad.notas[:800]}\n\n"
            "¿Descargar e instalar ahora? El programa se cerrará y se abrirá solo. Tus datos no se tocan.")
        if r != QMessageBox.Yes:
            return
        if not actualizador.es_exe():
            QMessageBox.information(self, "Actualizaciones",
                                    "Estás corriendo el código fuente; la actualización solo aplica al .exe.")
            return
        prog = QProgressDialog("Descargando…", None, 0, 100, self)
        prog.setWindowModality(Qt.WindowModal)
        prog.show()
        try:
            self.inv.respaldar()
            nuevo = actualizador.descargar(novedad, lambda p: (prog.setValue(p), QApplication.processEvents()))
        except Exception as e:
            prog.close()
            QMessageBox.warning(self, "Actualizaciones", f"Falló la descarga:\n{e}")
            return
        prog.close()
        actualizador.instalar_y_reiniciar(nuevo)
        self.close()

    def closeEvent(self, ev):
        # Antes de cerrar: tomar lo último del Excel y dejarlo al día.
        self.vigia_excel.stop()
        try:
            excel_vivo.importar(self.inv, self.ruta_excel)  # queda en morado para la próxima vez
            excel_vivo.exportar(self.inv, self.ruta_excel)
        except Exception:
            pass
        try:
            self.inv.respaldar()
        except Exception:
            pass
        self.inv.cerrar()
        ev.accept()


def paleta_clara():
    from PySide6.QtGui import QPalette
    p = QPalette()
    texto = QColor("#1B1F24")
    for grupo in (QPalette.Active, QPalette.Inactive):
        p.setColor(grupo, QPalette.Window, QColor("#F4F6F9"))
        p.setColor(grupo, QPalette.WindowText, texto)
        p.setColor(grupo, QPalette.Base, QColor("#FFFFFF"))
        p.setColor(grupo, QPalette.AlternateBase, QColor("#F5F9FC"))
        p.setColor(grupo, QPalette.Text, texto)
        p.setColor(grupo, QPalette.Button, QColor("#FFFFFF"))
        p.setColor(grupo, QPalette.ButtonText, texto)
        p.setColor(grupo, QPalette.ToolTipBase, QColor("#FFFFE1"))
        p.setColor(grupo, QPalette.ToolTipText, texto)
        p.setColor(grupo, QPalette.PlaceholderText, QColor("#8A96A3"))
        p.setColor(grupo, QPalette.Highlight, QColor("#BBD3EE"))
        p.setColor(grupo, QPalette.HighlightedText, texto)
        p.setColor(grupo, QPalette.BrightText, QColor("#B00020"))
        p.setColor(grupo, QPalette.Link, QColor("#1F4E79"))
    gris = QColor("#9AA5B1")
    for rol in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        p.setColor(QPalette.Disabled, rol, gris)
    p.setColor(QPalette.Disabled, QPalette.Base, QColor("#F0F0F0"))
    p.setColor(QPalette.Disabled, QPalette.Window, QColor("#F4F6F9"))
    p.setColor(QPalette.Disabled, QPalette.Button, QColor("#F0F0F0"))
    return p


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(NOMBRE_APP)
    # Siempre claro: si Windows o KDE están en modo oscuro, las letras quedaban blancas
    # sobre los fondos claros del programa.
    try:
        app.styleHints().setColorScheme(Qt.ColorScheme.Light)
    except AttributeError:
        pass
    app.setStyle("Fusion")
    app.setPalette(paleta_clara())
    app.setStyleSheet(ESTILO)
    f = QFont()
    f.setPointSize(10)
    app.setFont(f)
    try:
        from PySide6.QtCore import QLocale, QLibraryInfo, QTranslator
        QLocale.setDefault(QLocale(QLocale.Spanish, QLocale.Guatemala))
        tr = QTranslator(app)
        if tr.load(QLocale(), "qtbase", "_", QLibraryInfo.path(QLibraryInfo.TranslationsPath)):
            app.installTranslator(tr)
    except Exception:
        pass
    try:
        inv = Inventario()
        inv.respaldar()
    except Exception as e:
        QMessageBox.critical(None, NOMBRE_APP, f"No se pudo abrir la base de datos:\n{e}\n\n{carpeta_datos()}")
        return 1
    w = VentanaPrincipal(inv)
    w.showMaximized()
    return app.exec()
