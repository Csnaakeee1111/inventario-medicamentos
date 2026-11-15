"""Prueba de las pantallas sin mostrarlas: diálogos, casillas de la semana, cuadre y exportes.
Uso:  QT_QPA_PLATFORM=offscreen .venv/bin/python pruebas/prueba_pantallas.py
"""
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import tempfile
from datetime import date, timedelta
from pathlib import Path

os.environ["INVENTARIO_DATOS"] = tempfile.mkdtemp(prefix="inv-pantallas-")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from PySide6.QtCore import QDate, QStandardPaths  # noqa: E402
from PySide6.QtGui import QDesktopServices  # noqa: E402
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox  # noqa: E402

QStandardPaths.setTestModeEnabled(True)  # exportes a una carpeta de prueba, no a Documentos reales
avisos = []
QMessageBox.warning = staticmethod(lambda *a, **k: avisos.append(a[2]))
QMessageBox.information = staticmethod(lambda *a, **k: None)
QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.No)
QDesktopServices.openUrl = staticmethod(lambda u: True)
salida = Path(tempfile.mkdtemp(prefix="inv-exportes-"))

from inventario import ui  # noqa: E402
from inventario.db import Inventario, lunes_de  # noqa: E402

fallos = []


def revisar(c, t):
    print(("  ✔ " if c else "  ✗ ") + t)
    if not c:
        fallos.append(t)


app = QApplication([])
inv = Inventario()
w = ui.VentanaPrincipal(inv)
w.show()
lun = lunes_de(date.today()) - timedelta(days=7)
mar = lun + timedelta(days=1)

# Diálogo de medicamento
d = ui.DialogoMedicamento(inv)
d.nombre.setText("Amoxicilina 500 mg"); d.descripcion.setPlainText("Antibiótico"); d.minimo.setValue(10)
d.guardar(); revisar(d.result() == 1, "crear medicamento desde el diálogo")
d = ui.DialogoMedicamento(inv); d.nombre.setText("amoxicilina 500 MG"); avisos.clear(); d.guardar()
revisar(d.result() == 0 and avisos, "diálogo rechaza nombre repetido")


def movimiento(tipo, med, lote, cant, fecha, vence=None):
    avisos.clear()
    m = ui.DialogoMovimiento(inv, fecha, tipo=tipo, medicamento=med)
    m.lote.setEditText(lote)
    if vence:
        m.sin_vence.setChecked(False); m.vence.setDate(QDate(vence.year, vence.month, vence.day))
    m.cantidad.setValue(cant)
    m.guardar()
    return m.result() == 1


revisar(movimiento("ENTRADA", "Amoxicilina 500 mg", "amx-1", 50, lun, date(2027, 5, 1)), "entrada con lote nuevo")
l = inv.lotes(inv.medicamento_por_nombre("Amoxicilina 500 mg")["id"])
revisar(len(l) == 1 and l[0]["lote"] == "AMX-1" and l[0]["vencimiento"] == "2027-05-01" and l[0]["existencia"] == 50,
        "lote guardado en mayúsculas con vencimiento y 50 unidades")
revisar(movimiento("ENTRADA", "Amoxicilina 500 mg", "AMX-1", 10, lun), "segunda entrada mismo lote y día se suma")
revisar(inv.lotes(l[0]["medicamento_id"])[0]["existencia"] == 60, "existencia 60")
m = ui.DialogoMovimiento(inv, mar, tipo="SALIDA", medicamento="Amoxicilina 500 mg")
revisar("AMX-1" in m.lote.currentText() and "hay 60" in m.lote.currentText(), "salida ofrece el lote con existencia")
m.cantidad.setValue(15); m.guardar(); revisar(m.result() == 1, "salida eligiendo el lote de la lista")
revisar(not movimiento("SALIDA", "Amoxicilina 500 mg", "NO-EXISTE", 1, mar), "salida de lote inexistente rechazada")
revisar(not movimiento("SALIDA", "Amoxicilina 500 mg", "AMX-1", 999, mar), "salida mayor a existencia rechazada")
revisar(not movimiento("ENTRADA", "Medicina fantasma", "X", 1, lun), "medicamento que no existe rechazado")
revisar(not movimiento("ENTRADA", "Amoxicilina 500 mg", "AMX-1", 1, lun + timedelta(days=5)), "sábado rechazado")
revisar(movimiento("ENTRADA", "amoxicilina 500 mg", "", 5, lun), "nombre en minúsculas y sin lote → SIN LOTE")

# Tabla de la semana: escribir en las casillas
w.algo_cambio()
w.semana.lunes = lun; w.semana.recargar()
t = w.semana.t
fila = next(r for r, f in enumerate(w.semana.filas) if f["lote"] == "AMX-1")
col_fin = t.columnCount() - 1
revisar(t.item(fila, col_fin).text() == "45", "semana: final 60 − 15 = 45")
t.item(fila, 4 + 2 * 2 + 1).setText("5")  # miércoles sale 5
fila = next(r for r, f in enumerate(w.semana.filas) if f["lote"] == "AMX-1")
revisar(w.semana.t.item(fila, col_fin).text() == "40", "semana: escribir salida del miércoles → 40")
avisos.clear(); w.semana.t.item(fila, 4 + 2 * 3 + 1).setText("41")
fila = next(r for r, f in enumerate(w.semana.filas) if f["lote"] == "AMX-1")
revisar(avisos and w.semana.t.item(fila, 4 + 2 * 3 + 1).text() == "", "semana: salida que no alcanza se rechaza y se borra")
w.semana.t.item(fila, 4 + 2 * 2 + 1).setText("")
fila = next(r for r, f in enumerate(w.semana.filas) if f["lote"] == "AMX-1")
revisar(w.semana.t.item(fila, col_fin).text() == "45", "semana: borrar la casilla quita el movimiento")
w.semana.mover(7)
fila = next(r for r, f in enumerate(w.semana.filas) if f["lote"] == "AMX-1")
revisar(w.semana.t.item(fila, 3).text() == "45", "semana siguiente arranca con inicial 45")
w.semana.hoy(); revisar(w.semana.lunes == lunes_de(date.today()), "botón «Esta semana»")

# Cuadre: escribir conteos y ver diferencias
w.cuadre.fecha.setDate(QDate(lun.year, lun.month, lun.day).addDays(4))
ct = w.cuadre.t
r = next(i for i, f in enumerate(w.cuadre.filas) if f["lote"] == "AMX-1")
ct.item(r, 4).setText("43")
revisar(ct.item(r, 5).text() == "-2", "cuadre: diferencia −2")
ct.item(r, 4).setText("45")
revisar(ct.item(r, 5).text() == "0 ✔", "cuadre: diferencia 0")

# Kardex
w.kardex.medicamento.setText("Amoxicilina 500 mg")
revisar(w.kardex.t.rowCount() == 3 and w.kardex.t.item(2, 6).text() == "50", "kardex: 3 filas y saldo final 50")

# Inventario: búsqueda
w.inventario.buscar.setText("amx")
revisar(w.inventario.t.rowCount() == 1, "inventario: buscar por lote")
w.inventario.buscar.setText("")

# Exportes de todas las pestañas
for p in (w.semana, w.inventario, w.kardex, w.cuadre, w.vencimientos):
    for fmt in ("xlsx", "pdf"):
        destino = salida / f"{p.nombre_archivo()}.{fmt}"
        QFileDialog.getSaveFileName = staticmethod(lambda *a, _d=destino, **k: (str(_d), ""))
        p.exportar(fmt)
        hojas = p.hojas()
        if hojas and any(h.filas for h in hojas):
            revisar(destino.exists() and destino.stat().st_size > 1000, f"exporta {type(p).__name__} a {fmt}")
destino = salida / "todo.xlsx"
QFileDialog.getSaveFileName = staticmethod(lambda *a, **k: (str(destino), ""))
w.exportar_todo()
import openpyxl  # noqa: E402
revisar(openpyxl.load_workbook(destino).sheetnames == ["Inventario", "Semana", "Vencimientos", "Kardex"],
        "Excel completo con 4 hojas")
w.close()
revisar((Path(QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)) /
         "Inventario Medicamentos" / "Inventario_actual.xlsx").exists(), "al cerrar deja Inventario_actual.xlsx")
print(f"\n{len(fallos)} fallos")
sys.exit(1 if fallos else 0)
