"""Prueba de Inventario_actual.xlsx en los dos sentidos (Excel → programa).
Uso:  QT_QPA_PLATFORM=offscreen .venv/bin/python pruebas/prueba_excel.py
"""
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import tempfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path

os.environ["INVENTARIO_DATOS"] = tempfile.mkdtemp(prefix="inv-excel-")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook  # noqa: E402

from inventario import excel_vivo  # noqa: E402
from inventario.db import Inventario, lunes_de  # noqa: E402
from pruebas.prueba_movimientos import fallos as fallos_inv, invariantes  # noqa: E402

fallos = []


def revisar(c, t):
    print(("  ✔ " if c else "  ✗ ") + t)
    if not c:
        fallos.append(t)


def editar_excel(ruta, cambios, nuevas=(), quien="María López"):
    """Simula a una persona: abre, cambia casillas por encabezado y guarda."""
    wb = load_workbook(ruta)
    ws = wb["Inventario"]
    enc = {ws.cell(4, c).value: c for c in range(1, ws.max_column + 1)}
    for (fila_nombre, lote), col, valor in cambios:
        for r in range(5, ws.max_row + 1):
            if ws.cell(r, 1).value == fila_nombre and (lote is None or ws.cell(r, enc["Lote"]).value == lote):
                ws.cell(r, enc[col], valor)
                break
        else:
            raise AssertionError(f"no encontré {fila_nombre} {lote}")
    for datos in nuevas:
        r = ws.max_row + 1
        for col, valor in datos.items():
            ws.cell(r, enc[col], valor)
    wb.properties.lastModifiedBy = quien
    time.sleep(0.01)
    wb.save(ruta)


carpeta = Path(tempfile.mkdtemp(prefix="docs-"))
ruta = excel_vivo.ruta_por_defecto(carpeta)
inv = Inventario()
hoy_ef = excel_vivo.fecha_efectiva()
lun = lunes_de(hoy_ef) - timedelta(days=7)
para = inv.guardar_medicamento("Paracetamol 500 mg", "Analgésico", "Tableta", 10)
amox = inv.guardar_medicamento("Amoxicilina 500 mg", "Antibiótico", "Cápsula", 5)
sin = inv.guardar_medicamento("Vitamina C", "Sin lotes todavía")
la = inv.obtener_lote(para, "A1", "2027-01-31")
lb = inv.obtener_lote(amox, "B1", "2026-12-15")
inv.fijar_dia(la, lun.isoformat(), "ENTRADA", 100)
inv.fijar_dia(lb, lun.isoformat(), "ENTRADA", 40)

print("1) Exportar")
revisar(excel_vivo.exportar(inv, ruta), "se escribe el Excel")
revisar(not excel_vivo.hay_cambios(inv, ruta), "recién escrito no cuenta como cambio")
revisar(not excel_vivo.importar(inv, ruta).hay_algo, "importar sin cambios no hace nada")

print("2) María edita el Excel mientras en el programa registran una salida")
inv.fijar_dia(lb, (lun + timedelta(days=1)).isoformat(), "SALIDA", 5)  # programa: Amox B1 40 → 35
editar_excel(ruta, [
    (("Paracetamol 500 mg", "A1"), "Descripción", "Analgésico y antipirético"),
    (("Paracetamol 500 mg", "A1"), "Kardex", 90),
    (("Paracetamol 500 mg", "A1"), "Fecha de vencimiento", datetime(2027, 2, 28)),
    (("Paracetamol 500 mg", "A1"), "Existencias físicas", 88),
    (("Paracetamol 500 mg", "A1"), "Mínimo", "15"),
    (("Amoxicilina 500 mg", "B1"), "Kardex", 38),          # choque: el programa ya lo movió
    (("Amoxicilina 500 mg", "B1"), "Lote", "b1-x"),         # renombrar lote
    (("Vitamina C", None), "Lote", "VC9"),                  # lote nuevo a medicamento existente
    (("Vitamina C", None), "Kardex", 12),
], nuevas=[
    {"Medicamento": "Suero oral", "Descripción": "Sales", "Lote": "s-1", "Kardex": 30,
     "Fecha de vencimiento": "31/03/2027", "Mínimo": 5},
    {"Medicamento": "Malo", "Kardex": "muchos"},
    {"Medicamento": "Malo 2", "Fecha de vencimiento": "mañana"},
])
res = excel_vivo.importar(inv, ruta)
print("   aplicados:", *res.aplicados, sep="\n     ")
print("   choques:", *res.choques, sep="\n     ")
print("   errores:", *res.errores, sep="\n     ")
revisar(res.usuario == "María López", "se sabe quién lo cambió (María López)")
m = inv.medicamento(para)
revisar(m["descripcion"] == "Analgésico y antipirético" and m["minimo"] == 15, "descripción y mínimo aplicados")
revisar(inv.saldo_lote(la) == 90, "kardex 100 → 90 con ajuste")
k = inv.kardex(para)[-1]
revisar(k.concepto == "Ajuste desde Excel" and k.salida == 10 and "María López" in k.nota, "kardex muestra ajuste −10 por María")
revisar(inv.lotes(para)[0]["vencimiento"] == "2027-02-28", "vencimiento escrito como fecha de Excel")
fila = [f for f in inv.inventario_por_lote() if f["lote_id"] == la][0]
revisar(fila["fisico"] == 88, "existencias físicas 88 (conteo desde Excel)")
revisar(inv.saldo_lote(lb) == 35 and len(res.choques) == 1, "choque: kardex de Amoxicilina se queda en 35 y se avisa")
revisar(inv.lotes(amox)[0]["lote"] == "B1-X", "lote renombrado a B1-X")
vc = inv.lotes(sin)
revisar(len(vc) == 1 and vc[0]["lote"] == "VC9" and vc[0]["existencia"] == 12, "Vitamina C: lote VC9 con 12")
suero = inv.medicamento_por_nombre("Suero oral")
revisar(suero and inv.lotes(suero["id"])[0]["existencia"] == 30
        and inv.lotes(suero["id"])[0]["vencimiento"] == "2027-03-31", "fila nueva: Suero oral, lote S-1, 30")
revisar(not inv.medicamento_por_nombre("Malo") and not inv.medicamento_por_nombre("Malo 2")
        and len(res.errores) == 2, "filas con datos malos se rechazan con aviso")
pend = inv.cambios_pendientes()
revisar(pend and all(c["usuario"] == "María López" and c["origen"] == "EXCEL" for c in pend),
        f"{len(pend)} cambios anotados con quién y cuándo")

print("3) Volver a guardar el mismo Excel no repite nada")
saldos = {l: inv.saldo_lote(l) for (l,) in inv.cx.execute("SELECT id FROM lotes")}
n_cambios = len(inv.historial_cambios())
wb = load_workbook(ruta); time.sleep(0.01); wb.save(ruta)
res2 = excel_vivo.importar(inv, ruta)
revisar(not res2.aplicados, "nada se aplica dos veces")
revisar(saldos == {l: inv.saldo_lote(l) for (l,) in inv.cx.execute("SELECT id FROM lotes")}, "saldos iguales")
revisar(len(inv.historial_cambios()) == n_cambios, "no se duplican anotaciones")

print("4) Reescribir: casillas moradas con comentario")
revisar(excel_vivo.exportar(inv, ruta), "se reescribe")
wb = load_workbook(ruta); ws = wb["Inventario"]
enc = {ws.cell(4, c).value: c for c in range(1, ws.max_column + 1)}
r = next(r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value == "Paracetamol 500 mg")
c = ws.cell(r, enc["Kardex"])
revisar(c.fill.fgColor.rgb.endswith(excel_vivo.MORADO) and c.comment and "María López" in c.comment.text,
        "Kardex de Paracetamol morado con comentario de María")
revisar(ws.cell(r, enc["Lote"]).fill.fgColor.rgb.endswith("FFFFFF") or not ws.cell(r, enc["Lote"]).comment,
        "casilla sin cambios no está morada")
r2 = next(r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value == "Suero oral")
revisar(ws.cell(r2, enc["ID lote"]).value, "fila nueva ya tiene su ID")
revisar("Cambios" in wb.sheetnames and wb["Cambios"].max_row > 5, "hoja «Cambios» con el historial")

print("5) Después de reescribir, una edición normal sigue funcionando")
inv.fijar_dia(la, (lun + timedelta(days=2)).isoformat(), "SALIDA", 5)  # 90 → 85 en el programa
excel_vivo.exportar(inv, ruta)
editar_excel(ruta, [(("Paracetamol 500 mg", "A1"), "Kardex", 80)], quien="Pedro")
res3 = excel_vivo.importar(inv, ruta)
revisar(inv.saldo_lote(la) == 80 and not res3.choques and res3.usuario == "Pedro", "Pedro baja a 80 sin choque")
editar_excel(ruta, [(("Paracetamol 500 mg", "A1"), "Kardex", 9999)], quien="Pedro")
excel_vivo.importar(inv, ruta)
revisar(inv.saldo_lote(la) == 9999, "subir kardex también funciona (ajuste positivo)")

print("6) Revisado quita el morado")
inv.marcar_revisados([c["id"] for c in inv.cambios_pendientes()])
excel_vivo.exportar(inv, ruta)
wb = load_workbook(ruta); ws = wb["Inventario"]
morados = sum(1 for row in ws.iter_rows(min_row=5) for c in row
              if c.fill and c.fill.fgColor and str(c.fill.fgColor.rgb).endswith(excel_vivo.MORADO))
revisar(morados == 0, "ninguna casilla morada")

print("7) Excel abierto (archivo de bloqueo) → no se sobrescribe")
(carpeta / f"~${ruta.name}").write_text("x")
antes = ruta.stat().st_mtime_ns
revisar(not excel_vivo.exportar(inv, ruta) and ruta.stat().st_mtime_ns == antes, "no toca el archivo abierto")
(carpeta / f"~${ruta.name}").unlink()

print("8) Cuentas siguen cuadrando")
invariantes(inv, "excel")
revisar(not fallos_inv, "invariantes del inventario")

print("9) Pantalla: el programa detecta el Excel solo y pinta de morado")
from PySide6.QtCore import QStandardPaths  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402
QStandardPaths.setTestModeEnabled(True)
from inventario import ui  # noqa: E402
app = QApplication([])
w = ui.VentanaPrincipal(inv)
w.ruta_excel = ruta
w.revisar_excel()  # escribe el Excel al día
editar_excel(ruta, [(("Amoxicilina 500 mg", "B1-X"), "Presentación", "Cápsula dura")], quien="Ana")
w.revisar_excel(); w.revisar_excel()  # 1º ve el cambio, 2º confirma que ya no cambia e importa
revisar(inv.medicamento(amox)["presentacion"] == "Cápsula dura", "el programa tomó el cambio solo")
t = w.inventario.t
fila = next(r for r, f in enumerate(w.inventario.filas) if f["med_id"] == amox)
it = t.item(fila, 7)
revisar(it.background().color().name().upper().endswith(excel_vivo.MORADO), "casilla Presentación morada")
revisar("Ana" in it.toolTip(), f"al pasar el mouse dice quién: {it.toolTip().splitlines()[0]}")
abiertos = []
ui.DialogoCambios.exec = lambda self: abiertos.append(self) or 0
w.inventario.clic_celda(fila, 7)
revisar(abiertos and "Ana" in abiertos[0].findChild(ui.QTableWidget).item(0, 1).text(), "clic abre detalle con Ana")
revisar(w.inventario.aviso.isVisible() or not w.inventario.aviso.isHidden(), "aviso morado arriba")
w.close()
print(f"\n{len(fallos)} fallos")
sys.exit(1 if fallos else 0)
