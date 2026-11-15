"""Prueba de movimientos: casos puntuales + miles de movimientos al azar.

Después de cada movimiento revisa que las cuentas cuadren en todas las vistas.
Uso:  .venv/bin/python pruebas/prueba_movimientos.py
"""
from __future__ import annotations

import os
import random
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import tempfile
from datetime import date, timedelta
from pathlib import Path

os.environ["INVENTARIO_DATOS"] = tempfile.mkdtemp(prefix="inv-prueba-")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from inventario.db import ErrorInventario, Inventario, dias_habiles, lunes_de  # noqa: E402

fallos: list[str] = []
pasos = 0


def revisar(cond: bool, texto: str):
    global pasos
    pasos += 1
    if not cond:
        fallos.append(texto)
        print("  ✗", texto)


def debe_fallar(func, texto: str):
    try:
        func()
    except ErrorInventario:
        revisar(True, texto)
        return
    revisar(False, f"debía rechazarse: {texto}")


# ── Revisión de que todo cuadra ─────────────────────────────────────────────

def invariantes(inv: Inventario, contexto: str):
    cx = inv.cx
    # 1. Ningún lote con saldo negativo en ninguna fecha.
    for (lote_id,) in cx.execute("SELECT id FROM lotes"):
        saldo = 0
        for f, s in cx.execute("""SELECT fecha, SUM(CASE tipo WHEN 'SALIDA' THEN -cantidad ELSE cantidad END)
                                  FROM movimientos WHERE lote_id=? GROUP BY fecha ORDER BY fecha""", (lote_id,)):
            saldo += s
            if saldo < 0:
                revisar(False, f"[{contexto}] lote {lote_id} negativo ({saldo}) el {f}")
                break
    # 2. Sin fines de semana ni cantidades raras.
    for f, t, c in cx.execute("SELECT fecha, tipo, cantidad FROM movimientos"):
        if date.fromisoformat(f).weekday() >= 5:
            revisar(False, f"[{contexto}] movimiento en fin de semana {f}")
        if t in ("ENTRADA", "SALIDA") and c <= 0:
            revisar(False, f"[{contexto}] {t} con cantidad {c}")
        if t == "AJUSTE" and c == 0:
            revisar(False, f"[{contexto}] ajuste en cero")
    # 3. Kardex = suma de lotes = tabla de inventario = kardex de todos.
    inventario = {}
    for f in inv.inventario_por_lote(incluir_inactivos=True, incluir_vacios=True):
        inventario[f["med_id"]] = inventario.get(f["med_id"], 0) + (f["kardex"] or 0)
    todos = {}
    for f in inv.todos_los_movimientos():
        todos[f["nombre"]] = f["saldo"]
    for m in inv.medicamentos(incluir_inactivos=True):
        k = inv.kardex(m["id"])
        final = k[-1].saldo if k else 0
        lotes = sum(l["existencia"] for l in inv.lotes(m["id"]))
        if not (final == lotes == m["existencia"] == inventario.get(m["id"], 0) == todos.get(m["nombre"], 0)):
            revisar(False, f"[{contexto}] {m['nombre']}: kardex {final}, lotes {lotes}, catálogo "
                           f"{m['existencia']}, inventario {inventario.get(m['id'])}, todos {todos.get(m['nombre'])}")
        # Kardex por lote y con fechas debe terminar igual.
        for l in inv.lotes(m["id"]):
            kl = inv.kardex(m["id"], l["id"], desde="2026-01-01")
            if kl[-1].saldo != l["existencia"]:
                revisar(False, f"[{contexto}] kardex del lote {l['lote']} con fecha termina en {kl[-1].saldo}, "
                               f"debía {l['existencia']}")
    revisar(True, f"[{contexto}] invariantes")


def semanas_encadenan(inv: Inventario, desde: date, semanas: int, contexto: str):
    """El Final de una semana tiene que ser el Inicial de la siguiente, lote por lote."""
    lunes = lunes_de(desde)
    anterior = None
    for _ in range(semanas):
        filas = {f["lote_id"]: f for f in inv.filas_semana(lunes, solo_con_existencia=False)}
        for lid, f in filas.items():
            calc = f["inicial"] + sum(d["entrada"] - d["salida"] + d["ajuste"] for d in f["dias"])
            if calc != f["final"]:
                revisar(False, f"[{contexto}] semana {lunes} lote {lid}: final {f['final']} ≠ cuenta {calc}")
            if f["final"] != inv.saldo_lote(lid, hasta=(lunes + timedelta(days=4)).isoformat()):
                revisar(False, f"[{contexto}] semana {lunes} lote {lid}: final ≠ saldo del viernes")
            if anterior and lid in anterior and anterior[lid]["final"] != f["inicial"]:
                revisar(False, f"[{contexto}] semana {lunes} lote {lid}: inicial {f['inicial']} ≠ final "
                               f"anterior {anterior[lid]['final']}")
        anterior = filas
        lunes += timedelta(days=7)
    revisar(True, f"[{contexto}] semanas encadenan")


# ── 1. Casos puntuales ──────────────────────────────────────────────────────

def casos_puntuales():
    print("1) Casos puntuales")
    inv = Inventario(Path(os.environ["INVENTARIO_DATOS"]) / "puntual.db")
    lun = date(2026, 11, 9)
    L, M, X, J, V = [d.isoformat() for d in dias_habiles(lun)]
    sab, dom = "2026-11-14", "2026-11-15"

    para = inv.guardar_medicamento("Paracetamol 500 mg", "Analgésico", "Tableta", 10)
    debe_fallar(lambda: inv.guardar_medicamento("  paracetamol   500 MG "), "nombre repetido (mayúsculas/espacios)")
    debe_fallar(lambda: inv.guardar_medicamento("   "), "nombre vacío")
    a = inv.obtener_lote(para, "a-1", "2027-01-01")
    revisar(inv.obtener_lote(para, "A-1", None) == a, "mismo lote sin importar mayúsculas")
    revisar(inv.lotes(para)[0]["vencimiento"] == "2027-01-01", "vencimiento no se borra si no se manda")
    sl = inv.obtener_lote(para, "", None)
    revisar(inv.lotes(para)[-1]["lote"] == "SIN LOTE" and sl, "lote vacío = SIN LOTE")

    inv.fijar_dia(a, L, "ENTRADA", 100)
    inv.fijar_dia(a, M, "SALIDA", 30)
    revisar(inv.saldo_lote(a) == 70, "entrada 100 − salida 30 = 70")
    inv.sumar_dia(a, M, "SALIDA", 5)
    revisar(inv.cantidad_dia(a, M, "SALIDA") == 35, "sumar a la salida del mismo día = una sola fila")
    revisar(inv.cx.execute("SELECT COUNT(*) FROM movimientos WHERE lote_id=? AND fecha=?", (a, M)).fetchone()[0] == 1,
            "no se duplican filas del mismo día")
    debe_fallar(lambda: inv.fijar_dia(a, X, "SALIDA", 66), "salida mayor que la existencia")
    inv.fijar_dia(a, X, "SALIDA", 65)
    revisar(inv.saldo_lote(a) == 0, "sacar exactamente todo deja 0")
    debe_fallar(lambda: inv.fijar_dia(a, L, "ENTRADA", 99), "bajar una entrada pasada que ya se usó")
    debe_fallar(lambda: inv.fijar_dia(a, L, "ENTRADA", 0), "borrar una entrada pasada que ya se usó")
    inv.fijar_dia(a, X, "SALIDA", 60)
    inv.fijar_dia(a, L, "ENTRADA", 95)
    revisar(inv.saldo_lote(a) == 0, "bajar entrada pasada cuando sí alcanza")
    inv.fijar_dia(a, X, "SALIDA", 0)
    revisar(inv.cantidad_dia(a, X, "SALIDA") == 0 and inv.saldo_lote(a) == 60, "poner 0 borra el movimiento")
    debe_fallar(lambda: inv.fijar_dia(a, sab, "ENTRADA", 1), "sábado")
    debe_fallar(lambda: inv.fijar_dia(a, dom, "SALIDA", 1), "domingo")
    debe_fallar(lambda: inv.fijar_dia(a, L, "ENTRADA", -5), "cantidad negativa")
    debe_fallar(lambda: inv.sumar_dia(a, L, "ENTRADA", 0), "sumar cero")
    debe_fallar(lambda: inv.fijar_dia(sl, L, "SALIDA", 1), "salida de lote vacío")

    # Cuadre semanal con ajuste.
    debe_fallar(lambda: inv.guardar_cuadre(sab, "SEMANAL", {a: 60}, ajustar=True), "cuadre en sábado")
    debe_fallar(lambda: inv.guardar_cuadre(V, "SEMANAL", {}, ajustar=True), "cuadre vacío")
    debe_fallar(lambda: inv.guardar_cuadre(V, "SEMANAL", {a: -1}, ajustar=False), "conteo negativo")
    inv.guardar_cuadre(V, "SEMANAL", {a: 57}, ajustar=True)
    revisar(inv.saldo_lote(a) == 57, "ajuste baja el kardex a lo contado (60 → 57)")
    k = inv.kardex(para)
    revisar(k[-1].concepto == "Ajuste por cuadre" and k[-1].salida == 3, "kardex muestra ajuste −3")
    inv.guardar_cuadre(V, "SEMANAL", {a: 57}, ajustar=True)
    revisar(inv.saldo_lote(a) == 57, "segundo cuadre igual no cambia nada")
    revisar(inv.cx.execute("SELECT COUNT(*) FROM movimientos WHERE tipo='AJUSTE' AND cantidad=0").fetchone()[0] == 0,
            "no se crean ajustes en cero")
    inv.guardar_cuadre(V, "MENSUAL", {a: 70}, ajustar=False)
    revisar(inv.saldo_lote(a) == 57, "cuadre sin ajustar no toca el kardex")
    fila = [f for f in inv.inventario_por_lote() if f["lote_id"] == a][0]
    revisar(fila["fisico"] == 70 and fila["kardex"] == 57, "inventario muestra físico 70 y kardex 57")

    # Cuadre con fecha pasada que chocaría con salidas posteriores.
    sig = [d.isoformat() for d in dias_habiles(lun + timedelta(days=7))]
    inv.fijar_dia(a, sig[0], "SALIDA", 50)  # quedan 7
    debe_fallar(lambda: inv.guardar_cuadre(V, "SEMANAL", {a: 40}, ajustar=True),
                "ajuste pasado que dejaría negativa una salida posterior")
    inv.guardar_cuadre(V, "SEMANAL", {a: 50}, ajustar=True)
    revisar(inv.saldo_lote(a) == 0, "ajuste pasado que sí alcanza")

    # Lotes y catálogo.
    b = inv.obtener_lote(para, "B-2", "2026-12-01")
    debe_fallar(lambda: inv.editar_lote(b, "a-1", None), "renombrar lote a uno que ya existe")
    inv.editar_lote(b, "b-3", "2026-12-31")
    revisar(inv.lotes(para)[0]["lote"] in ("B-3", "A-1"), "renombrar lote")
    ibu = inv.guardar_medicamento("Ibuprofeno")
    debe_fallar(lambda: inv.guardar_medicamento("Paracetamol 500 mg", medicamento_id=ibu), "renombrar a uno existente")
    inv.activar_medicamento(ibu, False)
    debe_fallar(lambda: inv.guardar_medicamento("ibuprofeno"), "agregar uno que está desactivado")
    revisar("Ibuprofeno" not in inv.nombres_medicamentos(), "desactivado no sale en la lista")

    # Respaldo y restaurar.
    copia = inv.respaldar()
    inv.fijar_dia(b, sig[1], "ENTRADA", 999)
    inv.restaurar(copia)
    revisar(inv.saldo_lote(b) == 0, "restaurar respaldo regresa los datos")
    invariantes(inv, "puntual")
    semanas_encadenan(inv, lun - timedelta(days=7), 4, "puntual")
    inv.cerrar()


# ── 2. Miles de movimientos al azar ─────────────────────────────────────────

def al_azar(semilla: int, n: int = 3000):
    print(f"2) {n} movimientos al azar (semilla {semilla})")
    rnd = random.Random(semilla)
    inv = Inventario(Path(os.environ["INVENTARIO_DATOS"]) / f"azar{semilla}.db")
    meds = [inv.guardar_medicamento(f"Medicina {i}", f"desc {i}") for i in range(6)]
    lotes = [inv.obtener_lote(m, f"L{j}", f"2027-0{j + 1}-15") for m in meds for j in range(rnd.randint(1, 3))]
    inicio = date(2026, 9, 7)
    habiles = [inicio + timedelta(days=i) for i in range(120) if (inicio + timedelta(days=i)).weekday() < 5]
    aceptados = rechazados = 0
    for i in range(n):
        lote = rnd.choice(lotes)
        dia = rnd.choice(habiles).isoformat()
        op = rnd.random()
        try:
            if op < 0.35:
                inv.fijar_dia(lote, dia, "ENTRADA", rnd.choice([0, 1, 5, 10, 20, 50, 100]))
            elif op < 0.75:
                inv.fijar_dia(lote, dia, "SALIDA", rnd.choice([0, 1, 2, 3, 5, 10, 30]))
            elif op < 0.9:
                inv.sumar_dia(lote, dia, rnd.choice(["ENTRADA", "SALIDA"]), rnd.randint(1, 15))
            else:
                sistema = inv.saldo_lote(lote, hasta=dia)
                inv.guardar_cuadre(dia, rnd.choice(["SEMANAL", "MENSUAL"]),
                                   {lote: max(0, sistema + rnd.randint(-5, 5))}, ajustar=rnd.random() < 0.7)
            aceptados += 1
        except ErrorInventario:
            rechazados += 1
        if i % 150 == 0:
            invariantes(inv, f"azar{semilla} paso {i}")
    invariantes(inv, f"azar{semilla} final")
    semanas_encadenan(inv, inicio, 18, f"azar{semilla}")
    print(f"   aceptados {aceptados}, rechazados (correctamente) {rechazados}")
    inv.cerrar()


if __name__ == "__main__":
    casos_puntuales()
    for s in (1, 2, 3):
        al_azar(s)
    print(f"\n{pasos} revisiones · {len(fallos)} fallos")
    sys.exit(1 if fallos else 0)
