"""Inventario_actual.xlsx en los dos sentidos.

- `exportar`: escribe el Excel y guarda una "foto" de cada valor escrito.
- `importar`: si el archivo se guardó después, compara cada casilla con la foto.
  Lo que la gente cambió se aplica en el programa y queda anotado (cuándo, quién, antes, después).

Se compara contra la foto (lo que había en el Excel), no contra la base: así un Excel viejo
no deshace lo que se registró en el programa mientras tanto. Si el mismo dato cambió en los
dos lados, no se aplica y se avisa como choque.
"""
from __future__ import annotations

import getpass
import os
import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .db import ErrorInventario, Inventario, es_habil, texto_fecha

MARCA = "Programa Inventario"
MORADO = "E4D7F5"
GRIS = "E7E9EC"
AZUL = "1F4E79"

# (encabezado, campo, nivel, editable)
COLUMNAS = [
    ("Medicamento", "nombre", "med", True),
    ("Descripción", "descripcion", "med", True),
    ("Existencias físicas", "fisico", "lote", True),
    ("Kardex", "kardex", "lote", True),
    ("Lote", "lote", "lote", True),
    ("Fecha de vencimiento", "vencimiento", "lote", True),
    ("Presentación", "presentacion", "med", True),
    ("Mínimo", "minimo", "med", True),
    ("Diferencia", "diferencia", "-", False),
    ("ID medicamento", "med_id", "-", False),
    ("ID lote", "lote_id", "-", False),
]
NOMBRES_CAMPO = {"nombre": "Medicamento", "descripcion": "Descripción", "fisico": "Existencias físicas",
                 "kardex": "Kardex", "lote": "Lote", "vencimiento": "Fecha de vencimiento",
                 "presentacion": "Presentación", "minimo": "Mínimo", "nuevo": "Agregado"}
CAMPOS_NUMERO = {"fisico", "kardex", "minimo"}


@dataclass
class Resultado:
    usuario: str = ""
    aplicados: list[str] = field(default_factory=list)
    choques: list[str] = field(default_factory=list)
    errores: list[str] = field(default_factory=list)

    @property
    def hay_algo(self) -> bool:
        return bool(self.aplicados or self.choques or self.errores)


def ruta_por_defecto(carpeta: Path) -> Path:
    return carpeta / "Inventario_actual.xlsx"


def abierto_en_otro_programa(ruta: Path) -> bool:
    """Excel deja ~$archivo.xlsx y LibreOffice .~lock.archivo.xlsx# mientras está abierto."""
    return ((ruta.parent / f"~${ruta.name}").exists()
            or (ruta.parent / f".~lock.{ruta.name}#").exists())


def fecha_efectiva(hoy: date | None = None) -> date:
    """Hoy si es día hábil; si es fin de semana, el viernes anterior."""
    hoy = hoy or date.today()
    return hoy if es_habil(hoy) else hoy - timedelta(days=hoy.weekday() - 4)


# ── Normalizar lo que viene del Excel ──────────────────────────────────────

def normalizar(campo: str, valor) -> str:
    """Convierte un valor de casilla en texto comparable. Lanza ValueError si no sirve."""
    if valor is None:
        return ""
    if campo == "vencimiento":
        if isinstance(valor, datetime):
            return valor.date().isoformat()
        if isinstance(valor, date):
            return valor.isoformat()
        texto = str(valor).strip()
        if not texto:
            return ""
        for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(texto, formato).date().isoformat()
            except ValueError:
                pass
        raise ValueError(f"fecha no válida «{texto}» (usa dd/mm/aaaa)")
    if campo in CAMPOS_NUMERO:
        if isinstance(valor, str):
            valor = valor.strip()
            if not valor:
                return ""
        try:
            numero = float(valor)
        except (TypeError, ValueError):
            raise ValueError(f"«{valor}» no es un número") from None
        if numero != int(numero) or numero < 0:
            raise ValueError(f"«{valor}» debe ser un número entero sin negativos")
        return str(int(numero))
    texto = " ".join(str(valor).split()) if campo in ("nombre", "lote", "presentacion") else str(valor).strip()
    return texto.upper() if campo == "lote" else texto


def mostrar(campo: str, valor: str) -> str:
    if campo == "vencimiento":
        return texto_fecha(valor) or "(vacío)"
    return valor if valor != "" else "(vacío)"


def _valores_actuales(inv: Inventario) -> tuple[dict[int, dict], dict[int, dict]]:
    meds, lotes = {}, {}
    for f in inv.inventario_por_lote(incluir_inactivos=True, incluir_vacios=True):
        meds[f["med_id"]] = {
            "nombre": f["nombre"], "descripcion": f["descripcion"] or "",
            "presentacion": f["presentacion"] or "", "minimo": str(f["minimo"] or 0),
        }
        if f["lote_id"]:
            lotes[f["lote_id"]] = {
                "med_id": f["med_id"], "lote": f["lote"], "vencimiento": f["vencimiento"] or "",
                "kardex": str(f["kardex"] or 0), "fisico": "" if f["fisico"] is None else str(f["fisico"]),
            }
    return meds, lotes


def _clave(nivel: str, ident: int, campo: str) -> str:
    return f"{'m' if nivel == 'med' else 'l'}{ident}.{campo}"


# ── Exportar ────────────────────────────────────────────────────────────────

def exportar(inv: Inventario, ruta: Path) -> bool:
    """Escribe el Excel. Devuelve False si está abierto en Excel/LibreOffice (no se toca)."""
    if abierto_en_otro_programa(ruta):
        return False
    pendientes: dict[tuple[str, int, str], list] = {}
    for c in inv.cambios_pendientes():
        if c["lote_id"] and c["campo"] in ("lote", "vencimiento", "fisico", "kardex"):
            pendientes.setdefault(("lote", c["lote_id"], c["campo"]), []).append(c)
        elif c["medicamento_id"]:
            campo = "nombre" if c["campo"] == "nuevo" else c["campo"]
            pendientes.setdefault(("med", c["medicamento_id"], campo), []).append(c)

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"
    borde = Side(style="thin", color="999999")
    caja = Border(left=borde, right=borde, top=borde, bottom=borde)
    ws["A1"] = "Inventario de medicamentos"
    ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    ws["A2"] = (f"Actualizado {datetime.now():%d/%m/%Y %H:%M}. Puedes cambiar las casillas blancas y "
                "agregar filas nuevas abajo; guarda y el programa lo toma solo. Morado = cambiado desde "
                "el Excel (pasa el mouse para ver quién y cuándo). Gris = no se edita.")
    ws["A2"].font = Font(italic=True, color="555555")
    fila_enc = 4
    for c, (enc, campo, _, editable) in enumerate(COLUMNAS, 1):
        celda = ws.cell(fila_enc, c, enc)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=AZUL if editable else "6B7B8C")
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = caja
    foto: dict[str, str] = {}
    fila = fila_enc + 1
    for f in inv.inventario_por_lote():
        med_id, lote_id = f["med_id"], f["lote_id"]
        valores = {
            "nombre": f["nombre"], "descripcion": f["descripcion"] or "",
            "fisico": "" if f["fisico"] is None else str(f["fisico"]),
            "kardex": str(f["kardex"] or 0) if lote_id else "",
            "lote": f["lote"] or "", "vencimiento": f["vencimiento"] or "",
            "presentacion": f["presentacion"] or "", "minimo": str(f["minimo"] or 0),
        }
        for c, (enc, campo, nivel, editable) in enumerate(COLUMNAS, 1):
            celda = ws.cell(fila, c)
            celda.border = caja
            if campo == "med_id":
                celda.value = med_id
            elif campo == "lote_id":
                celda.value = lote_id
            elif campo == "diferencia":
                if f["fisico"] is not None and lote_id:
                    celda.value = f["fisico"] - (f["kardex"] or 0)
            else:
                v = valores[campo]
                if campo == "vencimiento" and v:
                    celda.value = date.fromisoformat(v)
                    celda.number_format = "DD/MM/YYYY"
                elif campo in CAMPOS_NUMERO and v != "":
                    celda.value = int(v)
                else:
                    celda.value = v or None
                ident = med_id if nivel == "med" else lote_id
                if ident:
                    foto[_clave(nivel, ident, campo)] = v
                    marcas = pendientes.get((nivel, ident, campo))
                    if marcas:
                        celda.fill = PatternFill("solid", fgColor=MORADO)
                        texto = "\n".join(
                            f"{datetime.strptime(m['cuando'], '%Y-%m-%d %H:%M:%S'):%d/%m/%Y %H:%M} "
                            f"{m['usuario'] or '¿?'}: {mostrar(campo, m['antes'])} → {mostrar(campo, m['despues'])}"
                            for m in marcas[:5])
                        celda.comment = Comment(f"Cambiado desde Excel\n{texto}", MARCA, width=320, height=120)
            if not editable:
                celda.fill = PatternFill("solid", fgColor=GRIS)
        fila += 1
    anchos = [30, 38, 12, 10, 14, 14, 16, 9, 11, 8, 8]
    for c, ancho in enumerate(anchos, 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = ancho
    ws.column_dimensions["J"].hidden = True
    ws.column_dimensions["K"].hidden = True
    ws.freeze_panes = ws.cell(fila_enc + 1, 2)
    ws.auto_filter.ref = f"A{fila_enc}:I{max(fila - 1, fila_enc)}"
    numeros = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                             showErrorMessage=True, error="Escribe un número entero (0 o más).")
    ws.add_data_validation(numeros)
    for letra in "CDH":
        numeros.add(f"{letra}{fila_enc + 1}:{letra}5000")

    h = wb.create_sheet("Cambios")
    h.append(["Cuándo", "Quién", "Desde", "Medicamento", "Lote", "Qué", "Antes", "Después", "Revisado"])
    for c in h[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
    for c in inv.historial_cambios():
        h.append([datetime.strptime(c["cuando"], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M"),
                  c["usuario"], c["origen"].capitalize(), c["nombre"] or "", c["nombre_lote"] or "",
                  NOMBRES_CAMPO.get(c["campo"], c["campo"]), mostrar(c["campo"], c["antes"]),
                  mostrar(c["campo"], c["despues"]), "Sí" if c["revisado"] else "No"])
    for letra, ancho in zip("ABCDEFGHI", [17, 16, 9, 28, 12, 20, 18, 18, 10]):
        h.column_dimensions[letra].width = ancho

    wb.properties.lastModifiedBy = MARCA
    wb.properties.creator = MARCA
    ruta.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=ruta.parent, prefix=".guardando-")
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, ruta)
    except PermissionError:
        Path(tmp).unlink(missing_ok=True)
        return False
    inv.guardar_foto_excel(foto, reemplazar=True)
    inv.poner_meta("excel_mtime", ruta.stat().st_mtime_ns)
    return True


# ── Importar ────────────────────────────────────────────────────────────────

def hay_cambios(inv: Inventario, ruta: Path) -> bool:
    try:
        return ruta.exists() and ruta.stat().st_mtime_ns != int(inv.meta("excel_mtime", "0"))
    except (OSError, ValueError):
        return False


def importar(inv: Inventario, ruta: Path) -> Resultado:
    res = Resultado()
    if not hay_cambios(inv, ruta):
        return res
    mtime = ruta.stat().st_mtime_ns
    try:
        wb = load_workbook(ruta, data_only=True)
    except Exception as e:  # archivo a medio guardar, dañado, etc.
        res.errores.append(f"No se pudo leer el Excel: {e}")
        return res
    usuario = (wb.properties.lastModifiedBy or "").strip()
    if not usuario or usuario == MARCA:
        usuario = getpass.getuser()
    res.usuario = usuario
    if "Inventario" not in wb.sheetnames:
        res.errores.append("El Excel no tiene la hoja «Inventario».")
        inv.poner_meta("excel_mtime", mtime)
        return res
    ws = wb["Inventario"]

    # Ubicar encabezados (por nombre, por si movieron columnas).
    fila_enc, cols = None, {}
    for r in range(1, 15):
        if str(ws.cell(r, 1).value or "").strip() == "Medicamento":
            fila_enc = r
            break
    if fila_enc is None:
        res.errores.append("No encontré la fila de encabezados («Medicamento») en la hoja Inventario.")
        inv.poner_meta("excel_mtime", mtime)
        return res
    por_encabezado = {enc: campo for enc, campo, _, _ in COLUMNAS}
    for c in range(1, ws.max_column + 1):
        enc = str(ws.cell(fila_enc, c).value or "").strip()
        if enc in por_encabezado:
            cols[por_encabezado[enc]] = c

    foto = inv.foto_excel()
    meds, lotes = _valores_actuales(inv)
    hoy = fecha_efectiva().isoformat()
    nueva_foto: dict[str, str] = {}
    cambios_med: dict[int, dict[str, tuple[str, str]]] = {}  # med_id -> campo -> (antes, después)
    nuevas: list[tuple[int, dict]] = []

    def entero(v):
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return None

    for r in range(fila_enc + 1, ws.max_row + 1):
        celdas = {campo: ws.cell(r, c).value for campo, c in cols.items()}
        if all(v in (None, "") for k, v in celdas.items() if k not in ("diferencia",)):
            continue
        datos, malos = {}, []
        for campo, v in celdas.items():
            if campo in ("diferencia", "med_id", "lote_id"):
                continue
            try:
                datos[campo] = normalizar(campo, v)
            except ValueError as e:
                malos.append(f"fila {r}, {NOMBRES_CAMPO.get(campo, campo)}: {e}")
        med_id, lote_id = entero(celdas.get("med_id")), entero(celdas.get("lote_id"))
        if med_id and med_id not in meds:
            res.errores.append(f"Fila {r}: el medicamento de esa fila ya no existe; se ignoró.")
            continue
        if lote_id and lote_id not in lotes:
            res.errores.append(f"Fila {r}: el lote de esa fila ya no existe; se ignoró.")
            continue
        if malos:
            res.errores.extend(malos)
            continue
        if not med_id or not lote_id:
            # Fila nueva, o medicamento sin lote al que le escribieron un lote / cantidades.
            nuevas.append((r, {**datos, "med_id": med_id}))
            if not med_id:
                continue
        # Campos del medicamento.
        for campo in ("nombre", "descripcion", "presentacion", "minimo"):
            if campo not in datos:
                continue
            clave = _clave("med", med_id, campo)
            antes = foto.get(clave)
            if antes is None or datos[campo] == antes:
                continue
            if campo == "nombre" and not datos[campo]:
                res.errores.append(f"Fila {r}: el nombre del medicamento no puede quedar vacío.")
                continue
            actual = meds[med_id][campo]
            if actual == datos[campo]:
                nueva_foto[clave] = datos[campo]
            elif actual != antes:
                res.choques.append(f"{meds[med_id]['nombre']} · {NOMBRES_CAMPO[campo]}: en el Excel "
                                   f"{mostrar(campo, datos[campo])}, pero en el programa ya cambió a "
                                   f"{mostrar(campo, actual)}. Se dejó lo del programa.")
                nueva_foto[clave] = datos[campo]
            else:
                cambios_med.setdefault(med_id, {}).setdefault(campo, (antes, datos[campo]))
        if not lote_id:
            continue
        # Campos del lote.
        actual_l = lotes[lote_id]
        nombre_med = meds[med_id]["nombre"]
        cambio_lote = {}
        for campo in ("lote", "vencimiento", "kardex", "fisico"):
            if campo not in datos:
                continue
            clave = _clave("lote", lote_id, campo)
            antes = foto.get(clave)
            if antes is None or datos[campo] == antes:
                continue
            actual = actual_l[campo]
            if actual == datos[campo]:
                nueva_foto[clave] = datos[campo]
            elif actual != antes:
                res.choques.append(f"{nombre_med} lote {actual_l['lote']} · {NOMBRES_CAMPO[campo]}: en el Excel "
                                   f"{mostrar(campo, datos[campo])}, pero en el programa ya está en "
                                   f"{mostrar(campo, actual)}. Se dejó lo del programa.")
                nueva_foto[clave] = datos[campo]
            elif campo in ("fisico", "kardex") and datos[campo] == "":
                res.errores.append(f"{nombre_med} lote {actual_l['lote']}: no se puede dejar vacío "
                                   f"«{NOMBRES_CAMPO[campo]}»; se ignoró.")
                nueva_foto[clave] = datos[campo]
            else:
                cambio_lote[campo] = (antes, datos[campo])
        if "lote" in cambio_lote or "vencimiento" in cambio_lote:
            nuevo_lote = cambio_lote.get("lote", (None, actual_l["lote"]))[1]
            nuevo_vence = cambio_lote.get("vencimiento", (None, actual_l["vencimiento"]))[1]
            try:
                inv.editar_lote(lote_id, nuevo_lote, nuevo_vence or None)
                for campo in ("lote", "vencimiento"):
                    if campo in cambio_lote:
                        a, d = cambio_lote[campo]
                        inv.anotar_cambio("EXCEL", usuario, campo, a, d, med_id, lote_id)
                        nueva_foto[_clave("lote", lote_id, campo)] = d
                        res.aplicados.append(f"{nombre_med} · {NOMBRES_CAMPO[campo]}: "
                                             f"{mostrar(campo, a)} → {mostrar(campo, d)}")
            except ErrorInventario as e:
                res.errores.append(f"{nombre_med}: {e}")
        if "kardex" in cambio_lote:
            a, d = cambio_lote["kardex"]
            try:
                dif = inv.ajustar_kardex(lote_id, int(d), hoy, f"Modificado en Excel por {usuario}")
                inv.anotar_cambio("EXCEL", usuario, "kardex", a, d, med_id, lote_id)
                nueva_foto[_clave("lote", lote_id, "kardex")] = d
                res.aplicados.append(f"{nombre_med} lote {actual_l['lote']} · Kardex: {a} → {d} "
                                     f"(ajuste {dif:+d} el {texto_fecha(hoy)})")
            except ErrorInventario as e:
                res.errores.append(f"{nombre_med} lote {actual_l['lote']} · Kardex: {e}")
        if "fisico" in cambio_lote:
            a, d = cambio_lote["fisico"]
            try:
                inv.guardar_cuadre(hoy, "SEMANAL", {lote_id: int(d)}, ajustar=False,
                                   nota=f"Conteo escrito en Excel por {usuario}")
                inv.anotar_cambio("EXCEL", usuario, "fisico", a, d, med_id, lote_id)
                nueva_foto[_clave("lote", lote_id, "fisico")] = d
                res.aplicados.append(f"{nombre_med} lote {actual_l['lote']} · Existencias físicas: "
                                     f"{mostrar('fisico', a)} → {d}")
            except ErrorInventario as e:
                res.errores.append(f"{nombre_med} lote {actual_l['lote']} · Existencias físicas: {e}")

    # Aplicar cambios del medicamento (juntos, una sola vez por medicamento).
    for med_id, campos in cambios_med.items():
        actual = meds[med_id]
        nuevo = {k: campos.get(k, (None, actual[k]))[1] for k in ("nombre", "descripcion", "presentacion", "minimo")}
        try:
            inv.guardar_medicamento(nuevo["nombre"], nuevo["descripcion"], nuevo["presentacion"],
                                    int(nuevo["minimo"] or 0), med_id)
        except ErrorInventario as e:
            res.errores.append(f"{actual['nombre']}: {e}")
            continue
        for campo, (a, d) in campos.items():
            inv.anotar_cambio("EXCEL", usuario, campo, a, d, med_id, None)
            nueva_foto[_clave("med", med_id, campo)] = d
            res.aplicados.append(f"{nuevo['nombre']} · {NOMBRES_CAMPO[campo]}: "
                                 f"{mostrar(campo, a)} → {mostrar(campo, d)}")

    # Filas nuevas.
    for r, d in nuevas:
        nombre = d.get("nombre", "")
        if not nombre:
            res.errores.append(f"Fila {r}: falta el nombre del medicamento; se ignoró.")
            continue
        med = inv.medicamento(d["med_id"]) if d.get("med_id") else inv.medicamento_por_nombre(nombre)
        lote_txt = d.get("lote", "")
        kardex = int(d["kardex"]) if d.get("kardex") else 0
        fisico = d.get("fisico", "")
        if med and not lote_txt and not kardex and fisico == "":
            continue  # ya existe y no trae nada nuevo
        try:
            if not med:
                med_id = inv.guardar_medicamento(nombre, d.get("descripcion", ""), d.get("presentacion", ""),
                                                 int(d.get("minimo") or 0))
                inv.anotar_cambio("EXCEL", usuario, "nuevo", "", nombre, med_id, None)
                res.aplicados.append(f"Nuevo medicamento: {nombre}")
            else:
                med_id = med["id"]
            if not (lote_txt or kardex or fisico != ""):
                continue
            existentes = {l["lote"].upper() for l in inv.lotes(med_id)}
            lote_nombre = (lote_txt or "SIN LOTE").upper()
            if lote_nombre in existentes and med:
                continue  # fila ya importada antes
            lote_id = inv.obtener_lote(med_id, lote_nombre, d.get("vencimiento") or None)
            inv.anotar_cambio("EXCEL", usuario, "lote", "", lote_nombre, med_id, lote_id)
            texto = f"{inv.medicamento(med_id)['nombre']}: lote nuevo {lote_nombre}"
            if kardex:
                inv.sumar_dia(lote_id, hoy, "ENTRADA", kardex, f"Agregado en Excel por {usuario}")
                inv.anotar_cambio("EXCEL", usuario, "kardex", "0", str(kardex), med_id, lote_id)
                texto += f" con entrada de {kardex}"
            if fisico != "":
                inv.guardar_cuadre(hoy, "SEMANAL", {lote_id: int(fisico)}, ajustar=False,
                                   nota=f"Conteo escrito en Excel por {usuario}")
                inv.anotar_cambio("EXCEL", usuario, "fisico", "", fisico, med_id, lote_id)
            res.aplicados.append(texto)
        except ErrorInventario as e:
            res.errores.append(f"Fila {r} ({nombre}): {e}")

    inv.guardar_foto_excel(nueva_foto, reemplazar=False)
    inv.poner_meta("excel_mtime", mtime)
    return res
