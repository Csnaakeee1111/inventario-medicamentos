"""Base de datos del inventario (SQLite).

Un solo archivo `inventario.db` en la carpeta de datos del usuario, separado del
programa: al actualizar el .exe los datos no se tocan.

Modelo:
- medicamentos: catálogo, se escribe una sola vez.
- lotes: cada medicamento puede tener varios lotes con su vencimiento.
- movimientos: ENTRADA / SALIDA (una fila por día, lote y tipo) y AJUSTE (de cuadres).
  El kardex y las existencias del sistema salen de aquí.
- cuadres + cuadre_detalle: conteos físicos semanales o mensuales.
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

SIN_LOTE = "SIN LOTE"


def carpeta_datos() -> Path:
    """Carpeta donde viven la base, los respaldos y las exportaciones."""
    if os.environ.get("INVENTARIO_DATOS"):
        base = Path(os.environ["INVENTARIO_DATOS"])
    elif sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home())) / "InventarioMedicamentos"
    else:
        base = Path.home() / ".local" / "share" / "InventarioMedicamentos"
    base.mkdir(parents=True, exist_ok=True)
    return base


ESQUEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS meta (
    clave TEXT PRIMARY KEY,
    valor TEXT
);

CREATE TABLE IF NOT EXISTS medicamentos (
    id           INTEGER PRIMARY KEY,
    nombre       TEXT NOT NULL UNIQUE COLLATE NOCASE,
    descripcion  TEXT NOT NULL DEFAULT '',
    presentacion TEXT NOT NULL DEFAULT '',
    minimo       INTEGER NOT NULL DEFAULT 0,
    activo       INTEGER NOT NULL DEFAULT 1,
    creado       TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS lotes (
    id             INTEGER PRIMARY KEY,
    medicamento_id INTEGER NOT NULL REFERENCES medicamentos(id) ON DELETE CASCADE,
    lote           TEXT NOT NULL COLLATE NOCASE,
    vencimiento    TEXT,
    UNIQUE (medicamento_id, lote)
);

CREATE TABLE IF NOT EXISTS movimientos (
    id       INTEGER PRIMARY KEY,
    fecha    TEXT NOT NULL,
    lote_id  INTEGER NOT NULL REFERENCES lotes(id) ON DELETE CASCADE,
    tipo     TEXT NOT NULL CHECK (tipo IN ('ENTRADA', 'SALIDA', 'AJUSTE')),
    cantidad INTEGER NOT NULL,
    nota     TEXT NOT NULL DEFAULT '',
    creado   TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE UNIQUE INDEX IF NOT EXISTS mov_dia
    ON movimientos (fecha, lote_id, tipo) WHERE tipo IN ('ENTRADA', 'SALIDA');
CREATE INDEX IF NOT EXISTS mov_lote ON movimientos (lote_id, fecha);

CREATE TABLE IF NOT EXISTS cuadres (
    id     INTEGER PRIMARY KEY,
    fecha  TEXT NOT NULL,
    tipo   TEXT NOT NULL CHECK (tipo IN ('SEMANAL', 'MENSUAL')),
    nota   TEXT NOT NULL DEFAULT '',
    ajustado INTEGER NOT NULL DEFAULT 0,
    creado TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS cambios (
    id             INTEGER PRIMARY KEY,
    cuando         TEXT NOT NULL,
    origen         TEXT NOT NULL,
    usuario        TEXT NOT NULL DEFAULT '',
    medicamento_id INTEGER REFERENCES medicamentos(id) ON DELETE CASCADE,
    lote_id        INTEGER REFERENCES lotes(id) ON DELETE CASCADE,
    campo          TEXT NOT NULL,
    antes          TEXT NOT NULL DEFAULT '',
    despues        TEXT NOT NULL DEFAULT '',
    revisado       INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS cambios_pend ON cambios (revisado, medicamento_id, lote_id);

-- Lo que se escribió en Inventario_actual.xlsx la última vez, para saber qué cambió la gente.
CREATE TABLE IF NOT EXISTS excel_foto (
    clave TEXT PRIMARY KEY,
    valor TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS cuadre_detalle (
    cuadre_id INTEGER NOT NULL REFERENCES cuadres(id) ON DELETE CASCADE,
    lote_id   INTEGER NOT NULL REFERENCES lotes(id) ON DELETE CASCADE,
    sistema   INTEGER NOT NULL,
    fisico    INTEGER NOT NULL,
    PRIMARY KEY (cuadre_id, lote_id)
);
"""

VERSION_ESQUEMA = 1


# ── Fechas ──────────────────────────────────────────────────────────────────

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
DIAS_CORTOS = ["Lun", "Mar", "Mié", "Jue", "Vie"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def lunes_de(d: date) -> date:
    return d - timedelta(days=d.weekday())


def dias_habiles(lunes: date) -> list[date]:
    return [lunes + timedelta(days=i) for i in range(5)]


def es_habil(d: date) -> bool:
    return d.weekday() < 5


def ultimo_habil_del_mes(d: date) -> date:
    siguiente = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    ultimo = siguiente - timedelta(days=1)
    while not es_habil(ultimo):
        ultimo -= timedelta(days=1)
    return ultimo


def texto_fecha(iso: str | None) -> str:
    """'2026-11-15' -> '15/11/2026'."""
    if not iso:
        return ""
    try:
        return datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return iso


# ── Errores para mostrar al usuario ─────────────────────────────────────────

class ErrorInventario(Exception):
    """Error con mensaje listo para enseñar en pantalla."""


@dataclass
class FilaKardex:
    fecha: str
    lote: str
    concepto: str
    entrada: int
    salida: int
    saldo: int
    nota: str


class Inventario:
    def __init__(self, ruta: Path | None = None):
        self.ruta = ruta or carpeta_datos() / "inventario.db"
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        self.cx = sqlite3.connect(self.ruta)
        self.cx.row_factory = sqlite3.Row
        self.cx.execute("PRAGMA foreign_keys = ON")
        self.cx.execute("PRAGMA journal_mode = WAL")
        self.cx.executescript(ESQUEMA)
        self.cx.execute(
            "INSERT OR IGNORE INTO meta VALUES ('version_esquema', ?)", (VERSION_ESQUEMA,))
        self.cx.commit()

    def cerrar(self) -> None:
        self.cx.close()

    # ── Respaldo ────────────────────────────────────────────────────────────

    def respaldar(self, guardar: int = 30) -> Path:
        """Copia diaria de la base en respaldos/; deja las últimas `guardar`."""
        carpeta = self.ruta.parent / "respaldos"
        carpeta.mkdir(exist_ok=True)
        destino = carpeta / f"inventario-{date.today().isoformat()}.db"
        copia = sqlite3.connect(destino)
        with copia:
            self.cx.backup(copia)
        copia.close()
        viejos = sorted(carpeta.glob("inventario-*.db"))[:-guardar]
        for v in viejos:
            v.unlink(missing_ok=True)
        return destino

    def restaurar(self, origen: Path) -> None:
        self.cx.close()
        shutil.copy2(origen, self.ruta)
        for extra in ("-wal", "-shm"):
            Path(str(self.ruta) + extra).unlink(missing_ok=True)
        self.__init__(self.ruta)

    # ── Medicamentos ────────────────────────────────────────────────────────

    def medicamentos(self, buscar: str = "", incluir_inactivos: bool = False) -> list[sqlite3.Row]:
        sql = """
        SELECT m.*,
            COALESCE((SELECT SUM(CASE mv.tipo WHEN 'SALIDA' THEN -mv.cantidad ELSE mv.cantidad END)
                      FROM movimientos mv JOIN lotes l ON l.id = mv.lote_id
                      WHERE l.medicamento_id = m.id), 0) AS existencia,
            (SELECT MIN(l.vencimiento) FROM lotes l
              WHERE l.medicamento_id = m.id AND l.vencimiento IS NOT NULL
                AND (SELECT COALESCE(SUM(CASE mv.tipo WHEN 'SALIDA' THEN -mv.cantidad ELSE mv.cantidad END), 0)
                     FROM movimientos mv WHERE mv.lote_id = l.id) > 0) AS proximo_vence,
            (SELECT COUNT(*) FROM lotes l WHERE l.medicamento_id = m.id) AS num_lotes
        FROM medicamentos m
        WHERE (m.activo = 1 OR ?)
          AND (m.nombre LIKE ? OR m.descripcion LIKE ?)
        ORDER BY m.nombre COLLATE NOCASE
        """
        patron = f"%{buscar.strip()}%"
        return self.cx.execute(sql, (int(incluir_inactivos), patron, patron)).fetchall()

    def ultimo_conteo(self, medicamento_id: int) -> tuple[int, str] | None:
        """Último conteo físico (suma de lotes del cuadre más reciente que lo incluye)."""
        fila = self.cx.execute("""
            SELECT c.id, c.fecha, SUM(d.fisico) AS fisico
            FROM cuadres c JOIN cuadre_detalle d ON d.cuadre_id = c.id
            JOIN lotes l ON l.id = d.lote_id
            WHERE l.medicamento_id = ?
            GROUP BY c.id ORDER BY c.fecha DESC, c.id DESC LIMIT 1
        """, (medicamento_id,)).fetchone()
        return (fila["fisico"], fila["fecha"]) if fila else None

    def medicamento(self, medicamento_id: int) -> sqlite3.Row:
        return self.cx.execute("SELECT * FROM medicamentos WHERE id = ?", (medicamento_id,)).fetchone()

    def medicamento_por_nombre(self, nombre: str) -> sqlite3.Row | None:
        return self.cx.execute(
            "SELECT * FROM medicamentos WHERE nombre = ? COLLATE NOCASE", (nombre.strip(),)).fetchone()

    def nombres_medicamentos(self) -> list[str]:
        return [r[0] for r in self.cx.execute(
            "SELECT nombre FROM medicamentos WHERE activo = 1 ORDER BY nombre COLLATE NOCASE")]

    def guardar_medicamento(self, nombre: str, descripcion: str = "", presentacion: str = "",
                            minimo: int = 0, medicamento_id: int | None = None) -> int:
        nombre = " ".join(nombre.split())
        if not nombre:
            raise ErrorInventario("El nombre del medicamento no puede ir vacío.")
        otro = self.medicamento_por_nombre(nombre)
        if otro and otro["id"] != medicamento_id:
            if not otro["activo"]:
                raise ErrorInventario(
                    f"«{otro['nombre']}» ya existe pero está desactivado. "
                    "Actívalo desde el catálogo (casilla «Ver desactivados»).")
            raise ErrorInventario(f"«{otro['nombre']}» ya está en el catálogo.")
        with self.cx:
            if medicamento_id is None:
                cur = self.cx.execute(
                    "INSERT INTO medicamentos (nombre, descripcion, presentacion, minimo) VALUES (?,?,?,?)",
                    (nombre, descripcion.strip(), presentacion.strip(), int(minimo)))
                return cur.lastrowid
            self.cx.execute(
                "UPDATE medicamentos SET nombre=?, descripcion=?, presentacion=?, minimo=? WHERE id=?",
                (nombre, descripcion.strip(), presentacion.strip(), int(minimo), medicamento_id))
            return medicamento_id

    def activar_medicamento(self, medicamento_id: int, activo: bool) -> None:
        with self.cx:
            self.cx.execute("UPDATE medicamentos SET activo=? WHERE id=?", (int(activo), medicamento_id))

    # ── Lotes ───────────────────────────────────────────────────────────────

    def lotes(self, medicamento_id: int) -> list[sqlite3.Row]:
        return self.cx.execute("""
            SELECT l.*, COALESCE(SUM(CASE mv.tipo WHEN 'SALIDA' THEN -mv.cantidad ELSE mv.cantidad END), 0)
                   AS existencia
            FROM lotes l LEFT JOIN movimientos mv ON mv.lote_id = l.id
            WHERE l.medicamento_id = ?
            GROUP BY l.id ORDER BY l.vencimiento IS NULL, l.vencimiento, l.lote
        """, (medicamento_id,)).fetchall()

    def obtener_lote(self, medicamento_id: int, lote: str, vencimiento: str | None) -> int:
        """Devuelve el id del lote; lo crea si no existe. Actualiza el vencimiento si viene."""
        lote = " ".join((lote or "").split()).upper() or SIN_LOTE
        fila = self.cx.execute(
            "SELECT id, vencimiento FROM lotes WHERE medicamento_id=? AND lote=? COLLATE NOCASE",
            (medicamento_id, lote)).fetchone()
        with self.cx:
            if fila:
                if vencimiento and fila["vencimiento"] != vencimiento:
                    self.cx.execute("UPDATE lotes SET vencimiento=? WHERE id=?", (vencimiento, fila["id"]))
                return fila["id"]
            return self.cx.execute(
                "INSERT INTO lotes (medicamento_id, lote, vencimiento) VALUES (?,?,?)",
                (medicamento_id, lote, vencimiento)).lastrowid

    def editar_lote(self, lote_id: int, lote: str, vencimiento: str | None) -> None:
        lote = " ".join(lote.split()).upper() or SIN_LOTE
        try:
            with self.cx:
                self.cx.execute("UPDATE lotes SET lote=?, vencimiento=? WHERE id=?",
                                (lote, vencimiento, lote_id))
        except sqlite3.IntegrityError:
            raise ErrorInventario(f"Ya existe otro lote «{lote}» para este medicamento.") from None

    # ── Existencias ─────────────────────────────────────────────────────────

    def saldo_lote(self, lote_id: int, hasta: str | None = None, antes_de: str | None = None) -> int:
        sql = """SELECT COALESCE(SUM(CASE tipo WHEN 'SALIDA' THEN -cantidad ELSE cantidad END), 0)
                 FROM movimientos WHERE lote_id = ?"""
        args: list = [lote_id]
        if hasta:
            sql += " AND fecha <= ?"
            args.append(hasta)
        if antes_de:
            sql += " AND fecha < ?"
            args.append(antes_de)
        return self.cx.execute(sql, args).fetchone()[0]

    # ── Movimientos diarios (tabla de la semana) ────────────────────────────

    def cantidad_dia(self, lote_id: int, fecha: str, tipo: str) -> int:
        fila = self.cx.execute(
            "SELECT cantidad FROM movimientos WHERE lote_id=? AND fecha=? AND tipo=?",
            (lote_id, fecha, tipo)).fetchone()
        return fila[0] if fila else 0

    def fijar_dia(self, lote_id: int, fecha: str, tipo: str, cantidad: int, nota: str = "") -> None:
        """Deja el total de ENTRADA o SALIDA de ese lote en ese día en `cantidad`."""
        if tipo not in ("ENTRADA", "SALIDA"):
            raise ValueError(tipo)
        if cantidad < 0:
            raise ErrorInventario("La cantidad no puede ser negativa.")
        d = date.fromisoformat(fecha)
        if not es_habil(d):
            raise ErrorInventario("Solo se registran movimientos de lunes a viernes.")
        anterior = self.cantidad_dia(lote_id, fecha, tipo)
        signo = 1 if tipo == "ENTRADA" else -1
        cambio = signo * (cantidad - anterior)
        if cambio < 0:
            self._revisar_que_alcance(lote_id, fecha, cambio)
        with self.cx:
            if cantidad == 0:
                self.cx.execute("DELETE FROM movimientos WHERE lote_id=? AND fecha=? AND tipo=?",
                                (lote_id, fecha, tipo))
            else:
                self.cx.execute("""
                    INSERT INTO movimientos (fecha, lote_id, tipo, cantidad, nota) VALUES (?,?,?,?,?)
                    ON CONFLICT (fecha, lote_id, tipo) WHERE tipo IN ('ENTRADA','SALIDA')
                    DO UPDATE SET cantidad=excluded.cantidad,
                                  nota=CASE WHEN excluded.nota <> '' THEN excluded.nota ELSE nota END
                """, (fecha, lote_id, tipo, cantidad, nota.strip()))

    def sumar_dia(self, lote_id: int, fecha: str, tipo: str, cantidad: int, nota: str = "") -> None:
        if cantidad <= 0:
            raise ErrorInventario("La cantidad debe ser mayor que cero.")
        self.fijar_dia(lote_id, fecha, tipo, self.cantidad_dia(lote_id, fecha, tipo) + cantidad, nota)

    def _revisar_que_alcance(self, lote_id: int, fecha: str, cambio: int) -> None:
        """El saldo no puede quedar negativo ni ese día ni después."""
        fechas = [fecha] + [r[0] for r in self.cx.execute(
            "SELECT DISTINCT fecha FROM movimientos WHERE lote_id=? AND fecha>? ORDER BY fecha",
            (lote_id, fecha))]
        for f in fechas:
            saldo = self.saldo_lote(lote_id, hasta=f) + cambio
            if saldo < 0:
                disponible = self.saldo_lote(lote_id, hasta=f)
                cuando = "ese día" if f == fecha else f"el {texto_fecha(f)}"
                raise ErrorInventario(
                    f"No alcanza: {cuando} el lote quedaría en {saldo}. "
                    f"Disponible {cuando}: {disponible}.")

    def filas_semana(self, lunes: date, solo_con_existencia: bool = True) -> list[dict]:
        """Una fila por lote para la tabla de lunes a viernes."""
        dias = [d.isoformat() for d in dias_habiles(lunes)]
        inicio, fin = dias[0], dias[-1]
        lotes = self.cx.execute("""
            SELECT l.id, l.lote, l.vencimiento, m.id AS med_id, m.nombre, m.descripcion
            FROM lotes l JOIN medicamentos m ON m.id = l.medicamento_id
            WHERE m.activo = 1
            ORDER BY m.nombre COLLATE NOCASE, l.vencimiento IS NULL, l.vencimiento, l.lote
        """).fetchall()
        movs: dict[tuple[int, str, str], int] = {}
        for r in self.cx.execute(
                "SELECT lote_id, fecha, tipo, SUM(cantidad) FROM movimientos "
                "WHERE fecha BETWEEN ? AND ? GROUP BY lote_id, fecha, tipo", (inicio, fin)):
            movs[(r[0], r[1], r[2])] = r[3]
        filas = []
        for l in lotes:
            inicial = self.saldo_lote(l["id"], antes_de=inicio)
            dia_datos = []
            hay_mov = False
            for f in dias:
                e = movs.get((l["id"], f, "ENTRADA"), 0)
                s = movs.get((l["id"], f, "SALIDA"), 0)
                a = movs.get((l["id"], f, "AJUSTE"), 0)
                hay_mov = hay_mov or bool(e or s or a)
                dia_datos.append({"fecha": f, "entrada": e, "salida": s, "ajuste": a})
            final = inicial + sum(d["entrada"] - d["salida"] + d["ajuste"] for d in dia_datos)
            if solo_con_existencia and not (inicial or hay_mov or final):
                continue
            filas.append({
                "lote_id": l["id"], "medicamento_id": l["med_id"], "nombre": l["nombre"],
                "descripcion": l["descripcion"], "lote": l["lote"], "vencimiento": l["vencimiento"],
                "inicial": inicial, "dias": dia_datos, "final": final,
            })
        return filas

    # ── Kardex ──────────────────────────────────────────────────────────────

    def kardex(self, medicamento_id: int, lote_id: int | None = None,
               desde: str | None = None, hasta: str | None = None) -> list[FilaKardex]:
        args: list = [medicamento_id]
        filtro_lote = ""
        if lote_id:
            filtro_lote = " AND l.id = ?"
            args.append(lote_id)
        saldo = 0
        if desde:
            saldo = self.cx.execute(f"""
                SELECT COALESCE(SUM(CASE mv.tipo WHEN 'SALIDA' THEN -mv.cantidad ELSE mv.cantidad END), 0)
                FROM movimientos mv JOIN lotes l ON l.id = mv.lote_id
                WHERE l.medicamento_id = ? {filtro_lote} AND mv.fecha < ?
            """, args + [desde]).fetchone()[0]
        sql = f"""
            SELECT mv.fecha, l.lote, mv.tipo, mv.cantidad, mv.nota
            FROM movimientos mv JOIN lotes l ON l.id = mv.lote_id
            WHERE l.medicamento_id = ? {filtro_lote}
        """
        filtro_args = list(args)
        if desde:
            sql += " AND mv.fecha >= ?"
            filtro_args.append(desde)
        if hasta:
            sql += " AND mv.fecha <= ?"
            filtro_args.append(hasta)
        sql += " ORDER BY mv.fecha, CASE mv.tipo WHEN 'ENTRADA' THEN 0 WHEN 'SALIDA' THEN 1 ELSE 2 END, mv.id"
        filas = []
        if desde:
            filas.append(FilaKardex(desde, "", "Saldo anterior", 0, 0, saldo, ""))
        for r in self.cx.execute(sql, filtro_args):
            cant = r["cantidad"]
            if r["tipo"] == "ENTRADA":
                entrada, salida, concepto = cant, 0, "Entrada"
            elif r["tipo"] == "SALIDA":
                entrada, salida, concepto = 0, cant, "Salida"
            else:
                entrada, salida = (cant, 0) if cant >= 0 else (0, -cant)
                concepto = "Ajuste desde Excel" if r["nota"].startswith("Modificado en Excel") else "Ajuste por cuadre"
            saldo += entrada - salida
            filas.append(FilaKardex(r["fecha"], r["lote"], concepto, entrada, salida, saldo, r["nota"]))
        return filas

    # ── Cuadres (conteo físico) ─────────────────────────────────────────────

    def filas_para_cuadre(self, fecha: str) -> list[dict]:
        """Lotes con existencia en el sistema a esa fecha (o con movimientos)."""
        filas = []
        for l in self.cx.execute("""
            SELECT l.id, l.lote, l.vencimiento, m.nombre
            FROM lotes l JOIN medicamentos m ON m.id = l.medicamento_id
            WHERE m.activo = 1 AND EXISTS (SELECT 1 FROM movimientos mv WHERE mv.lote_id = l.id AND mv.fecha <= ?)
            ORDER BY m.nombre COLLATE NOCASE, l.vencimiento IS NULL, l.vencimiento, l.lote
        """, (fecha,)):
            sistema = self.saldo_lote(l["id"], hasta=fecha)
            if sistema == 0:
                continue
            filas.append({"lote_id": l["id"], "nombre": l["nombre"], "lote": l["lote"],
                          "vencimiento": l["vencimiento"], "sistema": sistema})
        return filas

    def guardar_cuadre(self, fecha: str, tipo: str, conteos: dict[int, int],
                       ajustar: bool, nota: str = "") -> int:
        """`conteos` = {lote_id: cantidad física}. Si `ajustar`, crea AJUSTES por la diferencia."""
        if not conteos:
            raise ErrorInventario("No hay ningún conteo físico escrito.")
        if not es_habil(date.fromisoformat(fecha)):
            raise ErrorInventario("El cuadre debe tener fecha de lunes a viernes.")
        if any(v < 0 for v in conteos.values()):
            raise ErrorInventario("El conteo físico no puede ser negativo.")
        if ajustar:
            # Un ajuste hacia abajo en una fecha pasada no puede dejar sin saldo a salidas posteriores.
            for lote_id, fisico in conteos.items():
                cambio = fisico - self.saldo_lote(lote_id, hasta=fecha)
                if cambio < 0:
                    try:
                        self._revisar_que_alcance(lote_id, fecha, cambio)
                    except ErrorInventario as e:
                        lote = self.cx.execute(
                            "SELECT m.nombre, l.lote FROM lotes l JOIN medicamentos m ON m.id = l.medicamento_id "
                            "WHERE l.id = ?", (lote_id,)).fetchone()
                        raise ErrorInventario(
                            f"{lote['nombre']} (lote {lote['lote']}): ajustar a {fisico} choca con salidas "
                            f"registradas después del {texto_fecha(fecha)}. {e} Revisa esas salidas o "
                            "guarda el cuadre sin ajustar.") from None
        with self.cx:
            cid = self.cx.execute(
                "INSERT INTO cuadres (fecha, tipo, nota, ajustado) VALUES (?,?,?,?)",
                (fecha, tipo, nota.strip(), int(ajustar))).lastrowid
            for lote_id, fisico in conteos.items():
                sistema = self.saldo_lote(lote_id, hasta=fecha)
                self.cx.execute(
                    "INSERT INTO cuadre_detalle (cuadre_id, lote_id, sistema, fisico) VALUES (?,?,?,?)",
                    (cid, lote_id, sistema, fisico))
                if ajustar and fisico != sistema:
                    self.cx.execute(
                        "INSERT INTO movimientos (fecha, lote_id, tipo, cantidad, nota) VALUES (?,?,?,?,?)",
                        (fecha, lote_id, "AJUSTE", fisico - sistema,
                         f"Cuadre {tipo.lower()} #{cid}" + (f": {nota.strip()}" if nota.strip() else "")))
        return cid

    def cuadres(self) -> list[sqlite3.Row]:
        return self.cx.execute("""
            SELECT c.*, COUNT(d.lote_id) AS lineas,
                   SUM(CASE WHEN d.fisico <> d.sistema THEN 1 ELSE 0 END) AS con_diferencia
            FROM cuadres c LEFT JOIN cuadre_detalle d ON d.cuadre_id = c.id
            GROUP BY c.id ORDER BY c.fecha DESC, c.id DESC
        """).fetchall()

    def detalle_cuadre(self, cuadre_id: int) -> list[dict]:
        return [dict(r) for r in self.cx.execute("""
            SELECT d.lote_id, m.nombre, l.lote, l.vencimiento, d.sistema, d.fisico
            FROM cuadre_detalle d JOIN lotes l ON l.id = d.lote_id
            JOIN medicamentos m ON m.id = l.medicamento_id
            WHERE d.cuadre_id = ? ORDER BY m.nombre COLLATE NOCASE, l.lote
        """, (cuadre_id,))]

    # ── Vencimientos ────────────────────────────────────────────────────────

    def vencimientos(self, dias: int) -> list[dict]:
        limite = (date.today() + timedelta(days=dias)).isoformat()
        filas = []
        for l in self.cx.execute("""
            SELECT l.id, l.lote, l.vencimiento, m.nombre
            FROM lotes l JOIN medicamentos m ON m.id = l.medicamento_id
            WHERE m.activo = 1 AND l.vencimiento IS NOT NULL AND l.vencimiento <= ?
            ORDER BY l.vencimiento, m.nombre COLLATE NOCASE
        """, (limite,)):
            existencia = self.saldo_lote(l["id"])
            if existencia > 0:
                restan = (date.fromisoformat(l["vencimiento"]) - date.today()).days
                filas.append({"nombre": l["nombre"], "lote": l["lote"], "vencimiento": l["vencimiento"],
                              "existencia": existencia, "dias": restan})
        return filas

    # ── Vista general por lote (tabla principal) ────────────────────────────

    def inventario_por_lote(self, buscar: str = "", incluir_inactivos: bool = False,
                            incluir_vacios: bool = False) -> list[dict]:
        """Medicamento, descripción, existencia física (último conteo), kardex, lote, vencimiento."""
        patron = f"%{buscar.strip()}%"
        filas = []
        for r in self.cx.execute("""
            SELECT m.id AS med_id, m.nombre, m.descripcion, m.presentacion, m.minimo, m.activo,
                   l.id AS lote_id, l.lote, l.vencimiento,
                   COALESCE((SELECT SUM(CASE mv.tipo WHEN 'SALIDA' THEN -mv.cantidad ELSE mv.cantidad END)
                             FROM movimientos mv WHERE mv.lote_id = l.id), 0) AS kardex,
                   (SELECT d.fisico FROM cuadre_detalle d JOIN cuadres c ON c.id = d.cuadre_id
                     WHERE d.lote_id = l.id ORDER BY c.fecha DESC, c.id DESC LIMIT 1) AS fisico,
                   (SELECT c.fecha FROM cuadre_detalle d JOIN cuadres c ON c.id = d.cuadre_id
                     WHERE d.lote_id = l.id ORDER BY c.fecha DESC, c.id DESC LIMIT 1) AS fecha_conteo
            FROM medicamentos m LEFT JOIN lotes l ON l.medicamento_id = m.id
            WHERE (m.activo = 1 OR ?) AND (m.nombre LIKE ? OR m.descripcion LIKE ? OR l.lote LIKE ?)
            ORDER BY m.nombre COLLATE NOCASE, l.vencimiento IS NULL, l.vencimiento, l.lote
        """, (int(incluir_inactivos), patron, patron, patron)):
            filas.append(dict(r))
        # Lotes en cero sin conteo pendiente se esconden, salvo que el medicamento no tenga otro.
        if not incluir_vacios:
            con_algo = {f["med_id"] for f in filas if f["kardex"] or f["fisico"]}
            vistos = set()
            limpias = []
            for f in filas:
                if f["kardex"] or f["fisico"]:
                    limpias.append(f)
                elif f["med_id"] not in con_algo and f["med_id"] not in vistos:
                    vistos.add(f["med_id"])
                    limpias.append({**f, "lote_id": None, "lote": None, "vencimiento": None,
                                    "kardex": 0, "fisico": None, "fecha_conteo": None})
            filas = limpias
        return filas

    def todos_los_movimientos(self) -> list[dict]:
        """Kardex de todos los medicamentos, con saldo corrido por medicamento."""
        filas, saldos = [], {}
        for r in self.cx.execute("""
            SELECT m.id AS med_id, m.nombre, mv.fecha, l.lote, mv.tipo, mv.cantidad, mv.nota
            FROM movimientos mv JOIN lotes l ON l.id = mv.lote_id
            JOIN medicamentos m ON m.id = l.medicamento_id
            ORDER BY m.nombre COLLATE NOCASE, mv.fecha,
                     CASE mv.tipo WHEN 'ENTRADA' THEN 0 WHEN 'SALIDA' THEN 1 ELSE 2 END, mv.id
        """):
            c = r["cantidad"]
            ent = c if r["tipo"] == "ENTRADA" or (r["tipo"] == "AJUSTE" and c > 0) else 0
            sal = c if r["tipo"] == "SALIDA" else (-c if r["tipo"] == "AJUSTE" and c < 0 else 0)
            saldos[r["med_id"]] = saldos.get(r["med_id"], 0) + ent - sal
            concepto = {"ENTRADA": "Entrada", "SALIDA": "Salida"}.get(r["tipo"]) or (
                "Ajuste desde Excel" if r["nota"].startswith("Modificado en Excel") else "Ajuste por cuadre")
            filas.append({"nombre": r["nombre"], "fecha": r["fecha"], "lote": r["lote"],
                          "concepto": concepto, "entrada": ent, "salida": sal,
                          "saldo": saldos[r["med_id"]], "nota": r["nota"]})
        return filas

    # ── Kardex ajustado a mano (desde el Excel) ─────────────────────────────

    def ajustar_kardex(self, lote_id: int, nuevo: int, fecha: str, nota: str) -> int:
        """Deja el kardex del lote en `nuevo` con un AJUSTE en `fecha`. Devuelve la diferencia."""
        if nuevo < 0:
            raise ErrorInventario("El kardex no puede quedar negativo.")
        if not es_habil(date.fromisoformat(fecha)):
            raise ErrorInventario("El ajuste debe tener fecha de lunes a viernes.")
        cambio = nuevo - self.saldo_lote(lote_id, hasta=fecha)
        if cambio == 0:
            return 0
        if cambio < 0:
            self._revisar_que_alcance(lote_id, fecha, cambio)
        with self.cx:
            self.cx.execute("INSERT INTO movimientos (fecha, lote_id, tipo, cantidad, nota) VALUES (?,?,?,?,?)",
                            (fecha, lote_id, "AJUSTE", cambio, nota))
        return cambio

    # ── Historial de cambios ────────────────────────────────────────────────

    def anotar_cambio(self, origen: str, usuario: str, campo: str, antes, despues,
                      medicamento_id: int | None = None, lote_id: int | None = None) -> None:
        with self.cx:
            self.cx.execute("""INSERT INTO cambios (cuando, origen, usuario, medicamento_id, lote_id, campo,
                               antes, despues) VALUES (?,?,?,?,?,?,?,?)""",
                            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), origen, usuario, medicamento_id,
                             lote_id, campo, "" if antes is None else str(antes),
                             "" if despues is None else str(despues)))

    def cambios_pendientes(self) -> list[sqlite3.Row]:
        return self.cx.execute("""
            SELECT c.*, m.nombre, l.lote AS nombre_lote FROM cambios c
            LEFT JOIN medicamentos m ON m.id = c.medicamento_id
            LEFT JOIN lotes l ON l.id = c.lote_id
            WHERE c.revisado = 0 ORDER BY c.cuando DESC, c.id DESC""").fetchall()

    def historial_cambios(self, limite: int = 500) -> list[sqlite3.Row]:
        return self.cx.execute("""
            SELECT c.*, m.nombre, l.lote AS nombre_lote FROM cambios c
            LEFT JOIN medicamentos m ON m.id = c.medicamento_id
            LEFT JOIN lotes l ON l.id = c.lote_id
            ORDER BY c.cuando DESC, c.id DESC LIMIT ?""", (limite,)).fetchall()

    def marcar_revisados(self, ids: list[int]) -> None:
        with self.cx:
            self.cx.executemany("UPDATE cambios SET revisado = 1 WHERE id = ?", [(i,) for i in ids])

    # ── Valores guardados (meta y foto del Excel) ───────────────────────────

    def meta(self, clave: str, defecto: str | None = None) -> str | None:
        fila = self.cx.execute("SELECT valor FROM meta WHERE clave = ?", (clave,)).fetchone()
        return fila[0] if fila else defecto

    def poner_meta(self, clave: str, valor) -> None:
        with self.cx:
            self.cx.execute("INSERT OR REPLACE INTO meta VALUES (?, ?)", (clave, str(valor)))

    def foto_excel(self) -> dict[str, str]:
        return dict(self.cx.execute("SELECT clave, valor FROM excel_foto").fetchall())

    def guardar_foto_excel(self, foto: dict[str, str], reemplazar: bool) -> None:
        with self.cx:
            if reemplazar:
                self.cx.execute("DELETE FROM excel_foto")
            self.cx.executemany("INSERT OR REPLACE INTO excel_foto VALUES (?, ?)", list(foto.items()))
