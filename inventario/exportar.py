"""Exportar cualquier tabla del programa a Excel (.xlsx) o PDF."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

AZUL = "1F4E79"
AZUL_CLARO = "DDEBF7"


@dataclass
class Hoja:
    """Lo que se exporta: un título, encabezados y filas.

    `colores` = {(fila, columna): 'RRGGBB'} para resaltar celdas (p. ej. vencidos).
    `grupos` = encabezado superior opcional [(texto, desde_col, hasta_col)], para Lun–Vie.
    """
    titulo: str
    encabezados: list[str]
    filas: list[list]
    subtitulo: str = ""
    nombre_hoja: str = "Datos"
    colores: dict[tuple[int, int], str] = field(default_factory=dict)
    grupos: list[tuple[str, int, int]] = field(default_factory=list)
    totales: list | None = None


def _texto(v) -> str:
    return "" if v is None else str(v)


def a_excel(hojas: list[Hoja], destino: Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    borde = Side(style="thin", color="999999")
    caja = Border(left=borde, right=borde, top=borde, bottom=borde)
    for h in hojas:
        ws = wb.create_sheet(h.nombre_hoja[:31])
        ncol = len(h.encabezados)
        ws.cell(1, 1, h.titulo).font = Font(bold=True, size=14, color=AZUL)
        ws.cell(2, 1, h.subtitulo or f"Generado: {datetime.now():%d/%m/%Y %H:%M}").font = Font(italic=True, color="666666")
        fila = 4
        if h.grupos:
            for texto, c1, c2 in h.grupos:
                ws.merge_cells(start_row=fila, start_column=c1 + 1, end_row=fila, end_column=c2 + 1)
                celda = ws.cell(fila, c1 + 1, texto)
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill("solid", fgColor=AZUL)
                celda.alignment = Alignment(horizontal="center")
            fila += 1
        for c, enc in enumerate(h.encabezados, 1):
            celda = ws.cell(fila, c, enc)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor=AZUL)
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = caja
        ws.freeze_panes = ws.cell(fila + 1, 1)
        inicio = fila + 1
        for i, datos in enumerate(h.filas):
            for c, v in enumerate(datos, 1):
                celda = ws.cell(inicio + i, c, v)
                celda.border = caja
                color = h.colores.get((i, c - 1))
                if color:
                    celda.fill = PatternFill("solid", fgColor=color)
                elif i % 2:
                    celda.fill = PatternFill("solid", fgColor="F5F9FC")
        if h.totales:
            f = inicio + len(h.filas)
            for c, v in enumerate(h.totales, 1):
                celda = ws.cell(f, c, v)
                celda.font = Font(bold=True)
                celda.fill = PatternFill("solid", fgColor=AZUL_CLARO)
                celda.border = caja
        for c in range(1, ncol + 1):
            largo = max([len(_texto(h.encabezados[c - 1]))] +
                        [len(_texto(f[c - 1])) for f in h.filas if c - 1 < len(f)])
            ws.column_dimensions[get_column_letter(c)].width = min(max(8, largo + 2), 45)
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
    return destino


def a_pdf(hojas: list[Hoja], destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(destino), pagesize=landscape(letter),
                            leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm)
    estilos = getSampleStyleSheet()
    titulo = ParagraphStyle("t", parent=estilos["Title"], fontSize=15, textColor=colors.HexColor("#" + AZUL),
                            alignment=0, spaceAfter=2)
    sub = ParagraphStyle("s", parent=estilos["Normal"], fontSize=9, textColor=colors.grey)
    celda = ParagraphStyle("c", parent=estilos["Normal"], fontSize=7.5, leading=9)
    celda_enc = ParagraphStyle("ce", parent=celda, textColor=colors.white, fontName="Helvetica-Bold",
                               alignment=1)
    ancho_util = landscape(letter)[0] - 2.4 * cm
    partes = []
    for n, h in enumerate(hojas):
        if n:
            from reportlab.platypus import PageBreak
            partes.append(PageBreak())
        partes.append(Paragraph(h.titulo, titulo))
        partes.append(Paragraph(h.subtitulo or f"Generado: {datetime.now():%d/%m/%Y %H:%M}", sub))
        partes.append(Spacer(1, 6))
        datos = []
        if h.grupos:
            fila_grupo = [""] * len(h.encabezados)
            for texto, c1, _ in h.grupos:
                fila_grupo[c1] = Paragraph(texto, celda_enc)
            datos.append(fila_grupo)
        datos.append([Paragraph(_texto(e), celda_enc) for e in h.encabezados])
        datos += [[Paragraph(_texto(v), celda) for v in f] for f in h.filas]
        if h.totales:
            datos.append([Paragraph(f"<b>{_texto(v)}</b>", celda) for v in h.totales])
        # Ancho natural de cada columna (aprox. 4.6 pt por letra a 7.5 pt + 18 de relleno).
        naturales = []
        for c in range(len(h.encabezados)):
            largo_enc = max((len(p) for p in _texto(h.encabezados[c]).split()), default=0)
            largo = max([largo_enc] + [len(_texto(f[c])) for f in h.filas[:300] if c < len(f)])
            naturales.append(max(32, min(largo, 45) * 4.6 + 18))
        total = sum(naturales)
        if total > ancho_util:
            anchos = [n * ancho_util / total for n in naturales]
        else:
            # Lo que sobra se reparte entre las columnas de texto largo.
            anchas = [n if n > 80 else 0 for n in naturales]
            extra = ancho_util - total
            suma = sum(anchas) or 1
            anchos = [n + extra * a / suma for n, a in zip(naturales, anchas)]
        tabla = Table(datos, colWidths=anchos, repeatRows=2 if h.grupos else 1)
        enc_filas = 2 if h.grupos else 1
        estilo = [
            ("BACKGROUND", (0, 0), (-1, enc_filas - 1), colors.HexColor("#" + AZUL)),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, enc_filas), (-1, -1), [colors.white, colors.HexColor("#F5F9FC")]),
        ]
        for texto, c1, c2 in h.grupos:
            estilo.append(("SPAN", (c1, 0), (c2, 0)))
        for (f, c), color in h.colores.items():
            estilo.append(("BACKGROUND", (c, f + enc_filas), (c, f + enc_filas), colors.HexColor("#" + color)))
        if h.totales:
            estilo.append(("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#" + AZUL_CLARO)))
        tabla.setStyle(TableStyle(estilo))
        partes.append(tabla)

    def pie(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(landscape(letter)[0] - 1.2 * cm, 0.7 * cm, f"Página {doc_.page}")
        canvas.restoreState()

    doc.build(partes, onFirstPage=pie, onLaterPages=pie)
    return destino
