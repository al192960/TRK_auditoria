# -*- coding: utf-8 -*-
"""
Genera el Excel de la auditoría con el mismo acomodo que el libro
"Borrador de carga" (hoja HCARGA e INCIDENCIAS) para poder copiar/pegar.
"""
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NUM_POS = 26
FILAS_POR_POS = 6
POS_POR_GRUPO = 7
# Columnas (location, área, No. Parte, Cant.) de cada grupo, igual que HCARGA
GRUPOS = [(4, 5, 6, 7), (8, 9, 10, 11), (12, 13, 14, 15), (16, 17, 18, 19)]
FILA_INICIO = 12

AZUL = PatternFill("solid", fgColor="1F4E78")
GRIS = PatternFill("solid", fgColor="D9E1F2")
NARANJA = PatternFill("solid", fgColor="FCE4D6")
BLANCO_B = Font(bold=True, color="FFFFFF")
NEGRITA = Font(bold=True)
DELG = Side(style="thin", color="808080")
BORDE = Border(left=DELG, right=DELG, top=DELG, bottom=DELG)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _encabezado(ws, fila, titulos, anchos=None):
    for i, t in enumerate(titulos, start=1):
        c = ws.cell(fila, i, t)
        c.font = BLANCO_B
        c.fill = AZUL
        c.alignment = CENTRO
        c.border = BORDE
    if anchos:
        for i, a in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = ws.cell(fila + 1, 1)


def _fecha(audit):
    try:
        return datetime.fromisoformat(audit.get("fecha"))
    except Exception:
        return datetime.now()


def _hcarga(wb, audit):
    ws = wb.active
    ws.title = "HCARGA"
    fecha = _fecha(audit)

    ws.merge_cells("D1:V3")
    c = ws["D1"]
    c.value = "FORMATO DE CARGA"
    c.font = Font(bold=True, size=18)
    c.alignment = CENTRO

    ws.merge_cells("D5:F5")
    ws["D5"] = fecha
    ws["D5"].number_format = "dd/mm/yyyy hh:mm"
    ws["H5"] = audit.get("trk", "")
    ws.merge_cells("J5:P5")
    ws["J5"] = audit.get("auditor", "")
    ws.merge_cells("T5:U5")
    ws["T5"] = audit.get("turno", "")
    for ref in ("D5", "H5", "J5", "T5"):
        ws[ref].font = NEGRITA
        ws[ref].alignment = CENTRO
    ws["D4"], ws["H4"], ws["J4"], ws["T4"] = "Fecha", "TRK", "Auditor", "Turno"

    for loc, area, parte, cant in GRUPOS:
        ws.merge_cells(start_row=11, start_column=loc, end_row=11, end_column=area)
        for col, txt in ((loc, "location"), (parte, "No. Parte"), (cant, "Cant.")):
            c = ws.cell(11, col, txt)
            c.font = BLANCO_B
            c.fill = AZUL
            c.alignment = CENTRO
        ws.cell(11, area).fill = AZUL
        ws.column_dimensions[get_column_letter(loc)].width = 6
        ws.column_dimensions[get_column_letter(area)].width = 10
        ws.column_dimensions[get_column_letter(parte)].width = 16
        ws.column_dimensions[get_column_letter(cant)].width = 8

    posiciones = audit.get("posiciones", {})
    for pos in range(1, NUM_POS + 1):
        g, k = divmod(pos - 1, POS_POR_GRUPO)
        loc, area, parte, cant = GRUPOS[g]
        f0 = FILA_INICIO + k * FILAS_POR_POS
        f1 = f0 + FILAS_POR_POS - 1
        datos = posiciones.get(str(pos), {})
        lineas = datos.get("lineas", [])
        extra = len(lineas) > FILAS_POR_POS

        ws.merge_cells(start_row=f0, start_column=loc, end_row=f1, end_column=loc)
        ws.merge_cells(start_row=f0, start_column=area, end_row=f1, end_column=area)
        c = ws.cell(f0, loc, pos)
        c.font = Font(bold=True, size=14)
        c.alignment = CENTRO
        c = ws.cell(f0, area, datos.get("area", "") or None)
        c.alignment = CENTRO
        if extra:
            c.fill = NARANJA

        for i in range(FILAS_POR_POS):
            if i < len(lineas):
                ws.cell(f0 + i, parte, lineas[i]["parte"])
                ws.cell(f0 + i, cant, lineas[i]["cant"])
        for r in range(f0, f1 + 1):
            for col in (loc, area, parte, cant):
                ws.cell(r, col).border = BORDE
        if extra:
            ws.cell(f0, loc).fill = NARANJA
    ws.cell(55, 4, "Las posiciones en naranja tienen más de 6 números de parte; "
                   "ver hoja DETALLE e INCIDENCIAS.").font = Font(italic=True, color="C65911")
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1


def _detalle(wb, audit):
    ws = wb.create_sheet("DETALLE")
    _encabezado(ws, 1, ["Posición", "Área", "#", "No. Parte", "Cantidad", "Extra (>6)"],
                [10, 14, 5, 20, 11, 11])
    fila = 2
    for pos in range(1, NUM_POS + 1):
        datos = audit.get("posiciones", {}).get(str(pos), {})
        for i, ln in enumerate(datos.get("lineas", []), start=1):
            valores = [pos, datos.get("area", ""), i, ln["parte"], ln["cant"],
                       "SÍ" if i > FILAS_POR_POS else ""]
            for col, v in enumerate(valores, start=1):
                c = ws.cell(fila, col, v)
                c.border = BORDE
                if i > FILAS_POR_POS:
                    c.fill = NARANJA
            fila += 1
    if fila > 2:
        ws.auto_filter.ref = f"A1:F{fila - 1}"


def _resumen(wb, audit):
    ws = wb.create_sheet("RESUMEN")
    _encabezado(ws, 1, ["MATERIAL", "QTY", "Pallets / líneas", "Posiciones"], [20, 12, 16, 30])
    tot = {}
    for pos, datos in audit.get("posiciones", {}).items():
        for ln in datos.get("lineas", []):
            t = tot.setdefault(ln["parte"], [0, 0, set()])
            t[0] += ln["cant"] or 0
            t[1] += 1
            t[2].add(int(pos))
    fila = 2
    for parte in sorted(tot):
        q, n, ps = tot[parte]
        for col, v in enumerate([parte, q, n, ", ".join(str(p) for p in sorted(ps))], 1):
            ws.cell(fila, col, v).border = BORDE
        fila += 1
    c = ws.cell(fila, 1, "TOTAL")
    c.font = NEGRITA
    c = ws.cell(fila, 2, f"=SUM(B2:B{fila - 1})" if fila > 2 else 0)
    c.font = NEGRITA


def _incidencias(wb, audit):
    ws = wb.create_sheet("INCIDENCIAS")
    _encabezado(ws, 1, ["Nombre", "Tipo", "Categoria", "Fecha", "Mes", "Posición", "TRK", "Auditor"],
                [18, 45, 28, 18, 12, 10, 10, 18])
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
             "septiembre", "octubre", "noviembre", "diciembre"]
    fila = 2
    for inc in audit.get("incidencias", []):
        try:
            f = datetime.fromisoformat(inc["fecha"])
        except Exception:
            f = datetime.now()
        valores = [inc.get("nombre", "").upper(), inc.get("tipo", ""), inc.get("categoria", ""),
                   f, meses[f.month - 1], inc.get("posicion", ""), audit.get("trk", ""),
                   audit.get("auditor", "")]
        for col, v in enumerate(valores, 1):
            c = ws.cell(fila, col, v)
            c.border = BORDE
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(fila, 4).number_format = "dd/mm/yyyy hh:mm"
        fila += 1
    if fila == 2:
        ws.cell(2, 1, "Sin incidencias")


def generar_excel(audit, ruta):
    wb = Workbook()
    _hcarga(wb, audit)
    _detalle(wb, audit)
    _resumen(wb, audit)
    _incidencias(wb, audit)
    wb.save(ruta)
    return ruta


def nombre_archivo(audit):
    f = _fecha(audit)
    trk = "".join(ch for ch in str(audit.get("trk", "")) if ch.isalnum()) or "SINTRK"
    return f"Auditoria_{trk}_{f:%Y-%m-%d_%H%M}.xlsx"
