# -*- coding: utf-8 -*-
"""
1) Reporte de carga: llena la plantilla original (plantilla.xlsm) escribiendo
   SOLO las celdas de captura de la hoja HCARGA. Todo lo demás (COMPARAR,
   fórmulas, tablas, tablas dinámicas, gráficas, macros, formatos) se copia
   byte por byte, sin abrirlo con ninguna librería, para que quede idéntico.
2) Reporte de incidencias: archivo aparte (.xlsx).
"""
import re
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape

NUM_POS = 26                 # posiciones del camión
SLOTS_EXTRA = (27, 28)       # bloques sobrantes de HCARGA para partes extra (>6)
FILAS_POR_POS = 6
POS_POR_GRUPO = 7
# (location, área, No. Parte, Cant.) de cada grupo de HCARGA
GRUPOS = [("D", "E", "F", "G"), ("H", "I", "J", "K"), ("L", "M", "N", "O"), ("P", "Q", "R", "S")]
FILA_INICIO = 12
CELDA_FECHA, CELDA_TRK, CELDA_NOMBRE, CELDA_TURNO = "D5", "H5", "J5", "T5"
MAX_EXTRA = len(SLOTS_EXTRA) * FILAS_POR_POS
MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ------------------------------------------------------------ utilidades XML
def _col_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def _partes_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def _celda_xml(ref, estilo, valor):
    s = f' s="{estilo}"' if estilo is not None else ""
    if valor is None or valor == "":
        return f'<c r="{ref}"{s}/>'
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f'<c r="{ref}"{s}><v>{valor}</v></c>'
    txt = escape(str(valor))
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>'


def _poner(xml, ref, valor):
    """Reemplaza el valor de una celda conservando su estilo (s=...)."""
    col, fila = _partes_ref(ref)
    m_fila = re.search(r'<row r="%d"[^>]*?(/>|>(.*?)</row>)' % fila, xml, re.S)
    if not m_fila:
        # fila inexistente: crearla en orden
        nueva = f'<row r="{fila}">{_celda_xml(ref, None, valor)}</row>'
        sig = None
        for m in re.finditer(r'<row r="(\d+)"', xml):
            if int(m.group(1)) > fila:
                sig = m.start()
                break
        if sig is None:
            sig = xml.index("</sheetData>")
        return xml[:sig] + nueva + xml[sig:]

    fila_xml = m_fila.group(0)
    if fila_xml.endswith("/>") and "</row>" not in fila_xml:
        nueva_fila = fila_xml[:-2] + ">" + _celda_xml(ref, None, valor) + "</row>"
        return xml[:m_fila.start()] + nueva_fila + xml[m_fila.end():]

    m_c = re.search(r'<c r="%s"(\s[^>]*?)?(/>|>.*?</c>)' % ref, fila_xml, re.S)
    if m_c:
        attrs = m_c.group(1) or ""
        m_s = re.search(r'\ss="(\d+)"', attrs)
        nueva = _celda_xml(ref, m_s.group(1) if m_s else None, valor)
        fila_nueva = fila_xml[:m_c.start()] + nueva + fila_xml[m_c.end():]
    else:
        # celda inexistente: insertarla en orden de columna
        pos_ins = fila_xml.rindex("</row>")
        for m in re.finditer(r'<c r="([A-Z]+)\d+"', fila_xml):
            if _col_num(m.group(1)) > _col_num(col):
                pos_ins = m.start()
                break
        fila_nueva = fila_xml[:pos_ins] + _celda_xml(ref, None, valor) + fila_xml[pos_ins:]
    return xml[:m_fila.start()] + fila_nueva + xml[m_fila.end():]


def _ruta_hoja(z, nombre):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(nombre), wb)
    rid = m.group(1)
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
        re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    destino = m.group(1).lstrip("/")
    return destino if destino.startswith("xl/") else "xl/" + destino


def _serial_excel(dt):
    return round((dt - datetime(1899, 12, 30)).total_seconds() / 86400, 6)


# ------------------------------------------------------------ reporte de carga
def celdas_bloque(slot):
    """Celdas (área, [(parte, cant) x6]) del bloque/location 1..28 de HCARGA."""
    g, k = divmod(slot - 1, POS_POR_GRUPO)
    _, c_area, c_parte, c_cant = GRUPOS[g]
    f0 = FILA_INICIO + k * FILAS_POR_POS
    return f"{c_area}{f0}", [(f"{c_parte}{f0 + i}", f"{c_cant}{f0 + i}") for i in range(FILAS_POR_POS)]


def contar_extras(audit):
    return sum(max(0, len(d.get("lineas", [])) - FILAS_POR_POS) for d in audit.get("posiciones", {}).values())


def valores_hcarga(audit):
    """Diccionario {celda: valor} con todo lo que se escribe en HCARGA."""
    vals = {}
    try:
        fecha = datetime.fromisoformat(audit.get("fecha"))
    except Exception:
        fecha = datetime.now()
    vals[CELDA_FECHA] = _serial_excel(fecha)
    vals[CELDA_TRK] = audit.get("trk", "")
    vals[CELDA_NOMBRE] = audit.get("auditor", "")
    vals[CELDA_TURNO] = audit.get("turno", "")

    extras = []
    for pos in range(1, NUM_POS + 1):
        d = audit.get("posiciones", {}).get(str(pos), {})
        lineas = d.get("lineas", [])
        c_area, filas = celdas_bloque(pos)
        vals[c_area] = d.get("area", "") if lineas else ""
        for i, (cp, cc) in enumerate(filas):
            ln = lineas[i] if i < len(lineas) else None
            vals[cp] = ln["parte"] if ln else ""
            vals[cc] = ln["cant"] if ln else ""
        extras += [(pos, ln) for ln in lineas[FILAS_POR_POS:]]

    if len(extras) > MAX_EXTRA:
        raise ValueError(f"Hay {len(extras)} números extra (más de 6 por posición); "
                         f"el formato solo tiene lugar para {MAX_EXTRA}.")
    for j, slot in enumerate(SLOTS_EXTRA):
        c_area, filas = celdas_bloque(slot)
        bloque = extras[j * FILAS_POR_POS:(j + 1) * FILAS_POR_POS]
        vals[c_area] = ("EXTRA " + ",".join(f"P{p}" for p in sorted({p for p, _ in bloque}))) if bloque else ""
        for i, (cp, cc) in enumerate(filas):
            vals[cp] = bloque[i][1]["parte"] if i < len(bloque) else ""
            vals[cc] = bloque[i][1]["cant"] if i < len(bloque) else ""
    return vals


def generar_carga(audit, plantilla, salida):
    """Copia la plantilla y escribe la captura en HCARGA. Regresa la ruta."""
    with zipfile.ZipFile(plantilla) as zin:
        hoja = _ruta_hoja(zin, "HCARGA")
        xml = zin.read(hoja).decode("utf-8")
        for ref, v in valores_hcarga(audit).items():
            xml = _poner(xml, ref, v)
        wb = zin.read("xl/workbook.xml").decode("utf-8")
        # Excel recalcula COMPARAR al abrir
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wb, count=1)
        with zipfile.ZipFile(salida, "w") as zout:
            for info in zin.infolist():
                if info.filename == hoja:
                    datos = xml.encode("utf-8")
                elif info.filename == "xl/workbook.xml":
                    datos = wb.encode("utf-8")
                else:
                    datos = zin.read(info.filename)
                zout.writestr(info, datos, compress_type=info.compress_type)
    return salida


def nombre_carga(audit):
    try:
        f = datetime.fromisoformat(audit.get("fecha"))
    except Exception:
        f = datetime.now()
    trk = "".join(ch for ch in str(audit.get("trk", "")) if ch.isalnum()) or "SINTRK"
    return f"Carga_{trk}_{f:%Y-%m-%d_%H%M}.xlsm"


# ------------------------------------------------------------ reporte de incidencias
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def generar_incidencias(incidencias, salida):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "INCIDENCIAS"
    enc = ["Nombre", "Tipo", "Categoria", "Fecha", "Mes", "Posición", "TRK", "Auditor"]
    anchos = [20, 48, 30, 18, 12, 10, 12, 20]
    linea = Side(style="thin", color="808080")
    borde = Border(left=linea, right=linea, top=linea, bottom=linea)
    for i, (t, a) in enumerate(zip(enc, anchos), 1):
        c = ws.cell(1, i, t)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
        c.border = borde
        ws.column_dimensions[c.column_letter].width = a
    for r, inc in enumerate(incidencias, 2):
        try:
            f = datetime.fromisoformat(inc["fecha"])
        except Exception:
            f = datetime.now()
        fila = [inc.get("nombre", ""), inc.get("tipo", ""), inc.get("categoria", ""), f,
                MESES[f.month - 1], inc.get("posicion", ""), inc.get("trk", ""), inc.get("auditor", "")]
        for col, v in enumerate(fila, 1):
            c = ws.cell(r, col, v)
            c.border = borde
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(r, 4).number_format = "dd/mm/yyyy hh:mm"
    ws.freeze_panes = "A2"
    if incidencias:
        ws.auto_filter.ref = f"A1:H{len(incidencias) + 1}"
    wb.save(salida)
    return salida


def nombre_incidencias():
    return f"Incidencias_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
