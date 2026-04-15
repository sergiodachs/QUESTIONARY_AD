# -*- coding: utf-8 -*-
# listado_reservas_fantasma_y_ubicaciones_nav_mejorado.py
#
# OBJETIVO
# --------
# Generar un Excel práctico para:
# 1) Ver productos por ubicación a partir de Item Ledger Entry agrupado
# 2) Detectar candidatos a RESERVAS FANTASMA / residuos raros de stock
#
# HOJAS
# -----
# 1) DETALLE_ILE
# 2) RESUMEN_UBICACIONES_ILE
# 3) RESERVAS_FANTASMA
# 4) RESUMEN_FANTASMA
# 5) TOP_BINS_PROBLEMATICOS
# 6) TOP_ITEMS_PROBLEMATICOS
# 7) LEYENDA
#
# REQUISITOS
# ----------
# pip install pyodbc openpyxl
#
# NOTA
# ----
# Este informe NO sustituye al stock "bueno" de contenido ubicación.
# Está orientado a detectar residuos, negativos y reservas fantasma
# usando Item Ledger Entry agrupado.

import os
import sys
import math
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox

import pyodbc
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.page import PageMargins


# =========================================================
# CONFIG DB
# =========================================================
SERVER = "192.168.49.144"
DATABASE = "NAV2009"
USERNAME = "webadmin"
PASSWORD = "webadmin"

POSSIBLE_DRIVERS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server",
]

T_ILE = "[Item Ledger Entry]"
T_ITEM = "[Item]"

OUT_DIR = r"C:\INFORMEUBICACIONES"

# =========================================================
# EMPRESAS A IGNORAR
# =========================================================
# Si en vuestra BD Empresa usa códigos numéricos, añádelos aquí como texto.
EXCLUDED_EMPRESA_VALUES = {
    "DEVETECH",
    "PRUEBAS",
}

# =========================================================
# UMBRALES VISUALES Y DE CLASIFICACION
# =========================================================
QTY_SMALL_WARNING = 2
QTY_HIGH_STOCK = 100

# Para marcar casos especialmente sospechosos
GHOST_STRONG_NEGATIVE_THRESHOLD = -1e-12


# =========================================================
# GUI
# =========================================================
def ask_filters():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    messagebox.showinfo(
        "Listado ubicaciones + reservas fantasma",
        "Este programa genera un Excel con varias hojas de análisis.\n\n"
        "Se usa Item Ledger Entry agrupado para detectar residuos y reservas fantasma.\n"
        "Puedes dejar los filtros vacíos para sacar TODO."
    )

    location_code = simpledialog.askstring(
        "Filtro Location Code",
        "Location Code (almacén)\n\nEjemplo: 50\nDeja vacío para TODOS:",
        parent=root
    )

    bin_code = simpledialog.askstring(
        "Filtro Bin Code",
        "Bin Code exacto o parte del código\n\nEjemplo: 97A08DC\n"
        "También admite parcial: 97A08\n\nDeja vacío para TODOS:",
        parent=root
    )

    item_no = simpledialog.askstring(
        "Filtro Item No",
        "Código de producto (Item No)\n\nEjemplo: 214397\nDeja vacío para TODOS:",
        parent=root
    )

    if location_code is None and bin_code is None and item_no is None:
        root.destroy()
        sys.exit(0)

    location_code = (location_code or "").strip()
    bin_code = (bin_code or "").strip()
    item_no = (item_no or "").strip()

    root.destroy()
    return location_code, bin_code, item_no


def show_info(title, text):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(title, text)
    root.destroy()


def show_error(title, text):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showerror(title, text)
    root.destroy()


# =========================================================
# UTILIDADES
# =========================================================
def ensure_dir(path):
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)


def connect_any():
    last_err = None
    for drv in POSSIBLE_DRIVERS:
        try:
            cn = pyodbc.connect(
                f"DRIVER={{{drv}}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}",
                timeout=20
            )
            return cn
        except Exception as e:
            last_err = e
    raise RuntimeError(f"No se pudo abrir conexión ODBC.\n{last_err}")


def parse_sql_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    if not s:
        return None
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d")
    except Exception:
        return None


def safe_str(v):
    if v is None:
        return ""
    return str(v).strip()


def fmt_refs(refs, max_chars=25000):
    refs = sorted(set([safe_str(x) for x in refs if safe_str(x)]))
    if not refs:
        return "", ""

    out = []
    used = 0
    total = len(refs)

    for i, r in enumerate(refs):
        piece = (", " if i > 0 else "") + r
        current = "".join(out)
        if len(current) + len(piece) > max_chars:
            used = i
            break
        out.append(piece)
        used = i + 1

    txt = "".join(out)
    info = ""
    if used < total:
        info = f"Mostradas {used} de {total} referencias"
    return txt, info


def auto_row_height(text, chars_per_line=40, min_height=18, line_height=15):
    s = safe_str(text)
    if not s:
        return min_height
    lines = 0
    for part in s.split("\n"):
        lines += max(1, math.ceil(len(part) / chars_per_line))
    return max(min_height, lines * line_height)


def format_decimal_es(n, decimals=2):
    try:
        s = f"{float(n):,.{decimals}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return ""


def get_fill_for_qty(qty):
    if qty < 0:
        return PatternFill(fill_type="solid", fgColor="F4CCCC")  # rojo suave
    if abs(qty) < 1e-12:
        return PatternFill(fill_type="solid", fgColor="E7E6E6")  # gris
    if 0 < qty <= QTY_SMALL_WARNING:
        return PatternFill(fill_type="solid", fgColor="FFF2CC")  # amarillo
    if qty >= QTY_HIGH_STOCK:
        return PatternFill(fill_type="solid", fgColor="D9EAD3")  # verde
    return None


def get_fill_for_issue(issue_type):
    issue_type = safe_str(issue_type).upper()
    if "RESERVA_FANTASMA_CANDIDATA" in issue_type:
        return PatternFill(fill_type="solid", fgColor="FCE5CD")  # naranja suave
    if "NEGATIVO" in issue_type:
        return PatternFill(fill_type="solid", fgColor="F4CCCC")  # rojo
    if "CERO_CON_LOTE" in issue_type or "CERO_CON_FECHA" in issue_type:
        return PatternFill(fill_type="solid", fgColor="FFF2CC")  # amarillo
    return None


def priority_label(priority):
    if priority == 1:
        return "ALTA"
    if priority == 2:
        return "MEDIA"
    return "BAJA"


# =========================================================
# SQL
# =========================================================
def build_sql_detail_ile(location_code, bin_code, item_no):
    excluded = sorted(EXCLUDED_EMPRESA_VALUES)

    sql = rf"""
;WITH ILE_SUM AS (
    SELECT
        CAST(ile.[Empresa] AS NVARCHAR(100)) AS Empresa,
        ile.[Location Code] AS LocationCode,
        ile.[Cód_ ubicación] AS BinCode,
        ile.[Item No_] AS ItemNo,
        ile.[Lot No_] AS LotNo,
        SUM(CAST(ile.[Remaining Quantity] AS DECIMAL(38,5))) AS Cantidad,
        MAX(CASE WHEN ile.[Fecha fabricación] > '1900-01-01' THEN ile.[Fecha fabricación] END) AS FechaFab_Max,
        MAX(CASE WHEN ile.[Warranty Date] > '1900-01-01' THEN ile.[Warranty Date] END) AS Warranty_Max
    FROM {T_ILE} ile
    WHERE 1 = 1
"""
    params = []

    if excluded:
        placeholders = ",".join("?" for _ in excluded)
        sql += f" AND CAST(ile.[Empresa] AS NVARCHAR(100)) NOT IN ({placeholders})\n"
        params.extend(excluded)

    if location_code:
        sql += " AND ile.[Location Code] = ?\n"
        params.append(location_code)

    if bin_code:
        sql += " AND ile.[Cód_ ubicación] LIKE ?\n"
        params.append(f"%{bin_code}%")

    if item_no:
        sql += " AND ile.[Item No_] = ?\n"
        params.append(item_no)

    sql += rf"""
    GROUP BY
        CAST(ile.[Empresa] AS NVARCHAR(100)),
        ile.[Location Code],
        ile.[Cód_ ubicación],
        ile.[Item No_],
        ile.[Lot No_]
    HAVING SUM(CAST(ile.[Remaining Quantity] AS DECIMAL(38,5))) <> 0
)
SELECT
    s.Empresa,
    s.LocationCode,
    s.BinCode,
    s.ItemNo,
    it.[Description] AS Descripcion,
    CAST(s.Cantidad AS DECIMAL(38,5)) AS Cantidad,
    s.LotNo,
    COALESCE(s.FechaFab_Max, s.Warranty_Max) AS FechaReferencia,
    CAST(ISNULL(it.[Unit Cost], 0) AS DECIMAL(38,5)) AS CosteUnitario,
    CAST(s.Cantidad * ISNULL(it.[Unit Cost], 0) AS DECIMAL(38,5)) AS ValorStock
FROM ILE_SUM s
LEFT JOIN {T_ITEM} it
    ON it.[No_] = s.ItemNo
ORDER BY
    s.LocationCode,
    s.BinCode,
    s.ItemNo,
    s.LotNo;
"""
    return sql, params


def fetch_detail_ile_rows(location_code, bin_code, item_no):
    cn = connect_any()
    cur = cn.cursor()

    sql, params = build_sql_detail_ile(location_code, bin_code, item_no)
    cur.execute(sql, params)

    cols = [d[0] for d in cur.description]
    rows = []

    for tup in cur.fetchall():
        rec = dict(zip(cols, tup))
        rec["FechaReferencia"] = parse_sql_date(rec.get("FechaReferencia"))
        rec["Cantidad"] = float(rec.get("Cantidad") or 0)
        rec["CosteUnitario"] = float(rec.get("CosteUnitario") or 0)
        rec["ValorStock"] = float(rec.get("ValorStock") or 0)
        rec["Empresa"] = safe_str(rec.get("Empresa"))
        rec["LocationCode"] = safe_str(rec.get("LocationCode"))
        rec["BinCode"] = safe_str(rec.get("BinCode"))
        rec["ItemNo"] = safe_str(rec.get("ItemNo"))
        rec["Descripcion"] = safe_str(rec.get("Descripcion"))
        rec["LotNo"] = safe_str(rec.get("LotNo"))
        rows.append(rec)

    cur.close()
    cn.close()
    return rows


# =========================================================
# RESERVAS FANTASMA / HEURISTICAS
# =========================================================
def build_ghost_rows(detail_rows):
    ghost_rows = []

    for r in detail_rows:
        qty = float(r["Cantidad"] or 0)
        lot = safe_str(r["LotNo"])
        fecha = r["FechaReferencia"]

        tipo = None
        prioridad = 99
        comentario = ""

        if qty < GHOST_STRONG_NEGATIVE_THRESHOLD and (lot or fecha is not None):
            tipo = "RESERVA_FANTASMA_CANDIDATA"
            prioridad = 1
            comentario = "Cantidad negativa con lote y/o fecha. Muy sospechoso de reserva fantasma."
        elif qty < GHOST_STRONG_NEGATIVE_THRESHOLD:
            tipo = "NEGATIVO_ILE"
            prioridad = 2
            comentario = "Cantidad negativa en ILE agrupado."
        elif abs(qty) < 1e-12 and (lot or fecha is not None):
            tipo = "CERO_CON_LOTE_O_FECHA"
            prioridad = 3
            comentario = "Cantidad 0 pero conserva lote o fecha. Revisar residuo."
        else:
            continue

        ghost_rows.append({
            "TipoProblema": tipo,
            "Prioridad": prioridad,
            "Nivel": priority_label(prioridad),
            "Empresa": r["Empresa"],
            "LocationCode": r["LocationCode"],
            "BinCode": r["BinCode"],
            "ItemNo": r["ItemNo"],
            "Descripcion": r["Descripcion"],
            "Cantidad": qty,
            "LotNo": lot,
            "FechaReferencia": fecha,
            "CosteUnitario": float(r["CosteUnitario"] or 0),
            "ValorImpacto": float(r["ValorStock"] or 0),
            "Comentario": comentario,
        })

    ghost_rows.sort(
        key=lambda x: (
            x["Prioridad"],
            x["LocationCode"],
            x["BinCode"],
            x["ItemNo"],
            x["LotNo"],
        )
    )
    return ghost_rows


def build_summary_ubicaciones(detail_rows):
    grouped = {}

    for r in detail_rows:
        key = (r["LocationCode"], r["BinCode"])
        if key not in grouped:
            grouped[key] = {
                "LocationCode": r["LocationCode"],
                "BinCode": r["BinCode"],
                "UnidadesTotales": 0.0,
                "NumArticulos": 0,
                "ValorTotal": 0.0,
                "RefsSet": set(),
            }

        grouped[key]["UnidadesTotales"] += float(r["Cantidad"] or 0)
        grouped[key]["ValorTotal"] += float(r["ValorStock"] or 0)
        grouped[key]["RefsSet"].add(r["ItemNo"])

    rows = []
    for _, g in grouped.items():
        refs_txt, refs_info = fmt_refs(g["RefsSet"])
        rows.append({
            "LocationCode": g["LocationCode"],
            "BinCode": g["BinCode"],
            "UnidadesTotales": g["UnidadesTotales"],
            "NumArticulos": len(g["RefsSet"]),
            "ValorTotal": g["ValorTotal"],
            "Referencias": refs_txt,
            "RefsInfo": refs_info,
        })

    rows.sort(key=lambda x: (x["LocationCode"], x["BinCode"]))
    return rows


def build_summary_ghost(ghost_rows):
    grouped = {}

    for r in ghost_rows:
        tipo = r["TipoProblema"]
        if tipo not in grouped:
            grouped[tipo] = {
                "TipoProblema": tipo,
                "Casos": 0,
                "ImpactoCantidad": 0.0,
                "ImpactoValor": 0.0,
            }

        grouped[tipo]["Casos"] += 1
        grouped[tipo]["ImpactoCantidad"] += abs(float(r["Cantidad"] or 0))
        grouped[tipo]["ImpactoValor"] += abs(float(r["ValorImpacto"] or 0))

    rows = list(grouped.values())
    rows.sort(key=lambda x: x["TipoProblema"])
    return rows


def build_top_bins_problematicos(ghost_rows):
    grouped = {}

    for r in ghost_rows:
        key = (r["LocationCode"], r["BinCode"])
        if key not in grouped:
            grouped[key] = {
                "LocationCode": r["LocationCode"],
                "BinCode": r["BinCode"],
                "Casos": 0,
                "ItemsSet": set(),
                "ImpactoCantidad": 0.0,
                "ImpactoValor": 0.0,
                "TiposSet": set(),
            }

        grouped[key]["Casos"] += 1
        grouped[key]["ItemsSet"].add(r["ItemNo"])
        grouped[key]["TiposSet"].add(r["TipoProblema"])
        grouped[key]["ImpactoCantidad"] += abs(float(r["Cantidad"] or 0))
        grouped[key]["ImpactoValor"] += abs(float(r["ValorImpacto"] or 0))

    rows = []
    for _, g in grouped.items():
        tipos_txt, _ = fmt_refs(g["TiposSet"], max_chars=5000)
        rows.append({
            "LocationCode": g["LocationCode"],
            "BinCode": g["BinCode"],
            "Casos": g["Casos"],
            "NumItems": len(g["ItemsSet"]),
            "ImpactoCantidad": g["ImpactoCantidad"],
            "ImpactoValor": g["ImpactoValor"],
            "TiposProblema": tipos_txt,
        })

    rows.sort(key=lambda x: (-x["Casos"], -x["ImpactoValor"], x["LocationCode"], x["BinCode"]))
    return rows


def build_top_items_problematicos(ghost_rows):
    grouped = {}

    for r in ghost_rows:
        key = r["ItemNo"]
        if key not in grouped:
            grouped[key] = {
                "ItemNo": r["ItemNo"],
                "Descripcion": r["Descripcion"],
                "Casos": 0,
                "BinsSet": set(),
                "ImpactoCantidad": 0.0,
                "ImpactoValor": 0.0,
                "TiposSet": set(),
            }

        grouped[key]["Casos"] += 1
        grouped[key]["BinsSet"].add(f"{r['LocationCode']}|{r['BinCode']}")
        grouped[key]["TiposSet"].add(r["TipoProblema"])
        grouped[key]["ImpactoCantidad"] += abs(float(r["Cantidad"] or 0))
        grouped[key]["ImpactoValor"] += abs(float(r["ValorImpacto"] or 0))

    rows = []
    for _, g in grouped.items():
        tipos_txt, _ = fmt_refs(g["TiposSet"], max_chars=5000)
        rows.append({
            "ItemNo": g["ItemNo"],
            "Descripcion": g["Descripcion"],
            "Casos": g["Casos"],
            "NumBins": len(g["BinsSet"]),
            "ImpactoCantidad": g["ImpactoCantidad"],
            "ImpactoValor": g["ImpactoValor"],
            "TiposProblema": tipos_txt,
        })

    rows.sort(key=lambda x: (-x["Casos"], -x["ImpactoValor"], x["ItemNo"]))
    return rows


# =========================================================
# FORMATO EXCEL
# =========================================================
def apply_base_style(ws):
    thin = Side(style="thin", color="7F7F7F")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill(fill_type="solid", fgColor="D9E1F2")
    header_font = Font(name="Calibri", size=10, bold=True)
    normal_font = Font(name="Calibri", size=10)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = normal_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 22

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    ws.sheet_view.showGridLines = False

    ws.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.5,
        bottom=0.5,
        header=0.2,
        footer=0.2
    )


def set_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


# =========================================================
# EXPORTACION
# =========================================================
def export_xlsx(
    path,
    detail_rows,
    summary_rows,
    ghost_rows,
    ghost_summary_rows,
    top_bins_rows,
    top_items_rows
):
    wb = Workbook()

    # -----------------------------------------------------
    # HOJA DETALLE_ILE
    # -----------------------------------------------------
    ws1 = wb.active
    ws1.title = "DETALLE_ILE"

    headers1 = [
        "Empresa",
        "Location Code",
        "Bin Code",
        "Item No",
        "Descripción",
        "Cantidad",
        "Lote",
        "Fecha referencia",
        "Coste unitario",
        "Valor stock",
    ]
    ws1.append(headers1)

    for r in detail_rows:
        ws1.append([
            r["Empresa"],
            r["LocationCode"],
            r["BinCode"],
            r["ItemNo"],
            r["Descripcion"],
            float(r["Cantidad"]),
            r["LotNo"],
            r["FechaReferencia"],
            float(r["CosteUnitario"]),
            float(r["ValorStock"]),
        ])

    apply_base_style(ws1)

    for row in ws1.iter_rows(min_row=2):
        qty = float(row[5].value or 0)
        fill = get_fill_for_qty(qty)
        if fill:
            for cell in row:
                cell.fill = fill

        row[0].alignment = Alignment(horizontal="center", vertical="center")
        row[1].alignment = Alignment(horizontal="center", vertical="center")
        row[2].alignment = Alignment(horizontal="center", vertical="center")
        row[3].alignment = Alignment(horizontal="center", vertical="center")
        row[4].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        row[5].alignment = Alignment(horizontal="right", vertical="center")
        row[6].alignment = Alignment(horizontal="center", vertical="center")
        row[7].alignment = Alignment(horizontal="center", vertical="center")
        row[8].alignment = Alignment(horizontal="right", vertical="center")
        row[9].alignment = Alignment(horizontal="right", vertical="center")

        row[5].number_format = '#,##0.#####'
        row[7].number_format = 'yyyy-mm-dd'
        row[8].number_format = '#,##0.#####'
        row[9].number_format = '#,##0.00'

        desc = row[4].value or ""
        ws1.row_dimensions[row[0].row].height = auto_row_height(desc, chars_per_line=45)

    set_widths(ws1, {
        "A": 12,
        "B": 12,
        "C": 16,
        "D": 16,
        "E": 42,
        "F": 12,
        "G": 14,
        "H": 14,
        "I": 14,
        "J": 14,
    })

    # -----------------------------------------------------
    # HOJA RESUMEN_UBICACIONES_ILE
    # -----------------------------------------------------
    ws2 = wb.create_sheet("RESUMEN_UBICACIONES_ILE")

    headers2 = [
        "Location Code",
        "Bin Code",
        "Unidades totales",
        "Nº artículos",
        "Valor total",
        "Referencias",
    ]
    ws2.append(headers2)

    for r in summary_rows:
        ws2.append([
            r["LocationCode"],
            r["BinCode"],
            float(r["UnidadesTotales"]),
            int(r["NumArticulos"]),
            float(r["ValorTotal"]),
            r["Referencias"],
        ])

    apply_base_style(ws2)

    for row in ws2.iter_rows(min_row=2):
        qty = float(row[2].value or 0)
        fill = get_fill_for_qty(qty)
        if fill:
            for cell in row:
                cell.fill = fill

        row[0].alignment = Alignment(horizontal="center", vertical="center")
        row[1].alignment = Alignment(horizontal="center", vertical="center")
        row[2].alignment = Alignment(horizontal="right", vertical="center")
        row[3].alignment = Alignment(horizontal="right", vertical="center")
        row[4].alignment = Alignment(horizontal="right", vertical="center")
        row[5].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        row[2].number_format = '#,##0.#####'
        row[3].number_format = '#,##0'
        row[4].number_format = '#,##0.00'

        refs = row[5].value or ""
        ws2.row_dimensions[row[0].row].height = auto_row_height(refs, chars_per_line=60)

    set_widths(ws2, {
        "A": 12,
        "B": 16,
        "C": 14,
        "D": 12,
        "E": 14,
        "F": 60,
    })

    # -----------------------------------------------------
    # HOJA RESERVAS_FANTASMA
    # -----------------------------------------------------
    ws3 = wb.create_sheet("RESERVAS_FANTASMA")

    headers3 = [
        "Tipo problema",
        "Nivel",
        "Empresa",
        "Location Code",
        "Bin Code",
        "Item No",
        "Descripción",
        "Cantidad",
        "Lote",
        "Fecha referencia",
        "Coste unitario",
        "Valor impacto",
        "Comentario",
    ]
    ws3.append(headers3)

    for r in ghost_rows:
        ws3.append([
            r["TipoProblema"],
            r["Nivel"],
            r["Empresa"],
            r["LocationCode"],
            r["BinCode"],
            r["ItemNo"],
            r["Descripcion"],
            float(r["Cantidad"]),
            r["LotNo"],
            r["FechaReferencia"],
            float(r["CosteUnitario"]),
            float(r["ValorImpacto"]),
            r["Comentario"],
        ])

    apply_base_style(ws3)

    for row in ws3.iter_rows(min_row=2):
        issue = safe_str(row[0].value)
        fill = get_fill_for_issue(issue)
        if fill:
            for cell in row:
                cell.fill = fill

        row[0].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        row[1].alignment = Alignment(horizontal="center", vertical="center")
        row[2].alignment = Alignment(horizontal="center", vertical="center")
        row[3].alignment = Alignment(horizontal="center", vertical="center")
        row[4].alignment = Alignment(horizontal="center", vertical="center")
        row[5].alignment = Alignment(horizontal="center", vertical="center")
        row[6].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        row[7].alignment = Alignment(horizontal="right", vertical="center")
        row[8].alignment = Alignment(horizontal="center", vertical="center")
        row[9].alignment = Alignment(horizontal="center", vertical="center")
        row[10].alignment = Alignment(horizontal="right", vertical="center")
        row[11].alignment = Alignment(horizontal="right", vertical="center")
        row[12].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        row[7].number_format = '#,##0.#####'
        row[9].number_format = 'yyyy-mm-dd'
        row[10].number_format = '#,##0.#####'
        row[11].number_format = '#,##0.00'

        txt = f"{row[6].value or ''} {row[12].value or ''}"
        ws3.row_dimensions[row[0].row].height = auto_row_height(txt, chars_per_line=50)

    set_widths(ws3, {
        "A": 28,
        "B": 10,
        "C": 12,
        "D": 12,
        "E": 16,
        "F": 16,
        "G": 38,
        "H": 12,
        "I": 14,
        "J": 14,
        "K": 14,
        "L": 14,
        "M": 40,
    })

    # -----------------------------------------------------
    # HOJA RESUMEN_FANTASMA
    # -----------------------------------------------------
    ws4 = wb.create_sheet("RESUMEN_FANTASMA")

    headers4 = [
        "Tipo problema",
        "Casos",
        "Impacto cantidad",
        "Impacto valor",
    ]
    ws4.append(headers4)

    for r in ghost_summary_rows:
        ws4.append([
            r["TipoProblema"],
            int(r["Casos"]),
            float(r["ImpactoCantidad"]),
            float(r["ImpactoValor"]),
        ])

    apply_base_style(ws4)

    for row in ws4.iter_rows(min_row=2):
        issue = safe_str(row[0].value)
        fill = get_fill_for_issue(issue)
        if fill:
            for cell in row:
                cell.fill = fill

        row[0].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        row[1].alignment = Alignment(horizontal="right", vertical="center")
        row[2].alignment = Alignment(horizontal="right", vertical="center")
        row[3].alignment = Alignment(horizontal="right", vertical="center")

        row[1].number_format = '#,##0'
        row[2].number_format = '#,##0.#####'
        row[3].number_format = '#,##0.00'

    set_widths(ws4, {
        "A": 28,
        "B": 12,
        "C": 16,
        "D": 16,
    })

    # -----------------------------------------------------
    # HOJA TOP_BINS_PROBLEMATICOS
    # -----------------------------------------------------
    ws5 = wb.create_sheet("TOP_BINS_PROBLEMATICOS")

    headers5 = [
        "Location Code",
        "Bin Code",
        "Casos",
        "Nº items",
        "Impacto cantidad",
        "Impacto valor",
        "Tipos problema",
    ]
    ws5.append(headers5)

    for r in top_bins_rows:
        ws5.append([
            r["LocationCode"],
            r["BinCode"],
            int(r["Casos"]),
            int(r["NumItems"]),
            float(r["ImpactoCantidad"]),
            float(r["ImpactoValor"]),
            r["TiposProblema"],
        ])

    apply_base_style(ws5)

    for row in ws5.iter_rows(min_row=2):
        row[0].alignment = Alignment(horizontal="center", vertical="center")
        row[1].alignment = Alignment(horizontal="center", vertical="center")
        row[2].alignment = Alignment(horizontal="right", vertical="center")
        row[3].alignment = Alignment(horizontal="right", vertical="center")
        row[4].alignment = Alignment(horizontal="right", vertical="center")
        row[5].alignment = Alignment(horizontal="right", vertical="center")
        row[6].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        row[2].number_format = '#,##0'
        row[3].number_format = '#,##0'
        row[4].number_format = '#,##0.#####'
        row[5].number_format = '#,##0.00'

        txt = row[6].value or ""
        ws5.row_dimensions[row[0].row].height = auto_row_height(txt, chars_per_line=45)

    set_widths(ws5, {
        "A": 12,
        "B": 16,
        "C": 10,
        "D": 10,
        "E": 16,
        "F": 16,
        "G": 35,
    })

    # -----------------------------------------------------
    # HOJA TOP_ITEMS_PROBLEMATICOS
    # -----------------------------------------------------
    ws6 = wb.create_sheet("TOP_ITEMS_PROBLEMATICOS")

    headers6 = [
        "Item No",
        "Descripción",
        "Casos",
        "Nº bins",
        "Impacto cantidad",
        "Impacto valor",
        "Tipos problema",
    ]
    ws6.append(headers6)

    for r in top_items_rows:
        ws6.append([
            r["ItemNo"],
            r["Descripcion"],
            int(r["Casos"]),
            int(r["NumBins"]),
            float(r["ImpactoCantidad"]),
            float(r["ImpactoValor"]),
            r["TiposProblema"],
        ])

    apply_base_style(ws6)

    for row in ws6.iter_rows(min_row=2):
        row[0].alignment = Alignment(horizontal="center", vertical="center")
        row[1].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        row[2].alignment = Alignment(horizontal="right", vertical="center")
        row[3].alignment = Alignment(horizontal="right", vertical="center")
        row[4].alignment = Alignment(horizontal="right", vertical="center")
        row[5].alignment = Alignment(horizontal="right", vertical="center")
        row[6].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        row[2].number_format = '#,##0'
        row[3].number_format = '#,##0'
        row[4].number_format = '#,##0.#####'
        row[5].number_format = '#,##0.00'

        txt = f"{row[1].value or ''} {row[6].value or ''}"
        ws6.row_dimensions[row[0].row].height = auto_row_height(txt, chars_per_line=45)

    set_widths(ws6, {
        "A": 16,
        "B": 42,
        "C": 10,
        "D": 10,
        "E": 16,
        "F": 16,
        "G": 35,
    })

    # -----------------------------------------------------
    # HOJA LEYENDA
    # -----------------------------------------------------
    ws7 = wb.create_sheet("LEYENDA")
    ws7.append(["Elemento", "Significado"])
    ws7.append(["Rojo suave en cantidades", "Cantidad negativa"])
    ws7.append(["Gris en cantidades", "Cantidad cero"])
    ws7.append([f"Amarillo en cantidades", f"Cantidad positiva pequeña (<= {QTY_SMALL_WARNING})"])
    ws7.append([f"Verde en cantidades", f"Cantidad alta (>= {QTY_HIGH_STOCK})"])
    ws7.append(["RESERVA_FANTASMA_CANDIDATA", "Cantidad negativa con lote y/o fecha. Caso muy sospechoso"])
    ws7.append(["NEGATIVO_ILE", "Cantidad negativa en ILE agrupado"])
    ws7.append(["CERO_CON_LOTE_O_FECHA", "Cantidad nula pero conserva lote o fecha. Revisar"])
    ws7.append(["Empresas excluidas", ", ".join(sorted(EXCLUDED_EMPRESA_VALUES)) if EXCLUDED_EMPRESA_VALUES else "(ninguna)"])

    apply_base_style(ws7)
    set_widths(ws7, {
        "A": 28,
        "B": 70,
    })

    wb.save(path)


# =========================================================
# MAIN
# =========================================================
def main():
    try:
        location_code, bin_code, item_no = ask_filters()

        detail_rows = fetch_detail_ile_rows(location_code, bin_code, item_no)
        ghost_rows = build_ghost_rows(detail_rows)
        summary_rows = build_summary_ubicaciones(detail_rows)
        ghost_summary_rows = build_summary_ghost(ghost_rows)
        top_bins_rows = build_top_bins_problematicos(ghost_rows)
        top_items_rows = build_top_items_problematicos(ghost_rows)

        if not detail_rows and not ghost_rows:
            show_info("Sin resultados", "No se han encontrado registros con esos filtros.")
            return

        ensure_dir(OUT_DIR)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        tag_parts = []
        if location_code:
            tag_parts.append(f"LOC-{location_code}")
        if bin_code:
            safe_bin = bin_code.replace("*", "").replace("%", "").replace("/", "-").replace("\\", "-")
            tag_parts.append(f"BIN-{safe_bin}")
        if item_no:
            tag_parts.append(f"ITEM-{item_no}")

        tag = "_".join(tag_parts) if tag_parts else "TODOS"

        out_path = os.path.join(OUT_DIR, f"Listado_ReservasFantasma_MEJORADO_{tag}_{ts}.xlsx")
        export_xlsx(
            out_path,
            detail_rows,
            summary_rows,
            ghost_rows,
            ghost_summary_rows,
            top_bins_rows,
            top_items_rows
        )

        total_valor = sum(float(r["ValorTotal"]) for r in summary_rows) if summary_rows else 0.0
        total_fantasma = sum(int(r["Casos"]) for r in ghost_summary_rows) if ghost_summary_rows else 0

        show_info(
            "Proceso finalizado",
            f"Archivo generado correctamente:\n\n{out_path}\n\n"
            f"Líneas DETALLE_ILE: {len(detail_rows)}\n"
            f"Ubicaciones resumen: {len(summary_rows)}\n"
            f"Casos reservas fantasma: {total_fantasma}\n"
            f"Tipos de problema: {len(ghost_summary_rows)}\n"
            f"Top bins problemáticos: {len(top_bins_rows)}\n"
            f"Top items problemáticos: {len(top_items_rows)}\n"
            f"Valor total agrupado: {format_decimal_es(total_valor, 2)}"
        )

    except Exception as e:
        show_error("Error", str(e))


if __name__ == "__main__":
    main()