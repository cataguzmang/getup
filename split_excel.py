"""
split_excel.py — Dispatcher de reparto por marca (SP1)

Lee el/los Excel "sucios" del distribuidor en entrada/ (una hoja por estado+marca),
separa cada fila de pedido por marca según el prefijo del SKU, y deposita un .xlsx
canónico y limpio en fuentes/ de cada marca. Luego regenera el data.js de las marcas
cuyo generador ya entiende el formato canónico (por ahora: GOA).

Uso:
    python split_excel.py [--input entrada] [--no-build] [--only GOA,KOM]
"""

import argparse
import datetime
import re
import subprocess
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent


# ── Configuración de marcas ──────────────────────────────────────────────
@dataclass(frozen=True)
class Brand:
    name: str        # nombre legible
    code: str        # prefijo del nombre de archivo canónico
    folder: str      # carpeta destino (relativa a ROOT)
    generator: str   # script generador dentro de la carpeta
    sheet: str       # nombre de la hoja del archivo canónico
    sp1_ready: bool  # ¿su generador ya lee el formato canónico?


BRANDS = {
    "GET": Brand("San José", "San-Jose", "reporte-sanjose", "generar_data.py", "San Jose", False),
    "GOA": Brand("Garden of the Andes", "GOA", "goa-inventory", "parse_excel.py", "GOA", True),
    "KOM": Brand("Kombuchacha", "KOM", "kombuchacha", "build_data.py", "KOM", False),
    "ROB": Brand("Robinson Crusoe", "ROB", "robinson-crusoe", "build_data.py", "ROB", False),
}

HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson",
    "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]

# Índices de columna (0-based)
COL_DATE, COL_ORDER, COL_VARIANT, COL_CUSTOMER, COL_SP, COL_COMPANY = 0, 1, 2, 3, 4, 5

SKU_PREFIX_RE = re.compile(r"^\s*\[([A-Z]{2,4})\d+\]")

STATE_BY_COMPANY = {
    "latinfood florida": "Florida",
    "latinfood us corp.": "Nueva York",
    "latinfood us corp": "Nueva York",
}
STATE_BY_SHEET_PREFIX = {"MIA": "Florida", "NY": "Nueva York"}


# ── Helpers puros ─────────────────────────────────────────────────────────
def sku_prefix(variant):
    """'[GOA02] Pure Chamoline' -> 'GOA'. None si no matchea o no es str."""
    if not isinstance(variant, str):
        return None
    m = SKU_PREFIX_RE.match(variant)
    return m.group(1) if m else None


def is_transaction_row(row):
    """Fila de pedido real: datetime en col A y SKU [XXX##] en col C."""
    return (
        isinstance(row[COL_DATE], datetime.datetime)
        and sku_prefix(row[COL_VARIANT]) is not None
    )


def month_key(date):
    """datetime -> 'YYYY-MM'."""
    return f"{date.year:04d}-{date.month:02d}"


def state_of(row, sheet_name):
    """Estado desde Company; respaldo por prefijo de hoja; '?' si desconocido."""
    company = row[COL_COMPANY]
    if isinstance(company, str):
        key = re.sub(r"\s+", " ", company).strip().lower()
        if key in STATE_BY_COMPANY:
            return STATE_BY_COMPANY[key]
    parts = str(sheet_name).split()
    prefix = parts[0].upper() if parts else ""
    return STATE_BY_SHEET_PREFIX.get(prefix, "?")


def find_incentive_block(rows):
    """Filas del bloque de incentivos (desde el ancla 'SOLD' al final), o [].

    Ancla: alguna celda de la fila == 'SOLD' (case-insensitive).
    """
    for i, row in enumerate(rows):
        if any(isinstance(c, str) and c.strip().upper() == "SOLD" for c in row):
            return [list(r) for r in rows[i:]]
    return []


@dataclass
class Collected:
    rows: dict         # code -> {month -> [row, ...]}
    incentives: dict   # code -> {month -> [block_row, ...]}
    unmatched: list    # [row, ...]
    stats: dict        # code -> {month -> Counter(state)}


def collect(input_files):
    """Recorre los workbooks y separa filas por marca/mes, guardando además los
    SKUs no mapeados, los bloques de incentivos y los conteos por estado."""
    rows = defaultdict(lambda: defaultdict(list))
    incentives = defaultdict(dict)
    unmatched = []
    stats = defaultdict(lambda: defaultdict(Counter))

    for path in input_files:
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:  # archivo corrupto / no es xlsx
            print(f"  ⚠ No se pudo leer {Path(path).name}: {e} (se omite)")
            continue

        for ws in wb.worksheets:
            sheet_rows = list(ws.iter_rows(values_only=True))
            tx_codes, tx_months = [], []

            for row in sheet_rows:
                if not is_transaction_row(row):
                    continue
                prefix = sku_prefix(row[COL_VARIANT])
                mk = month_key(row[COL_DATE])
                if prefix not in BRANDS:
                    unmatched.append(list(row))
                    continue
                rows[prefix][mk].append(list(row))
                stats[prefix][mk][state_of(row, ws.title)] += 1
                tx_codes.append(prefix)
                tx_months.append(mk)

            block = find_incentive_block(sheet_rows)
            if block and tx_codes:
                dom_code = Counter(tx_codes).most_common(1)[0][0]
                dom_month = Counter(tx_months).most_common(1)[0][0]
                prev = incentives[dom_code].get(dom_month, [])
                incentives[dom_code][dom_month] = prev + block
            elif block:
                print(f"  ⚠ Bloque de incentivos en '{ws.title}' sin filas de pedido; se omite")

    # Convertir defaultdicts a dicts normales para asserts predecibles
    return Collected(
        rows={c: dict(m) for c, m in rows.items()},
        incentives={c: dict(m) for c, m in incentives.items()},
        unmatched=unmatched,
        stats={c: {mk: cnt for mk, cnt in m.items()} for c, m in stats.items()},
    )
