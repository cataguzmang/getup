"""
column_resolver.py — Resolución de columnas por nombre de encabezado.

El distribuidor cambia el layout de sus Excel sin aviso (columnas insertadas,
quitadas o reordenadas). Este módulo ubica cada campo lógico por el NOMBRE de su
encabezado en vez de por posición fija, con fallback posicional cuando no hay
encabezado reconocible. Lo comparten `report_core.py` y `split_excel.py`.

Caso real (2026-07-10): 'Sales Kombuchacha.xlsx' llegó con una columna Address
insertada y sin Company; el mapeo posicional leyó direcciones como vendedores.
"""

import re

# Campo lógico -> encabezado esperado (normalizado: minúsculas, espacios simples)
FIELD_HEADERS = {
    "date": "order date",
    "order": "order",
    "variant": "product variant",
    "customer": "customer",
    "salesperson": "salesperson",
    "company": "company",
    "qty_delivered": "qty delivered",
    "qty_invoiced": "qty invoiced",
    "qty_ordered": "qty ordered",
    "unit_price": "unit price",
    "total": "total",
}

# Posiciones del formato canónico (fallback cuando no hay encabezado)
CANONICAL_POSITIONS = {
    "date": 0, "order": 1, "variant": 2, "customer": 3, "salesperson": 4,
    "company": 5, "qty_delivered": 6, "qty_invoiced": 7, "qty_ordered": 8,
    "unit_price": 9, "total": 10,
}

# Columnas imprescindibles para reconocer una fila como encabezado
_REQUIRED = ("order date", "product variant")


def _norm(cell):
    """Normaliza una celda para comparar: str, trim, espacios simples, minúsculas."""
    if cell is None:
        return ""
    return re.sub(r"\s+", " ", str(cell)).strip().lower()


def find_header_row(rows, max_scan=10):
    """Índice de la primera fila (entre las primeras `max_scan`) que contiene los
    encabezados imprescindibles. None si no aparece."""
    for i, row in enumerate(rows[:max_scan]):
        normed = {_norm(c) for c in row}
        if all(r in normed for r in _REQUIRED):
            return i
    return None


def resolve_columns(header_row, label=""):
    """{campo_lógico: índice} para cada encabezado reconocido en la fila.

    Los campos no encontrados NO aparecen en el dict; por cada faltante se
    imprime un aviso (una vez por llamada). `label` identifica archivo/hoja."""
    normed = [_norm(c) for c in header_row]
    cols = {}
    for field, header in FIELD_HEADERS.items():
        if header in normed:
            cols[field] = normed.index(header)
        else:
            where = f"{label}: " if label else ""
            print(f"  ⚠ {where}columna '{header}' no encontrada en el encabezado; "
                  f"sus valores quedarán vacíos")
    return cols


def get(row, cols, field, default=None):
    """row[cols[field]] con tolerancia: campo sin resolver o fila corta -> default."""
    i = cols.get(field)
    if i is None or i >= len(row):
        return default
    return row[i]
