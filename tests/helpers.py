"""Constructores de workbooks sintéticos para los tests de split_excel."""
import datetime as _dt

import openpyxl

D = _dt.datetime

HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson",
    "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]


def _pad(row):
    """Rellena la fila a 11 columnas con None (como hace openpyxl al leer)."""
    assert len(row) <= len(HEADER), f"fila con {len(row)} columnas (máx {len(HEADER)}): {row!r}"
    return list(row) + [None] * (len(HEADER) - len(row))


def tx(date, order, variant, customer, sp, company="LatinFood US Corp.",
       qd=1, qi=1, qo=1, price=21.85, total=21.85):
    """Fila de pedido real."""
    return [date, order, variant, customer, sp, company, qd, qi, qo, price, total]


def subtotal(label, qd=1, qi=1, qo=1, price=21.85, total=21.85):
    """Fila-subtotal (texto en col A, sin fecha)."""
    return [label, None, None, None, None, None, qd, qi, qo, price, total]


def make_workbook(path, sheets):
    """`sheets`: dict {nombre_hoja: [fila, ...]}. Devuelve `path`."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        for r in rows:
            ws.append(_pad(r))
    wb.save(path)
    return path
