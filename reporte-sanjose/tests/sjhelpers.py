"""Constructores de workbooks sintéticos para los tests de San José."""
import datetime as _dt

import openpyxl

D = _dt.datetime

# Esquema viejo del Historico (13 columnas)
HIST_HEADER = [
    "Company", "State", "Customer", "Order", "Order Date", "Cod Prod", "Product",
    "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Salesperson", "Total", "Unit Price",
]

# Formato canónico (11 columnas)
CANON_HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson", "Company",
    "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]


def _mkwb(path, header, rows, sheet):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(header)
    for r in rows:
        ws.append(list(r) + [None] * (len(header) - len(r)))
    wb.save(path)
    return path


def make_historico(path, rows, sheet="Historico"):
    return _mkwb(path, HIST_HEADER, rows, sheet)


def make_canonical(path, rows, sheet="San Jose"):
    return _mkwb(path, CANON_HEADER, rows, sheet)


def hist_row(company, state, customer, order, date, cod, product,
             qd, salesperson, total, unit_price, qi=None, qo=None):
    qi = qd if qi is None else qi
    qo = qd if qo is None else qo
    return [company, state, customer, order, date, cod, product,
            qd, qi, qo, salesperson, total, unit_price]


def canon_row(date, order, variant, customer, salesperson, company,
              qd, unit_price, total, qi=None, qo=None):
    qi = qd if qi is None else qi
    qo = qd if qo is None else qo
    return [date, order, variant, customer, salesperson, company,
            qd, qi, qo, unit_price, total]
