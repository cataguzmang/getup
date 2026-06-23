import datetime
import sys
from pathlib import Path

# Hacer importable parse_excel.py (vive en el directorio padre de tests/)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
import pytest


def make_xlsx(path: Path, rows, sheet_name="GOA"):
    """Crea un .xlsx con `rows` (lista de listas) en la hoja indicada."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson",
    "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]


def line(date, order, sku_name, customer, sp, qty=1, price=21.85, total=21.85):
    return [date, order, sku_name, customer, sp, "LatinFood US Corp",
            qty, qty, qty, price, total]


@pytest.fixture
def fixtures():
    return {"make_xlsx": make_xlsx, "HEADER": HEADER, "line": line, "dt": datetime.datetime}
