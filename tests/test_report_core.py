import datetime
import openpyxl

import report_core as rc

HEADER = ["Order Date", "Order", "Product Variant", "Customer", "Salesperson",
          "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total"]
dt = datetime.datetime


def _xlsx(path, rows, sheet="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def _line(date, order, variant, customer, sp, qty=1, price=21.85, total=21.85,
          company="LatinFood US Corp."):
    return [date, order, variant, customer, sp, company, qty, qty, qty, price, total]


def test_sku_pattern_por_prefijo():
    pk = rc.sku_pattern("KOM")
    assert rc.split_sku("[KOM01] Kombuchacha Zero Blueberry", pk) == \
        ("KOM01", "Kombuchacha Zero Blueberry")
    # un prefijo distinto no matchea con el patrón de KOM
    assert rc.split_sku("[ROB05] Mussel en Brine", pk)[0] is None


def test_parse_transactions_preserva_cantidad_fraccionaria():
    pk = rc.sku_pattern("KOM")
    rows = [HEADER, _line(dt(2026, 6, 18), "S1", "[KOM01] Zero Blueberry",
                          "Tienda", "Mireya", qty=0.5, price=50.64, total=25.32)]
    lines = rc.parse_transactions(rows, pk)
    assert lines[0]["qtyOrdered"] == 0.5     # NO truncado a 0
    assert lines[0]["total"] == 25.32


def test_read_sheet_rows_respeta_sheet_candidates(tmp_path):
    p = _xlsx(tmp_path / "s.xlsx",
              [HEADER, _line(dt(2026, 6, 1), "S1", "[KOM01] X", "T", "A", qty=0.5)],
              sheet="KOM")
    rows = rc.read_sheet_rows(p, ("KOM", "Sheet1"))
    assert rows[0][0] == "Order Date" and rows[1][1] == "S1"


def test_build_report_kom_shape(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "KOM-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
              "Convenience store flora llc", "Mireya Fernandez", qty=0.5, price=50.64, total=25.32),
        _line(dt(2026, 6, 18), "S44804", "[KOM02] Kombuchacha Zero Raspberry",
              "Convenience store flora llc", "Mireya Fernandez", qty=0.5, price=50.64, total=25.32),
    ], sheet="KOM")
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp", "KOM", ("KOM", "Sheet1"))
    assert len(rep["products"]) == 2
    assert rep["totals"]["units"] == 1.0
    assert rep["totals"]["revenue"] == 50.64
    assert rep["totals"]["customers"] == 1
    assert len(rep["salespeople"]) == 1
    assert rep["meta"]["brand"] == "Kombuchacha"
    assert len(rep["months"]) == 1


def test_build_report_rob_shape(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "ROB-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 1), "S43486", "[ROB05] Mussel en Brine /Mejillones",
              "Gala foods", "Luis Soler", qty=1, price=60, total=60),
    ], sheet="ROB")
    rep = rc.build_report(tmp_path, "Robinson Crusoe", "LatinFood US Corp", "ROB", ("ROB", "Sheet1"))
    assert len(rep["products"]) == 1
    assert rep["totals"]["revenue"] == 60.0
    assert rep["totals"]["units"] == 1
    assert rep["meta"]["distributor"] == "LatinFood US Corp"
