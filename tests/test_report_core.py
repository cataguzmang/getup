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


HEADER_ADDR = ["Order Date", "Order", "Product Variant", "Customer", "Address",
               "Salesperson", "Qty Delivered", "Qty Invoiced", "Qty Ordered",
               "Unit Price", "Total"]


def _line_addr(date, order, variant, customer, address, sp,
               qty=1, price=21.85, total=21.85):
    """Fila con el layout nuevo: Address insertada, sin Company."""
    return [date, order, variant, customer, address, sp, qty, qty, qty, price, total]


def test_layout_con_address_lee_vendedor_correcto(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "Sales Kombuchacha.xlsx", [HEADER_ADDR,
        _line_addr(dt(2026, 5, 12), "S42061", "[KOM01] Kombuchacha Zero Blueberry",
                   "Tropical Sprmkt - Dunellen", "446 North Ave, Dunellen, NJ 08812",
                   "Mireya Fernandez", qty=0.5, price=50.64, total=25.32),
    ])
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    names = [s["name"] for s in rep["salespeople"]]
    assert names == ["Mireya Fernandez"]        # persona, NO la dirección
    assert rep["totals"]["revenue"] == 25.32


def test_mezcla_canonico_y_layout_nuevo(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    # archivo canónico (11 columnas estándar)
    _xlsx(fuentes / "KOM-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
              "Convenience store flora llc", "Mireya Fernandez",
              qty=0.5, price=50.64, total=25.32),
    ], sheet="KOM")
    # archivo con layout nuevo (Address, sin Company)
    _xlsx(fuentes / "Sales Kombuchacha.xlsx", [HEADER_ADDR,
        _line_addr(dt(2026, 5, 4), "S41495", "[KOM01] Kombuchacha Zero Blueberry",
                   "Castellanos Grocery", "2 Orchard St, Morristown, NJ 07860",
                   "Katherine Osorio Duque", qty=1, price=54.64, total=54.64),
    ])
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    names = sorted(s["name"] for s in rep["salespeople"])
    assert names == ["Katherine Osorio Duque", "Mireya Fernandez"]
    assert rep["totals"]["revenue"] == 79.96      # 25.32 + 54.64
    assert len(rep["months"]) == 2                 # 2026-05 y 2026-06


def test_mes_en_cero_desde_archivo_solo_encabezado(tmp_path):
    """D3: un canónico mensual sin líneas (KOM julio: 'No sales reported') registra
    el mes vacío explícito — "vendió $0" y "no cargué el mes" deben verse distinto."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "KOM-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
              "Convenience store flora llc", "Mireya Fernandez",
              qty=0.5, price=50.64, total=25.32),
    ], sheet="KOM")
    _xlsx(fuentes / "KOM-2026-07.xlsx", [HEADER], sheet="KOM")   # solo encabezado
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    assert [m["key"] for m in rep["months"]] == ["2026-06", "2026-07"]
    jul = rep["months"][1]
    assert jul["empty"] is True
    assert jul["totals"]["units"] == 0 and jul["totals"]["revenue"] == 0
    assert jul["products"] == [] and jul["salespeople"] == []
    assert jul["periodStart"] == "2026-07-01" and jul["periodEnd"] == "2026-07-31"
    assert rep["meta"]["periodLabel"] == "Junio 2026 – Julio 2026"
    # el mes vacío no toca la vista general
    assert rep["totals"]["revenue"] == 25.32


def test_archivo_sin_mes_en_nombre_no_crea_mes_vacio(tmp_path):
    """Un fuente sin '-AAAA-MM' en el nombre (ej. el crudo del distribuidor) que no
    aporta líneas NO inventa un mes."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "KOM-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 18), "S1", "[KOM01] Zero Blueberry", "T", "Mireya",
              qty=0.5, price=50.64, total=25.32),
    ], sheet="KOM")
    _xlsx(fuentes / "Sales Kombuchacha.xlsx", [HEADER], sheet="KOM")
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    assert [m["key"] for m in rep["months"]] == ["2026-06"]


def test_mes_con_lineas_no_se_marca_vacio(tmp_path):
    """Si el mes del nombre de archivo recibe líneas desde OTRO fuente, no es mes
    en cero (y no gana la marca `empty`)."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "KOM-2026-06.xlsx", [HEADER], sheet="KOM")   # vacío
    _xlsx(fuentes / "Sales Kombuchacha.xlsx", [HEADER,
        _line(dt(2026, 6, 4), "S2", "[KOM01] Zero Blueberry", "T", "Katherine",
              qty=1, price=54.64, total=54.64),
    ], sheet="KOM")
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    assert [m["key"] for m in rep["months"]] == ["2026-06"]
    assert "empty" not in rep["months"][0]


def test_sin_encabezado_usa_posiciones_canonicas(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    # sin fila de encabezado: debe caer al mapeo posicional canónico
    _xlsx(fuentes / "raw.xlsx", [
        _line(dt(2026, 6, 1), "S1", "[KOM01] Kombuchacha Zero Blueberry",
              "Tienda", "Mireya Fernandez", qty=1, price=50.64, total=50.64),
    ])
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    assert rep["salespeople"][0]["name"] == "Mireya Fernandez"
    assert rep["totals"]["revenue"] == 50.64
