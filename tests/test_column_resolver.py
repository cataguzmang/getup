import column_resolver as cr

HEADER = ["Order Date", "Order", "Product Variant", "Customer", "Salesperson",
          "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total"]
# Layout real recibido el 2026-07-10: Address insertada, sin Company
HEADER_ADDR = ["Order Date", "Order", "Product Variant", "Customer", "Address",
               "Salesperson", "Qty Delivered", "Qty Invoiced", "Qty Ordered",
               "Unit Price", "Total"]


def test_find_header_row_primera_fila():
    assert cr.find_header_row([HEADER, ["x"] * 11]) == 0


def test_find_header_row_en_medio():
    rows = [["basura"], [None, None], HEADER, ["x"] * 11]
    assert cr.find_header_row(rows) == 2


def test_find_header_row_ausente():
    assert cr.find_header_row([["a", "b", "c"], [1, 2, 3]]) is None


def test_find_header_row_respeta_max_scan():
    rows = [["x"]] * 12 + [HEADER]
    assert cr.find_header_row(rows, max_scan=10) is None


def test_resolve_canonico_mapea_identico():
    assert cr.resolve_columns(HEADER) == cr.CANONICAL_POSITIONS


def test_resolve_con_address_insertada_y_sin_company(capsys):
    cols = cr.resolve_columns(HEADER_ADDR, label="Sales Kombuchacha.xlsx")
    assert cols["customer"] == 3
    assert cols["salesperson"] == 5      # corrida por Address
    assert cols["qty_delivered"] == 6
    assert "company" not in cols          # ausente -> no aparece
    out = capsys.readouterr().out
    assert "company" in out.lower() and "Sales Kombuchacha.xlsx" in out


def test_resolve_normaliza_mayusculas_y_espacios():
    header = ["  order DATE ", "ORDER", "product   variant", "Customer",
              "SalesPerson", "company", "qty delivered", "qty invoiced",
              "qty ordered", "unit price", "total"]
    cols = cr.resolve_columns(header)
    assert cols == cr.CANONICAL_POSITIONS


def test_get_tolerante():
    row = ["2026-06-01", "S1", "[KOM01] x"]
    cols = {"date": 0, "variant": 2, "total": 10}
    assert cr.get(row, cols, "date") == "2026-06-01"
    assert cr.get(row, cols, "total") is None            # índice fuera de rango
    assert cr.get(row, cols, "company", "") == ""        # campo no resuelto
