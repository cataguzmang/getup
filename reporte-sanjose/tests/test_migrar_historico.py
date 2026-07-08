import openpyxl

from sjhelpers import make_historico, hist_row, D, CANON_HEADER
import migrar_historico as mig


def test_migrate_reorders_to_canonical(tmp_path):
    src = make_historico(tmp_path / "hist.xlsx", [
        hist_row("LatinFood Florida", "Florida", "Cliente A", "S100",
                 D(2025, 10, 9), "[GET01]", "Jack Mackerel in Brine",
                 4, "Alexandra J", 182.4, 45.6),
        hist_row("LatinFood US Corp", "Nueva York", "Cliente B", "S200",
                 D(2026, 1, 5), "[GET02]", "Jack Mackerel in Tomato Sauce",
                 1, "Katherine", 0, 0),  # promo
    ])
    out = tmp_path / "fuentes" / "San-Jose-historico.xlsx"
    n = mig.migrate(src=src, out=out)
    assert n == 0  # ninguna fila sin Cod Prod

    ws = openpyxl.load_workbook(out).active
    assert ws.title == "San Jose"
    grid = list(ws.iter_rows(values_only=True))
    assert list(grid[0]) == CANON_HEADER
    # fila 1: Product Variant reconstruido, Company preservado, sin columna State
    r = grid[1]
    assert r[0] == D(2025, 10, 9)                 # Order Date
    assert r[1] == "S100"                          # Order
    assert r[2] == "[GET01] Jack Mackerel in Brine"  # Product Variant
    assert r[3] == "Cliente A"                      # Customer
    assert r[4] == "Alexandra J"                    # Salesperson
    assert r[5] == "LatinFood Florida"             # Company
    assert r[9] == 45.6                             # Unit Price
    assert r[10] == 182.4                           # Total
    assert len(grid) == 3                           # header + 2 filas


def test_migrate_counts_missing_cod(tmp_path):
    src = make_historico(tmp_path / "hist.xlsx", [
        hist_row("LatinFood Florida", "Florida", "C", "S1",
                 D(2025, 10, 9), None, "Producto sin cod",
                 1, "SP", 45.6, 45.6),
    ])
    out = tmp_path / "fuentes" / "h.xlsx"
    assert mig.migrate(src=src, out=out) == 1
