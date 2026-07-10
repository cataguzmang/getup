import openpyxl

import split_excel as sx
from helpers import D, make_workbook

HEADER_ADDR = ["Order Date", "Order", "Product Variant", "Customer", "Address",
               "Salesperson", "Qty Delivered", "Qty Invoiced", "Qty Ordered",
               "Unit Price", "Total"]


def _wb_addr(tmp_path):
    """Workbook de entrada con el layout nuevo (Address insertada, sin Company)."""
    sheets = {
        "NY Jun - KOM": [
            HEADER_ADDR,
            [D(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
             "Convenience store flora llc", "132-134 Bloomfield Ave, Newark, NJ",
             "Mireya Fernandez", 0.5, 0.5, 0.5, 50.64, 25.32],
        ],
    }
    return make_workbook(tmp_path / "GET Jun 2026.xlsx", sheets)


def test_collect_reordena_a_canonico(tmp_path):
    data = sx.collect([_wb_addr(tmp_path)])
    row = data.rows["KOM"]["2026-06"][0]
    assert row[sx.COL_SP] == "Mireya Fernandez"     # persona en la col canónica 4
    assert row[sx.COL_COMPANY] is None              # Company ausente -> vacía
    assert row[sx.COL_CUSTOMER] == "Convenience store flora llc"
    assert len(row) == len(sx.HEADER)               # siempre 11 columnas


def test_estado_cae_al_prefijo_de_hoja_sin_company(tmp_path):
    data = sx.collect([_wb_addr(tmp_path)])
    # sin Company, el estado sale del prefijo de la hoja ('NY ...')
    assert data.stats["KOM"]["2026-06"]["Nueva York"] == 1


def test_write_canonical_desde_layout_nuevo(tmp_path):
    data = sx.collect([_wb_addr(tmp_path)])
    out = tmp_path / "KOM-2026-06.xlsx"
    sx.write_canonical(out, data.rows["KOM"]["2026-06"], [], "KOM")
    ws = openpyxl.load_workbook(out)["KOM"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows[0] == tuple(sx.HEADER)
    assert rows[1][4] == "Mireya Fernandez"         # Salesperson en su lugar
    assert rows[1][6] == 0.5                        # Qty Delivered en su lugar


def test_canonico_sin_cambios_sigue_identico(messy_wb):
    """El workbook canónico de siempre produce exactamente lo mismo (regresión)."""
    data = sx.collect([messy_wb])
    row = data.rows["KOM"]["2026-06"][0]
    assert row[sx.COL_SP] == "Mireya Fernandez"
    assert row[sx.COL_COMPANY] == "LatinFood US Corp."
