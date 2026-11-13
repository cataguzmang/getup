import datetime

import split_excel as sx

D = datetime.datetime


def _rows_with_block():
    return [
        list(sx.HEADER),
        [D(2026, 6, 22), "S44946", "[GET01] x", "c", "Alexandra J",
         "LatinFood Florida", 20, 20, 20, 45.6, 912],
        [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None],
        ["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68,
         None, None, None, None, None, None],
    ]


def test_find_incentive_block_returns_from_sold():
    block = sx.find_incentive_block(_rows_with_block())
    assert len(block) == 2
    assert block[0][1] == "SOLD"
    assert block[1][0] == "Alexandra Jimenez"


def test_find_incentive_block_absent():
    rows = [
        list(sx.HEADER),
        [D(2026, 6, 22), "S44946", "[GET01] x", "c", "sp",
         "LatinFood Florida", 20, 20, 20, 45.6, 912],
    ]
    assert sx.find_incentive_block(rows) == []


def test_find_incentive_block_layout_goa_oct_2026():
    """Desde oct-2026 el bloque de GOA no trae 'SOLD': encabezado Vendedor/Gratis."""
    rows = [
        list(sx.HEADER),
        [D(2026, 10, 9), "S53248", "[GOA01] x", "c", "Katherine Osorio Duque",
         "LatinFood US Corp.", 1, 1, 1, 21.85, 21.85],
        [None, None, None, None, None, "Vendedor", "Punto de venta Nvo", "Vendidos",
         "Gratis", "Bono x Punto Nvo", None],
        [None, None, None, None, None, "Katherine Osorio", 2, 23, 4, 20, None],
    ]
    block = sx.find_incentive_block(rows)
    assert len(block) == 2
    assert block[0][5] == "Vendedor"


def test_find_incentive_block_vendedor_sin_columnas_de_bloque_no_es_ancla():
    rows = [list(sx.HEADER), [None, None, None, "Vendedor", None, None, None, None, None, None, None]]
    assert sx.find_incentive_block(rows) == []
