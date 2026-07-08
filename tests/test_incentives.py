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
