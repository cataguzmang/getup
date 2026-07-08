import sys
from pathlib import Path

# Hacer importable split_excel.py (vive en la raíz, padre de tests/)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest

from helpers import HEADER, tx, subtotal, make_workbook, D


@pytest.fixture
def messy_wb(tmp_path):
    """Workbook sucio de ejemplo: 4 marcas (GOA, GET, KOM, XYZ desconocida),
    dos estados, filas-subtotal y un bloque de incentivos en la hoja de MIA."""
    sheets = {
        "NY Jun - GOA": [
            HEADER,
            subtotal("Katherine Osorio Duque (6)", qd=5, qi=5, qo=5, price=25.8, total=109.25),
            tx(D(2026, 6, 24), "S45227", "[GOA02] Pure Chamoline Herbal Tea",
               "Food Bazaar # 14 -West NY", "Katherine Osorio Duque"),
            tx(D(2026, 6, 17), "S44722", "[GOA03] Ginger Lemongrass Herbal Tea",
               "Food Bazaar # 14 -West NY", "Katherine Osorio Duque"),
        ],
        "MIA Jun - Jack Mckarel": [
            HEADER,
            subtotal("[GET01] Jack Mackerel in Brine (8)", qd=68, qi=68, qo=68, price=39.9, total=2964),
            subtotal("    Alexandra J (2)", qd=35, qi=35, qo=35, price=45.6, total=1596),
            tx(D(2026, 6, 22), "S44946", "[GET01] Jack Mackerel in Brine",
               "Key Food 10400", "Alexandra J", company="LatinFood Florida",
               qd=20, qi=20, qo=20, price=45.6, total=912),
            # bloque de incentivos (col A con fecha de mayo, col B == 'SOLD')
            [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None],
            ["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68,
             None, None, None, None, None, None],
        ],
        "NY Jun - KOM": [
            HEADER,
            tx(D(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
               "Convenience store flora llc", "Mireya Fernandez",
               qd=0.5, qi=0.5, qo=0.5, price=50.64, total=25.32),
        ],
        "NY Jun - XYZ": [
            HEADER,
            tx(D(2026, 6, 10), "S99999", "[XYZ01] Mystery Product", "Some Store", "Someone"),
        ],
    }
    return make_workbook(tmp_path / "GET Jun 2026.xlsx", sheets)
