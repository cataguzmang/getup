import datetime
from pathlib import Path

import parse_excel as pe
from conftest import make_xlsx, HEADER, line


def test_list_source_files_ignora_temporales(tmp_path):
    make_xlsx(tmp_path / "a.xlsx", [HEADER])
    make_xlsx(tmp_path / "b.xlsx", [HEADER])
    (tmp_path / "~$a.xlsx").write_text("lock")  # archivo de bloqueo de Excel
    files = pe.list_source_files(tmp_path)
    names = sorted(f.name for f in files)
    assert names == ["a.xlsx", "b.xlsx"]


def test_read_sheet_rows_prefiere_GOA_luego_Sheet1(tmp_path):
    dt = datetime.datetime
    p = make_xlsx(tmp_path / "s.xlsx",
                  [HEADER, line(dt(2026, 2, 3), "S1", "[GOA01] Green Tea", "Tienda", "Ana")],
                  sheet_name="Sheet1")
    rows = pe.read_sheet_rows(p)
    assert rows[0][0] == "Order Date"
    assert rows[1][1] == "S1"


def test_parse_transactions_agrega_mes():
    dt = datetime.datetime
    rows = [HEADER,
            line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana"),
            line(dt(2026, 5, 3), "S2", "[GOA02] Pure Chamoline", "Tienda B", "Luis")]
    lines = pe.parse_transactions(rows)
    assert lines[0]["month"] == "2026-02"
    assert lines[1]["month"] == "2026-05"


def test_month_label_en_espanol():
    assert pe.month_label("2026-02") == "Febrero 2026"
    assert pe.month_label("2026-05") == "Mayo 2026"


def test_load_all_sources_dedup_por_orden(tmp_path):
    dt = datetime.datetime
    # archivo 1: orden S1 (feb)
    make_xlsx(tmp_path / "1.xlsx", [HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana")])
    # archivo 2: repite S1 y agrega S2 (mar)
    make_xlsx(tmp_path / "2.xlsx", [HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana"),
        line(dt(2026, 3, 1), "S2", "[GOA02] Pure Chamoline", "Tienda B", "Luis")])
    lines, inc_by_month, sources, dups = pe.load_all_sources(tmp_path)
    orders = sorted({ln["order"] for ln in lines})
    assert orders == ["S1", "S2"]          # S1 no se cuenta dos veces
    assert len(lines) == 2
    assert "S1" in dups                      # se reporta el duplicado
    assert sorted(sources) == ["1.xlsx", "2.xlsx"]
