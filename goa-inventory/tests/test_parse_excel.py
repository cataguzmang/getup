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


def test_build_view_totales_e_incentivos():
    dt = datetime.datetime
    lines = pe.parse_transactions([HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana", qty=2, total=43.70),
        line(dt(2026, 2, 25), "S2", "[GOA01] Green Tea", "Tienda B", "Ana", qty=1, total=21.85),
        line(dt(2026, 2, 25), "S3", "[GOA02] Pure Chamoline", "Tienda A", "Luis", qty=1, total=21.85)])
    inc = {"freeUnitsBySalesperson": [{"name": "Ana", "free": 5}],
           "costItems": [], "totalDescuentos": 70.95}
    v = pe.build_view(lines, inc)
    assert v["totals"]["units"] == 4
    assert v["totals"]["revenue"] == 87.40
    assert v["totals"]["orders"] == 3
    assert v["totals"]["customers"] == 2          # Tienda A y Tienda B
    ana = next(s for s in v["salespeople"] if s["name"] == "Ana")
    assert ana["freeUnits"] == 5                   # inyectado desde incentivos


def test_build_months_separa_por_mes():
    dt = datetime.datetime
    lines = pe.parse_transactions([HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana"),
        line(dt(2026, 3, 1), "S2", "[GOA02] Pure Chamoline", "Tienda B", "Luis"),
        line(dt(2026, 3, 2), "S3", "[GOA01] Green Tea", "Tienda C", "Luis")])
    months = pe.build_months(lines, {})
    keys = [m["key"] for m in months]
    assert keys == ["2026-02", "2026-03"]                 # cronológico
    assert months[0]["label"] == "Febrero 2026"
    assert months[1]["totals"]["orders"] == 2             # S2 y S3
    assert months[0]["periodStart"] == "2026-02-24"


def test_merge_views_suma_y_une():
    dt = datetime.datetime
    feb = pe.build_view(pe.parse_transactions([HEADER,
        line(dt(2026, 2, 1), "S1", "[GOA01] Green Tea", "Tienda A", "Ana")]),
        pe.empty_incentives())
    mar = pe.build_view(pe.parse_transactions([HEADER,
        line(dt(2026, 3, 1), "S2", "[GOA01] Green Tea", "Tienda A", "Ana"),
        line(dt(2026, 3, 2), "S3", "[GOA02] Pure Chamoline", "Tienda B", "Luis")]),
        pe.empty_incentives())
    g = pe.merge_views([feb, mar])
    assert g["totals"]["units"] == 3
    assert g["totals"]["orders"] == 3                 # suma (órdenes únicas por mes)
    assert g["totals"]["customers"] == 2              # unión: Tienda A (feb+mar) + Tienda B
    green = next(p for p in g["products"] if p["sku"] == "GOA01")
    assert green["units"] == 2                         # GOA01 en feb y mar
    assert len(g["daily"]) == 3                        # concatenado


def test_carry_forward_conserva_mes_ausente(tmp_path):
    prev = {"months": [
        {"key": "2026-01", "label": "Enero 2026", "periodStart": "2026-01-05",
         "periodEnd": "2026-01-20", "totals": {"units": 9}, "products": [],
         "salespeople": [], "customers": [], "daily": [], "incentives": pe.empty_incentives()},
        {"key": "2026-02", "label": "Febrero 2026", "totals": {"units": 1}, "products": [],
         "salespeople": [], "customers": [], "daily": [], "incentives": pe.empty_incentives()},
    ]}
    months = [{"key": "2026-02", "label": "Febrero 2026", "totals": {"units": 1},
               "products": [], "salespeople": [], "customers": [], "daily": [],
               "incentives": pe.empty_incentives()}]
    merged, carried = pe.carry_forward(months, prev)
    keys = [m["key"] for m in merged]
    assert keys == ["2026-01", "2026-02"]    # enero conservado, ordenado
    assert carried == ["2026-01"]


def test_load_previous_report_parsea_datajs(tmp_path):
    import json
    p = tmp_path / "data.js"
    p.write_text("const reportData = " + json.dumps({"months": []}) + ";\n", encoding="utf-8")
    rep = pe.load_previous_report(p)
    assert rep == {"months": []}
    assert pe.load_previous_report(tmp_path / "noexiste.js") is None
