import datetime

import openpyxl
import pytest

import parse_altagama as pa


# ── Fixture: mini-Excel con la misma forma que el archivo de Alta Gama ───────
HEADER = ["Products",
          datetime.datetime(2026, 1, 26), datetime.datetime(2026, 2, 26),
          "Jul 1 - 25, 26"]

ROWS = [
    HEADER,
    ["GA Green Tea 1.4oz (Garden of the Andes - Green Tea 1.4oz)", 10, 20, 5],
    ["GA Chai Tea 1.4oz (Garden of the Andes - Chai Tea 1.4oz)", 0, 4, 1],
    ["TOTAL", 10, 24, 6],
]


def make_sheet(path, rows, sheet_name="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


# ── Encabezados de periodo ──────────────────────────────────────────────────
def test_fecha_de_excel_es_mes_completo_y_el_26_es_el_anio():
    p = pa.parse_period_header(datetime.datetime(2026, 1, 26))
    assert p["key"] == "2026-01"
    assert p["label"] == "Enero 2026"
    assert (p["start"], p["end"]) == ("2026-01-01", "2026-01-31")
    assert p["partial"] is False


def test_texto_de_periodo_parcial():
    p = pa.parse_period_header("Jul 1 - 25, 26")
    assert p["key"] == "2026-07"
    assert p["partial"] is True
    assert (p["start"], p["end"]) == ("2026-07-01", "2026-07-25")
    assert (p["coverageDays"], p["monthDays"]) == (25, 31)


def test_mes_completo_escrito_como_texto():
    assert pa.parse_period_header("Jan-26")["key"] == "2026-01"
    assert pa.parse_period_header("Mar 2026")["key"] == "2026-03"


def test_encabezado_no_reconocido_falla():
    with pytest.raises(ValueError):
        pa.parse_period_header("Q1")


# ── Productos ───────────────────────────────────────────────────────────────
def test_split_product_separa_sku_y_nombre():
    sku, name = pa.split_product(
        "GA Roseship & Hibiscus 1.8oz (Garden of the Andes - Roseship and Hibiscus 1.8oz)")
    assert sku == "GA Roseship & Hibiscus 1.8oz"
    assert name == "Garden of the Andes - Roseship and Hibiscus 1.8oz"


def test_split_product_sin_parentesis_no_inventa_sku():
    sku, name = pa.split_product("  Producto   raro  ")
    assert sku is None
    assert name == "Producto raro"


def test_short_name_recorta_la_marca():
    assert pa.short_name("Garden of the Andes - Green Tea 1.4oz", "GA X") == "Green Tea 1.4oz"


# ── Celdas: cero real vs dato ausente ───────────────────────────────────────
def test_cero_real_se_conserva_como_cero():
    assert pa.units(0) == (0, None)


def test_celda_vacia_es_no_disponible_no_cero():
    value, reason = pa.units(None)
    assert value is None and reason == "vacía"


def test_celda_no_numerica_es_no_disponible():
    value, reason = pa.units("n/a")
    assert value is None and "no numérica" in reason


def test_sum_units_ignora_los_no_disponibles():
    assert pa.sum_units([10, None, 0, 5]) == 15


# ── Hoja completa ───────────────────────────────────────────────────────────
def test_fila_total_no_se_importa_como_producto(tmp_path):
    periods, products, total_row, issues = pa.parse_sheet(ROWS)
    assert [p["sku"] for p in products] == [
        "GA Green Tea 1.4oz", "GA Chai Tea 1.4oz"]
    assert total_row == {"2026-01": 10, "2026-02": 24, "2026-07": 6}
    assert issues == []


def test_build_report_valida_contra_la_fila_total(tmp_path):
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    report = pa.build_report(path)
    v = report["meta"]["validation"]
    assert v["totalRowPresent"] is True
    assert v["allMatch"] is True
    assert [c["computed"] for c in v["checks"]] == [10, 24, 6]


def test_build_report_detecta_total_que_no_cuadra(tmp_path):
    rows = [r[:] for r in ROWS]
    rows[-1] = ["TOTAL", 10, 24, 999]           # julio mal sumado en el Excel
    path = make_sheet(tmp_path / "malo.xlsx", rows)
    report = pa.build_report(path)
    v = report["meta"]["validation"]
    assert v["allMatch"] is False
    assert [c["ok"] for c in v["checks"]] == [True, True, False]


def test_totales_separan_meses_completos_del_parcial(tmp_path):
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    t = pa.build_report(path)["totals"]
    assert t["units"] == 40             # 10 + 24 + 6
    assert t["unitsFullMonths"] == 34   # sin el parcial de julio
    assert (t["fullMonths"], t["partialPeriods"]) == (2, 1)
    assert t["avgFullMonth"] == 17.0


def test_el_periodo_parcial_no_calcula_variacion(tmp_path):
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    periods = pa.build_report(path)["periods"]
    enero, febrero, julio = periods
    assert enero["delta"] is None                    # primer mes: sin referencia
    assert (febrero["delta"], febrero["deltaPct"]) == (14, 140.0)
    assert julio["partial"] is True
    assert julio["delta"] is None and julio["deltaPct"] is None


def test_mejor_mes_por_producto_ignora_el_parcial(tmp_path):
    rows = [r[:] for r in ROWS]
    rows[1] = ["GA Green Tea 1.4oz (Garden of the Andes - Green Tea 1.4oz)", 10, 20, 900]
    rows[-1] = ["TOTAL", 10, 24, 901]
    path = make_sheet(tmp_path / "sell.xlsx", rows)
    green = pa.build_report(path)["products"][0]
    assert green["best"]["key"] == "2026-02"         # no 2026-07 pese a las 900 u.


def test_celda_vacia_no_se_convierte_en_cero(tmp_path):
    rows = [r[:] for r in ROWS]
    rows[2] = ["GA Chai Tea 1.4oz (Garden of the Andes - Chai Tea 1.4oz)", None, 4, 1]
    rows[-1] = ["TOTAL", 10, 24, 6]
    path = make_sheet(tmp_path / "hueco.xlsx", rows)
    report = pa.build_report(path)
    chai = next(p for p in report["products"] if p["sku"] == "GA Chai Tea 1.4oz")
    assert chai["byPeriod"]["2026-01"] is None
    assert chai["missingPeriods"] == ["2026-01"]
    assert report["meta"]["validation"]["issues"]    # queda registrado el hueco


def test_productos_duplicados_fallan(tmp_path):
    rows = [r[:] for r in ROWS]
    rows.insert(2, ROWS[1][:])                       # mismo producto dos veces
    with pytest.raises(RuntimeError, match="duplicados"):
        pa.parse_sheet(rows)


def test_cabecera_inesperada_falla():
    with pytest.raises(RuntimeError, match="Products"):
        pa.parse_sheet([["Producto", datetime.datetime(2026, 1, 26)], ["X", 1]])


def test_find_source_ignora_temporales(tmp_path):
    make_sheet(tmp_path / "GOA_CM Sell Through Data_2026_YTD.xlsx", ROWS)
    (tmp_path / "~$GOA_CM Sell Through Data_2026_YTD.xlsx").write_text("lock")
    (tmp_path / "otra cosa.xlsx").write_text("x")
    found = pa.find_source(tmp_path)
    assert found.name == "GOA_CM Sell Through Data_2026_YTD.xlsx"


def test_find_source_sin_carpeta_devuelve_none(tmp_path):
    assert pa.find_source(tmp_path / "no-existe") is None


# ── El archivo real, si está disponible localmente ──────────────────────────
@pytest.mark.skipif(pa.find_source() is None,
                    reason="El Excel de Alta Gama no está en esta copia (carpeta privada)")
def test_archivo_real_cuadra_con_su_fila_total():
    report = pa.build_report(pa.find_source())
    assert report["meta"]["validation"]["allMatch"] is True
    assert report["totals"]["units"] == sum(p["units"] for p in report["periods"])
    assert all(p["sku"] for p in report["products"])
