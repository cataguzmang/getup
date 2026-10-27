import datetime
import json

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
    src = pa.parse_sheet(ROWS)
    assert [p["sku"] for p in src["products"]] == [
        "GA Green Tea 1.4oz", "GA Chai Tea 1.4oz"]
    assert src["totalRow"] == {"2026-01": 10, "2026-02": 24, "2026-07": 6}
    assert src["issues"] == []


def test_build_report_valida_contra_la_fila_total(tmp_path):
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    report = pa.build_report([pa.load_excel(path)])
    v = report["meta"]["validation"]
    assert v["totalRowPresent"] is True
    assert v["allMatch"] is True
    assert [c["computed"] for c in v["checks"]] == [10, 24, 6]


def test_build_report_detecta_total_que_no_cuadra(tmp_path):
    rows = [r[:] for r in ROWS]
    rows[-1] = ["TOTAL", 10, 24, 999]           # julio mal sumado en el Excel
    path = make_sheet(tmp_path / "malo.xlsx", rows)
    report = pa.build_report([pa.load_excel(path)])
    v = report["meta"]["validation"]
    assert v["allMatch"] is False
    assert [c["ok"] for c in v["checks"]] == [True, True, False]


def test_totales_incluyen_el_parcial_y_conservan_el_corte(tmp_path):
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    t = pa.build_report([pa.load_excel(path)])["totals"]
    assert t["units"] == 40             # 10 + 24 + 6
    assert t["avgPeriod"] == round(40 / 3, 1)        # el parcial cuenta como un mes
    assert t["unitsFullMonths"] == 34   # sigue disponible para el uso interno
    assert (t["fullMonths"], t["partialPeriods"]) == (2, 1)
    assert t["avgFullMonth"] == 17.0


def test_el_periodo_parcial_tambien_calcula_variacion(tmp_path):
    """Decisión comercial: el parcial se presenta como un mes más."""
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    periods = pa.build_report([pa.load_excel(path)])["periods"]
    enero, febrero, julio = periods
    assert enero["delta"] is None                    # primer mes: sin referencia
    assert (febrero["delta"], febrero["deltaPct"]) == (14, 140.0)
    assert (julio["delta"], julio["deltaPct"]) == (-18, -75.0)
    assert julio["partial"] is True                  # el corte real no se pierde
    assert julio["short"] == "Jul"                   # sin adornos en la etiqueta


def test_el_aviso_del_corte_queda_disponible_para_el_correo(tmp_path):
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    meta = pa.build_report([pa.load_excel(path)])["meta"]
    assert meta["partialNote"]["key"] == "2026-07"
    assert "hasta el 25 de julio" in meta["partialNote"]["short"]
    assert meta["cutoffDate"] == "2026-07-25"
    assert "(al" not in meta["periodLabel"]          # el encabezado va limpio


def test_mejor_mes_por_producto_considera_todos_los_periodos(tmp_path):
    rows = [r[:] for r in ROWS]
    rows[1] = ["GA Green Tea 1.4oz (Garden of the Andes - Green Tea 1.4oz)", 10, 20, 900]
    rows[-1] = ["TOTAL", 10, 24, 901]
    path = make_sheet(tmp_path / "sell.xlsx", rows)
    green = pa.build_report([pa.load_excel(path)])["products"][0]
    assert green["best"]["key"] == "2026-07"


def test_highlights_del_dashboard(tmp_path):
    path = make_sheet(tmp_path / "sell.xlsx", ROWS)
    h = pa.build_report([pa.load_excel(path)])["highlights"]
    assert h["bestMonth"] == {"key": "2026-02", "label": "Febrero 2026", "units": 24}
    assert h["lastMonth"]["key"] == "2026-07"
    assert h["lastMonth"]["units"] == 6


def test_celda_vacia_no_se_convierte_en_cero(tmp_path):
    rows = [r[:] for r in ROWS]
    rows[2] = ["GA Chai Tea 1.4oz (Garden of the Andes - Chai Tea 1.4oz)", None, 4, 1]
    rows[-1] = ["TOTAL", 10, 24, 6]
    path = make_sheet(tmp_path / "hueco.xlsx", rows)
    report = pa.build_report([pa.load_excel(path)])
    chai = next(p for p in report["products"] if p["sku"] == "GA Chai Tea 1.4oz")
    assert chai["byPeriod"]["2026-01"] is None
    assert chai["missingPeriods"] == ["2026-01"]
    assert report["meta"]["validation"]["issues"]    # queda registrado el hueco


def test_productos_duplicados_fallan(tmp_path):
    rows = [r[:] for r in ROWS]
    rows.insert(2, ROWS[1][:])                       # mismo producto dos veces
    with pytest.raises(RuntimeError, match="duplicados"):
        pa.parse_sheet(rows)


def test_sin_encabezados_de_periodo_falla():
    with pytest.raises(RuntimeError, match="encabezados de periodo"):
        pa.parse_sheet([["Producto", "Q1"], ["X", 1]])


def test_find_sources_ignora_temporales(tmp_path):
    make_sheet(tmp_path / "b.xlsx", ROWS)
    make_sheet(tmp_path / "a.xlsx", ROWS)
    (tmp_path / "~$a.xlsx").write_text("lock")
    assert [p.name for p in pa.find_sources(tmp_path)] == ["a.xlsx", "b.xlsx"]


def test_find_sources_sin_carpeta_devuelve_lista_vacia(tmp_path):
    assert pa.find_sources(tmp_path / "no-existe") == []


# ── Layout jul–oct 2026: tabla corrida, fila Qty, marcas mezcladas ──────────
NUEVO = [
    [None, None, None, None, None, "Jul 26", "Aug 26", "Oct 1 - 26, 26", "TOTAL"],
    [None, None, None, None, None, "Qty", "Qty", "Qty", "Qty"],
    [None, "Products"],
    [None, None, None, None, "GA Green Tea 1.4oz (Garden of the Andes - Green Tea 1.4oz)", 72, 72, 42, 186],
    [None, None, None, None, "GA Chai Tea 1.4oz (Garden of the Andes - Chai Tea 1.4oz)", 12, 29, 18, 59],
    [None, None, None, None, "Kombuchacha Ginger 12oz (Kombuchacha- USA-Ginger- 12oz)", 0, 240, 24, 264],
    [None, None, None, None, "El Charrua Chimichurri 200g (El Charrua Gourmet - Peru - Chimichurri - 200 g)", 168, 96, 72, 336],
    [None, None, None, "Total Inventory", None, 252, 437, 156, 845],
]


def test_layout_nuevo_solo_importa_goa():
    src = pa.parse_sheet(NUEVO)
    assert [p["key"] for p in src["periods"]] == ["2026-07", "2026-08", "2026-10"]
    assert src["periods"][-1]["partial"] is True
    assert [p["sku"] for p in src["products"]] == ["GA Green Tea 1.4oz", "GA Chai Tea 1.4oz"]
    assert [e["label"].split(" (")[0] for e in src["excluded"]] == [
        "Kombuchacha Ginger 12oz", "El Charrua Chimichurri 200g"]
    assert src["issues"] == []


def test_total_inventory_es_fila_de_total_y_cuadra_con_todas_las_filas():
    src = pa.parse_sheet(NUEVO)
    assert src["totalRow"] == {"2026-07": 252, "2026-08": 437, "2026-10": 156}
    assert src["allRows"] == src["totalRow"]


def test_columna_total_por_fila_que_no_cuadra_es_incidencia():
    rows = [r[:] for r in NUEVO]
    rows[3][-1] = 999
    assert any("columna TOTAL" in i for i in pa.parse_sheet(rows)["issues"])


# ── Combinación de fuentes ──────────────────────────────────────────────────
def test_mes_completo_reemplaza_al_parcial_y_conserva_lo_anterior(tmp_path):
    viejo = pa.load_excel(make_sheet(tmp_path / "ytd.xlsx", ROWS))      # ene, feb, jul 1-25
    nuevo = pa.load_excel(make_sheet(tmp_path / "jul-oct.xlsx", NUEVO))  # jul, ago, oct 1-26
    report = pa.build_report([nuevo, viejo])        # el orden de entrada no importa
    assert [p["key"] for p in report["periods"]] == [
        "2026-01", "2026-02", "2026-07", "2026-08", "2026-10"]
    julio = next(p for p in report["periods"] if p["key"] == "2026-07")
    assert julio["partial"] is False and julio["units"] == 84
    v = report["meta"]["validation"]
    assert [(o["key"], o["before"], o["after"]) for o in v["overrides"]] == [("2026-07", 6, 84)]
    assert v["issues"] == []
    assert [s["file"] for s in report["meta"]["sources"]] == ["ytd.xlsx", "jul-oct.xlsx"]
    assert report["meta"]["partialNote"]["key"] == "2026-10"


def test_producto_que_baja_al_completar_el_mes_es_incidencia(tmp_path):
    rows = [r[:] for r in NUEVO]
    rows[4][5] = 0                                   # Chai jul completo < jul 1-25 (1)
    viejo = pa.load_excel(make_sheet(tmp_path / "ytd.xlsx", ROWS))
    nuevo = pa.load_excel(make_sheet(tmp_path / "jul-oct.xlsx", rows))
    issues = pa.build_report([viejo, nuevo])["meta"]["validation"]["issues"]
    assert any("bajó de 1 a 0" in i for i in issues)


def test_el_historico_siempre_pierde(tmp_path):
    hist = tmp_path / "hist.json"
    hist.write_text(json.dumps({
        "source": "ytd.xlsx",
        "periods": [{"key": "2026-01", "start": "2026-01-01", "end": "2026-01-31"}],
        "products": [{"sku": "GA Green Tea 1.4oz",
                      "name": "Garden of the Andes - Green Tea 1.4oz",
                      "byPeriod": {"2026-01": 999}}],
        "totalRow": {"2026-01": 999},
    }))
    historico = pa.load_history(hist)
    excel = pa.load_excel(make_sheet(tmp_path / "ytd.xlsx", ROWS))
    report = pa.build_report([excel, historico])
    green = next(p for p in report["products"] if p["sku"] == "GA Green Tea 1.4oz")
    assert green["byPeriod"]["2026-01"] == 10


def test_historico_del_repo_reproduce_lo_publicado():
    report = pa.build_report([pa.load_history()])
    assert [p["units"] for p in report["periods"]][:6] == [330, 748, 474, 456, 324, 264]
    assert report["meta"]["validation"]["allMatch"] is True


# ── Los archivos reales, si están disponibles localmente ────────────────────
@pytest.mark.skipif(not pa.find_sources(),
                    reason="El Excel de Alta Gama no está en esta copia (carpeta privada)")
def test_archivos_reales_cuadran_y_no_dejan_incidencias():
    sources = [pa.load_history()] + [pa.load_excel(p) for p in pa.find_sources()]
    report = pa.build_report(sources)
    v = report["meta"]["validation"]
    assert v["allMatch"] is True
    assert v["issues"] == []
    assert report["totals"]["units"] == sum(p["units"] for p in report["periods"])
    assert all(p["sku"].startswith(pa.SKU_PREFIX) for p in report["products"])
