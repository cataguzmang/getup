"""Núcleo de Alta Gama con otra marca: el mismo Excel, filtrado por prefijo de SKU.

Los casos de lectura, validación y combinación de fuentes viven en
goa-inventory/tests/test_parse_altagama.py (vía el wrapper de GOA)."""
import altagama_core as core

KOM = "Kombuchacha "

ROWS = [
    [None, None, None, None, None, "Jul 26", "Aug 26", "Oct 1 - 26, 26", "TOTAL"],
    [None, None, None, None, None, "Qty", "Qty", "Qty", "Qty"],
    [None, "Products"],
    [None, None, None, None, "GA Green Tea 1.4oz (Garden of the Andes - Green Tea 1.4oz)", 72, 72, 42, 186],
    [None, None, None, None, "Kombuchacha Blueberry 12oz (Kombuchacha- USA- Blueberry- 12oz)", 0, 216, 0, 216],
    [None, None, None, None, "Kombuchacha Ginger 12oz (Kombuchacha- USA-Ginger- 12oz)", 0, 240, 24, 264],
    [None, None, None, None, "El Charrua Chimichurri 200g (El Charrua Gourmet - Peru - Chimichurri - 200 g)", 168, 96, 72, 336],
]


def test_kombuchacha_solo_importa_sus_filas():
    src = core.parse_sheet(ROWS, KOM, "Kombuchacha")
    assert [p["sku"] for p in src["products"]] == [
        "Kombuchacha Blueberry 12oz", "Kombuchacha Ginger 12oz"]
    assert len(src["excluded"]) == 2
    assert src["issues"] == []


def test_nombre_corto_sin_guion_de_marca_sale_del_sku():
    src = core.parse_sheet(ROWS, KOM, "Kombuchacha")
    assert [p["shortName"] for p in src["products"]] == ["Blueberry 12oz", "Ginger 12oz"]


def test_short_name_de_goa_no_cambia():
    assert core.short_name("Garden of the Andes - Green Tea 1.4oz",
                           "GA Green Tea 1.4oz", "GA ") == "Green Tea 1.4oz"


def test_marca_sin_historico(tmp_path):
    assert core.load_history(None) is None
    assert core.load_history(tmp_path / "no-existe.json") is None


def test_reporte_de_kombuchacha_los_ceros_cuentan(tmp_path):
    report = core.build_report([dict(core.parse_sheet(ROWS, KOM, "Kombuchacha"), name="x.xlsx")],
                               "Kombuchacha")
    assert report["meta"]["brand"] == "Kombuchacha"
    assert [p["units"] for p in report["periods"]] == [0, 456, 24]
    assert report["highlights"]["bestMonth"]["key"] == "2026-08"
