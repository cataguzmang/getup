from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def _build_fuentes(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", [
        # venta real FL
        canon_row(D(2026, 6, 22), "S1", "[GET01] Jack Mackerel in Brine ",
                  "Key Food", "Alexandra J", "LatinFood Florida", 20, 45.6, 912),
        # venta real NY
        canon_row(D(2026, 6, 18), "S2", "[GET01] Jack Mackerel in Brine",
                  "Food Fair", "Katherine", "LatinFood US Corp.", 1, 21.85, 21.85),
        # caja promo (ambos montos 0)
        canon_row(D(2026, 6, 3), "S3", "[GET02] Jack Mackerel in Tomato Sauce",
                  "Price Choice", "Raquel L", "LatinFood Florida", 3, 0, 0),
        # entregado 0 → se omite
        canon_row(D(2026, 6, 3), "S4", "[GET02] Jack Mackerel in Tomato Sauce",
                  "Price Choice", "Raquel L", "LatinFood Florida", 0, 45.6, 0),
        # bloque de incentivos (se ignora): fila SOLD + fila con nombre suelto
        [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None],
        ["Alexandra J", None, None, None, None, None, None, None, None, None, None],
    ])
    return fuentes


def test_load_rows_filters_and_derives(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", _build_fuentes(tmp_path))
    rows = g.load_rows()
    # 3 filas reales (las dos ventas + la promo); qty0 y el bloque se descartan
    assert len(rows) == 3

    fl = [r for r in rows if r["state"] == "Florida"]
    ny = [r for r in rows if r["state"] == "Nueva York"]
    assert len(ny) == 1 and ny[0]["company"] == "LatinFood US Corp."
    assert len(fl) == 2

    # nombre de producto sin [SKU] y sin espacios sobrantes
    assert {r["product"] for r in rows} == {"Jack Mackerel in Brine",
                                            "Jack Mackerel in Tomato Sauce"}
    # promo detectada (ambos montos 0)
    promo = [r for r in rows if r["is_promo"]]
    assert len(promo) == 1 and promo[0]["qty"] == 3
    # regla de montos actual (max) conservada
    venta = [r for r in rows if r["order"] == "S1"][0]
    assert venta["line_rev"] == 912 and venta["box_price"] == 912 / 20


def test_load_rows_swap_orientation(tmp_path, monkeypatch):
    """Mayo 2026 viene con columnas intercambiadas (Unit Price = total de línea,
    Total = precio por caja). La regla max() debe seguir sacando el total de línea."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "San-Jose-2026-05.xlsx", [
        # qty 5, precio por caja 45.6 → total 228, pero VIENE INTERCAMBIADO:
        # Unit Price = 228 (total de línea), Total = 45.6 (precio por caja)
        canon_row(D(2026, 5, 20), "S9", "[GET01] Jack Mackerel in Brine",
                  "Cliente", "SP", "LatinFood Florida", 5, 228, 45.6),
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    r = g.load_rows()[0]
    assert r["line_rev"] == 228            # max(45.6, 228)
    assert r["box_price"] == 228 / 5       # 45.6 por caja
    assert r["is_promo"] is False


def test_load_rows_warns_on_unmapped_company(tmp_path, monkeypatch, capsys):
    """Un Company que no mapea a estado se omite pero se avisa (no drena revenue en silencio)."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", [
        canon_row(D(2026, 6, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "Distribuidora Nueva SA", 3, 45.6, 136.8),
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    assert rows == []
    assert "Distribuidora Nueva SA" in capsys.readouterr().out
