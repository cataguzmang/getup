from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def _load(tmp_path, monkeypatch, rows):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", rows)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    return g.load_rows()


def test_mes_normal(tmp_path, monkeypatch):
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 4, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 45.6, 182.4),
    ])
    r = rows[0]
    assert r["line_rev"] == 182.4 and r["box_price"] == 45.6 and not r["is_promo"]


def test_mes_swap(tmp_path, monkeypatch):
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 5, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 5, 228, 45.6),   # UP=228=5*45.6
        canon_row(D(2026, 5, 3), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 182.4, 45.6),
    ])
    r = [x for x in rows if x["order"] == "S1"][0]
    assert r["line_rev"] == 228 and r["box_price"] == 45.6


def test_caja_gratis_total_cero(tmp_path, monkeypatch):
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 3, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 5, 45.6, 228),   # ancla normal
        canon_row(D(2026, 3, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 3, 45.6, 0),     # Total=0 → gratis
    ])
    gratis = [x for x in rows if x["order"] == "S2"][0]
    assert gratis["is_promo"] is True and gratis["qty"] == 3


def test_venta_con_descuento_inline(tmp_path, monkeypatch):
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 2, 1), "S0", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 45.6, 182.4),  # ancla normal
        canon_row(D(2026, 2, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 1, 45.6, 41.04),
    ])
    r = [x for x in rows if x["order"] == "S1"][0]
    assert r["line_rev"] == 41.04 and not r["is_promo"]
