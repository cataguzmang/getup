import generar_data as g


def test_helpers_de_razon():
    assert g._promo_pct(100, 25) == 25.0
    assert g._promo_pct(0, 0) == 0.0
    assert g._roi(900, 300) == 3.0
    assert g._roi(900, 0) == 0.0
    assert g._rev_por_caja(912, 20, 0) == 45.6
    assert g._rev_por_caja(0, 0, 0) == 0.0


def test_summary_y_states_traen_razones(tmp_path, monkeypatch):
    from sjhelpers import make_canonical, canon_row, D
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 6, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 20, 45.6, 912),
        canon_row(D(2026, 6, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 5, 45.6, 0),   # promo
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    period = g.build_period_data(rows, months, op, pa, ga, month_set=set(months))
    s = period["summary"]
    assert s["promoPct"] == g._promo_pct(s["cajas"], s["pc"])
    assert s["roi"] == g._roi(s["rev"], s["pv"])
    assert s["revPorCaja"] == g._rev_por_caja(s["rev"], s["cajas"], s["pc"])
    fl = period["states"]["Florida"]
    assert fl["promoPct"] == g._promo_pct(fl["cajas"], fl["pc"])
    assert "roi" in fl and "revPorCaja" in fl
