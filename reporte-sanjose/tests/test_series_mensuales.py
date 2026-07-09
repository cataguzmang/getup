from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def test_d_total_y_series(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 6, 1), "S1", "[GET01] x", "C", "SP", "LatinFood Florida", 20, 45.6, 912),
        canon_row(D(2026, 6, 2), "S2", "[GET01] x", "C", "SP", "LatinFood US Corp.", 10, 45.6, 456),
        canon_row(D(2026, 6, 3), "S3", "[GET01] x", "C", "SP", "LatinFood Florida", 5, 45.6, 0),  # promo
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    Dd = g.build_monthly_state_data(rows, months, op, pa, ga)
    i = 0
    # total = fl + ny
    assert Dd["total"]["rev"][i] == round(Dd["fl"]["rev"][i] + Dd["ny"]["rev"][i], 2)
    assert Dd["total"]["cajas"][i] == Dd["fl"]["cajas"][i] + Dd["ny"]["cajas"][i]
    # series de razón presentes y coherentes
    assert Dd["fl"]["promoPct"][i] == g._promo_pct(Dd["fl"]["cajas"][i], Dd["fl"]["pc"][i])
    assert Dd["total"]["roi"][i] == g._roi(Dd["total"]["rev"][i], Dd["total"]["pv"][i])
