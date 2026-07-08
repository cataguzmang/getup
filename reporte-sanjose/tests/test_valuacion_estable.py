from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def _fuentes(tmp_path, extra_junio=False):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir(parents=True)
    rows_may = [
        canon_row(D(2026, 5, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 45.6, 182.4),
        canon_row(D(2026, 5, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 2, 0, 0),
    ]
    make_canonical(fuentes / "may.xlsx", rows_may)
    if extra_junio:
        make_canonical(fuentes / "jun.xlsx", [
            canon_row(D(2026, 6, 1), "S9", "[GET01] Jack Mackerel in Brine",
                      "C", "SP", "LatinFood Florida", 10, 60.0, 600.0),
        ])
    return fuentes


def _pv_mayo(fuentes, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    D_ = g.build_monthly_state_data(rows, months, op, pa, ga)
    i = months.index((2026, 5))
    return D_["fl"]["pv"][i]


def test_pv_estable_al_agregar_mes(tmp_path, monkeypatch):
    pv_solo = _pv_mayo(_fuentes(tmp_path / "a", extra_junio=False), monkeypatch)
    pv_con_junio = _pv_mayo(_fuentes(tmp_path / "b", extra_junio=True), monkeypatch)
    assert pv_solo == pv_con_junio
    assert pv_solo == round(2 * 45.6, 2)
