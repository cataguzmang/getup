from sjhelpers import make_canonical, canon_row, D
import generar_data as g

HEADER_LEN = 11
def _pad(row): return list(row) + [None] * (HEADER_LEN - len(row))


def test_crosscheck_free_match(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    rows = [
        canon_row(D(2026, 6, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", 20, 45.6, 912),
        canon_row(D(2026, 6, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", 21, 45.6, 0),  # 21 promo
    ]
    # SOLD del vendedor = total entregado (41 = 20 pagadas + 21 free); FREE = 21
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 5, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 41, 21, None, None]),   # SOLD 41 (entregado), FREE 21
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    r = g.load_rows()
    months = sorted(set(g.month_key(x["date"]) for x in r))
    op, pa, ga = g.get_order_unit_prices(r)
    inc = g.load_incentivos()
    period = g.build_period_data(r, months, op, pa, ga, inc, {}, month_set=set(months))
    cc = period["crossCheck"]["Florida"]
    # free: 21 cajas promo calculadas == 21 reportadas
    assert cc["free"]["computado"] == 21 and cc["free"]["reportado"] == 21
    # sold = total entregado: 41 calculadas == 41 reportadas (cuadra)
    assert cc["sold"]["computado"] == 41 and cc["sold"]["reportado"] == 41


def _period_con_sold(tmp_path, monkeypatch, pagadas, promo, sold_reportado, mes=7):
    """Mes con `pagadas` cajas pagadas + `promo` gratis y un bloque que reporta
    SOLD `sold_reportado` / FREE `promo`."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    rows = [
        canon_row(D(2026, mes, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", pagadas, 38.4,
                  round(pagadas * 38.4, 2)),
        canon_row(D(2026, mes, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", promo, 38.4, 0),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, mes - 1, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", sold_reportado, promo]),
    ]
    make_canonical(fuentes / f"San-Jose-2026-{mes:02d}.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    r = g.load_rows()
    months = sorted(set(g.month_key(x["date"]) for x in r))
    op, pa, ga = g.get_order_unit_prices(r)
    inc = g.load_incentivos()
    return g.build_period_data(r, months, op, pa, ga, inc, {}, month_set=set(months))


def test_crosscheck_sold_base_entregadas_sin_campo(tmp_path, monkeypatch):
    """Junio: SOLD = entregadas. Calza como siempre y NO gana el campo base
    (el data.js histórico debe quedar idéntico)."""
    cc = _period_con_sold(tmp_path, monkeypatch, pagadas=20, promo=21,
                          sold_reportado=41, mes=6)["crossCheck"]["Florida"]
    assert cc["sold"] == {"computado": 41, "reportado": 41}


def test_crosscheck_sold_base_pagadas(tmp_path, monkeypatch):
    """Julio: SOLD = pagadas (entregadas − free). Calza contra pagadas y lo registra."""
    cc = _period_con_sold(tmp_path, monkeypatch, pagadas=51, promo=10,
                          sold_reportado=51)["crossCheck"]["Florida"]
    assert cc["sold"] == {"computado": 51, "reportado": 51, "base": "pagadas"}


def test_crosscheck_sold_sin_base_no_calza(tmp_path, monkeypatch):
    """SOLD que no calza ni con entregadas ni con pagadas → mismatch real."""
    cc = _period_con_sold(tmp_path, monkeypatch, pagadas=20, promo=21,
                          sold_reportado=40)["crossCheck"]["Florida"]
    assert cc["sold"] == {"computado": 41, "reportado": 40}   # sin base; computado ≠ reportado
