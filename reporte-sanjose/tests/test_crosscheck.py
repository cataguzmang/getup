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
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 5, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 20, 21, None, None]),   # SOLD 20, FREE 21
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    r = g.load_rows()
    months = sorted(set(g.month_key(x["date"]) for x in r))
    op, pa, ga = g.get_order_unit_prices(r)
    inc = g.load_incentivos()
    period = g.build_period_data(r, months, op, pa, ga, inc, {}, month_set=set(months))
    cc = period["crossCheck"]["Florida"]
    assert cc["free"]["computado"] == 21 and cc["free"]["reportado"] == 21
    assert cc["sold"]["reportado"] == 20   # sold reportado del bloque
