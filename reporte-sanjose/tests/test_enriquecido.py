from sjhelpers import make_canonical, canon_row, D
import generar_data as g

HEADER_LEN = 11
def _pad(row): return list(row) + [None] * (HEADER_LEN - len(row))


def _fuentes(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    rows = [
        canon_row(D(2026, 6, 22), "S1", "[GET01] Jack Mackerel in Brine",
                  "Key Food", "Alexandra J", "LatinFood Florida", 20, 45.6, 912),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 5, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68]),
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    return fuentes


def test_period_incluye_neto_y_descuentos(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", _fuentes(tmp_path))
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    inc = g.load_incentivos()
    nc = g.new_customers_by_month(rows)
    period = g.build_period_data(rows, months, op, pa, ga, inc, nc, month_set=set(months))

    assert period["summary"]["descuentos"] == 88.68
    assert period["summary"]["revNeto"] == round(912 - 88.68, 2)
    assert period["states"]["Florida"]["descuentos"] == 88.68
    assert period["states"]["Florida"]["revNeto"] == round(912 - 88.68, 2)
    prods = {d["product"]: d for d in period["descuentosPorProducto"]}
    assert prods["In Brine"]["amount"] == 88.68 and prods["In Brine"]["state"] == "Florida"
