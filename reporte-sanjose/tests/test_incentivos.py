from sjhelpers import make_canonical, canon_row, D
import generar_data as g

HEADER_LEN = 11


def _pad(row):
    return list(row) + [None] * (HEADER_LEN - len(row))


def _fuentes_con_bloque(tmp_path):
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
        _pad(["Angela Quimbay", 5, None, "Tomate == > 18 cases", 540]),
        _pad(["Lina Raquel", 50, 11, None, 628.68]),
        _pad([None, None, 21, "Clientes Nuevos", 15]),
        _pad(["NEW CUSTOMERS"]),
        _pad(["Alexandra J"]),
        _pad(["Angela Q"]),
        _pad(["Raquel L", 1]),
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    return fuentes


def test_load_incentivos(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", _fuentes_con_bloque(tmp_path))
    inc = g.load_incentivos()
    fl = inc[(2026, 6)]["Florida"]
    productos = {d["product"]: d for d in fl["descuentos"]}
    assert productos["In Brine"]["cases"] == 3 and productos["In Brine"]["amount"] == 88.68
    assert productos["Tomate"]["cases"] == 18 and productos["Tomate"]["amount"] == 540
    assert round(fl["totalDescuentos"], 2) == 628.68
    assert fl["freeReportado"]["Alexandra Jimenez"] == 10
    assert fl["soldReportado"]["Lina Raquel"] == 50
    assert "Alexandra J" in fl["vendedoresNuevos"]


def test_load_incentivos_vacio(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 6, 1), "S1", "[GET01] x", "C", "SP", "LatinFood Florida", 1, 45.6, 45.6),
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    assert g.load_incentivos() == {}


def test_nombre_vendedor_no_se_confunde_con_producto(tmp_path, monkeypatch):
    """Un vendedor con 'Tomate'/'In Brine' en el nombre (col A) NO debe generar un
    descuento espurio, y un descuento real en la misma fila (col D) se respeta."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    rows = [
        canon_row(D(2026, 6, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "Tomate Gonzalez", "LatinFood Florida", 20, 45.6, 912),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 5, 1), "SOLD", "FREE", "Descuentos"]),
        # nombre con 'Tomate' + descuento real de In Brine en col D
        _pad(["Tomate Gonzalez", 60, 10, "In Brine == > 3 cases", 88.68]),
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    fl = g.load_incentivos()[(2026, 6)]["Florida"]
    productos = {d["product"] for d in fl["descuentos"]}
    assert productos == {"In Brine"}          # NO aparece 'Tomate' por el nombre
    assert fl["totalDescuentos"] == 88.68
    assert fl["soldReportado"]["Tomate Gonzalez"] == 60
