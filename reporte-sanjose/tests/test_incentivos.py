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


def test_descuento_sin_monto_no_agarra_el_free(tmp_path, monkeypatch):
    """Julio: 'In Brine == > 0 cases' con la celda de monto vacía → $0, no el 5
    de la columna FREE (bug C4: $305 en vez de $300)."""
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    rows = [
        canon_row(D(2026, 7, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", 10, 38.4, 384),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 6, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 20, 5, "In Brine == > 0 cases"]),   # monto vacío
        _pad(["Angela Quimbay", 31, 5, "Tomate == > 12 cases", 300]),
    ]
    make_canonical(fuentes / "San-Jose-2026-07.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    fl = g.load_incentivos()[(2026, 7)]["Florida"]
    productos = {d["product"]: d for d in fl["descuentos"]}
    assert productos["In Brine"]["amount"] == 0.0
    assert productos["Tomate"]["amount"] == 300
    assert fl["totalDescuentos"] == 300.0


# ── Septiembre 2026: layout "matriz" ─────────────────────────────────────────
# El distribuidor pasó a un bloque con el nombre del vendedor en la columna
# inmediatamente a la izquierda de SOLD (que ya no es la col A), una subfila
# Brine/Tomato bajo cada grupo, y los descuentos como montos por producto
# (Florida) o como una sola columna "Descuentos" (Nueva York).

def _bloque_matriz_ny():
    return [
        _pad(["__INCENTIVOS__", "Nueva York"]),
        [None] * 5 + [D(2026, 9, 1), "SOLD", None, "FREE", None, "Free Total", "Descuentos"],
        [None] * 6 + ["Brine", "Tomato", "Brine", "Tomato", None, None],
        [None] * 5 + ["Daniela Davila", None, 1, None, None, 0, None],
        [None] * 5 + ["Luis Soler", None, None, None, None, 0, None],
        [None] * 5 + ["Mireya Fernandez", 5, 5, 3, None, 3, 88.67999999999999],
        [None] * 5 + ["Total", 5, 7, 3, 0, 3, None],
    ]


def _bloque_matriz_fl():
    return [
        _pad(["__INCENTIVOS__", "Florida"]),
        [None] * 4 + [D(2026, 9, 1), "SOLD", None, "Units FREE", None,
                      "Total Units FREE", None, None, None],
        [None] * 5 + ["Brine", "Tomato"] * 4,
        [None] * 4 + ["Alexandra", 1, 1, None, None, 0, 0, 0, 0],
        [None] * 4 + ["Angela", None, 1, None, None, None, None, None, None],
        [None] * 4 + ["Raquel", 9, 9, 1, 2, 1, 2, 29.56, 60],
        [None] * 4 + ["Total", 10, 11, 1, 2, 1, 2, 29.56, 60],
        [None] * 11 + [89.56, None],
    ]


def _fuentes_matriz(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    rows = [
        canon_row(D(2026, 9, 30), "S1", "[GET01] Jack Mackerel in Brine",
                  "Extra", "Mireya Fernandez", "LatinFood US Corp.", 2, 45.6, 91.2),
        canon_row(D(2026, 9, 11), "S2", "[GET01] Jack Mackerel in Brine",
                  "Ideal", "Raquel L", "LatinFood Florida", 5, 45.6, 228),
    ]
    make_canonical(fuentes / "San-Jose-2026-09.xlsx",
                   rows + _bloque_matriz_ny() + [_pad([])] + _bloque_matriz_fl())
    return fuentes


def test_matriz_descuentos_por_producto_florida(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", _fuentes_matriz(tmp_path))
    fl = g.load_incentivos()[(2026, 9)]["Florida"]
    productos = {d["product"]: d for d in fl["descuentos"]}
    assert productos["In Brine"] == {"product": "In Brine", "cases": 1, "amount": 29.56}
    assert productos["Tomate"] == {"product": "Tomate", "cases": 2, "amount": 60.0}
    assert fl["totalDescuentos"] == 89.56
    assert fl["soldReportado"] == {"Alexandra": 2, "Angela": 1, "Raquel": 18}
    assert fl["freeReportado"] == {"Raquel": 3}
    assert "Total" not in fl["soldReportado"]      # la fila de totales no es vendedor


def test_matriz_descuento_unico_se_atribuye_al_producto_con_free(tmp_path, monkeypatch):
    """NY trae una sola columna 'Descuentos' sin producto: se atribuye al único
    producto con cajas gratis de ese vendedor (Brine, 3 cajas)."""
    monkeypatch.setattr(g, "SOURCES_DIR", _fuentes_matriz(tmp_path))
    ny = g.load_incentivos()[(2026, 9)]["Nueva York"]
    assert ny["descuentos"] == [{"product": "In Brine", "cases": 3, "amount": 88.68}]
    assert ny["totalDescuentos"] == 88.68
    assert ny["soldReportado"] == {"Daniela Davila": 1, "Mireya Fernandez": 10}
    assert ny["freeReportado"] == {"Mireya Fernandez": 3}
    assert ny["vendedoresNuevos"] == []


def test_layout_legacy_no_se_confunde_con_matriz(tmp_path, monkeypatch):
    """'In Brine == > 3 cases' bajo el encabezado SOLD NO es una subfila de productos."""
    monkeypatch.setattr(g, "SOURCES_DIR", _fuentes_con_bloque(tmp_path))
    fl = g.load_incentivos()[(2026, 6)]["Florida"]
    assert round(fl["totalDescuentos"], 2) == 628.68
    assert fl["soldReportado"]["Lina Raquel"] == 50
