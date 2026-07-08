from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def test_new_customers(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 4, 1), "S1", "[GET01] x", "Cliente A", "SP", "LatinFood Florida", 1, 45.6, 45.6),
        canon_row(D(2026, 5, 1), "S2", "[GET01] x", "Cliente A", "SP", "LatinFood Florida", 1, 45.6, 45.6),
        canon_row(D(2026, 5, 2), "S3", "[GET01] x", "Cliente B", "Ana", "LatinFood Florida", 1, 45.6, 45.6),
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    nc = g.new_customers_by_month(rows)
    # Cliente A aparece por primera vez en abril (primer mes → NO cuenta como nuevo)
    assert (2026, 4) not in nc or nc[(2026, 4)] == []
    # En mayo, Cliente B es nuevo; Cliente A no
    nombres_may = {c["name"] for c in nc.get((2026, 5), [])}
    assert nombres_may == {"Cliente B"}
    assert nc[(2026, 5)][0]["salesperson"] == "Ana"
