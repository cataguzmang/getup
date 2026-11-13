"""Empalme: sin las fuentes viejas (viven en otro computador), los meses nuevos se
agregan sobre el data.js anterior en vez de regenerarlo y perder el histórico.
Invariante: empalmar marzo sobre (feb) da lo mismo que generar (feb + mar) juntos,
salvo el orden de filas empatadas."""
import json

import pytest

from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def _feb():
    return [
        canon_row(D(2026, 2, 3), "S1", "[GET01] Jack Mackerel in Brine", "Cliente A", "Ana",
                  "LatinFood Florida", 10, 45.6, 456),
        canon_row(D(2026, 2, 3), "S1", "[GET01] Jack Mackerel in Brine", "Cliente A", "Ana",
                  "LatinFood Florida", 2, 0, 0),
        canon_row(D(2026, 2, 9), "S2", "[GET02] Jack Mackerel in Tomato Sauce", "Cliente B",
                  "Luis", "LatinFood US Corp.", 3, 45.6, 136.8),
    ]


def _mar():
    return [
        canon_row(D(2026, 3, 4), "S3", "[GET01] Jack Mackerel in Brine", "Cliente A", "Ana",
                  "LatinFood Florida", 5, 45.6, 228),
        canon_row(D(2026, 3, 5), "S4", "[GET02] Jack Mackerel in Tomato Sauce", "Cliente C",
                  "Eva", "LatinFood US Corp.", 4, 45.6, 182.4),
        canon_row(D(2026, 3, 5), "S4", "[GET02] Jack Mackerel in Tomato Sauce", "Cliente C",
                  "Eva", "LatinFood US Corp.", 1, 0, 0),
    ]


def _run(tmp_path, monkeypatch, files, out):
    fuentes = tmp_path / "fuentes"
    if fuentes.exists():
        for p in fuentes.glob("*.xlsx"):
            p.unlink()
    fuentes.mkdir(exist_ok=True)
    for name, rows in files.items():
        make_canonical(fuentes / name, rows)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    monkeypatch.setattr(g, "OUTPUT_FILE", out)
    g.main()
    return g.load_previous_data(out)


def _sin_orden(x):
    if isinstance(x, dict):
        return {k: _sin_orden(v) for k, v in x.items()}
    if isinstance(x, list):
        return sorted((_sin_orden(v) for v in x), key=lambda v: json.dumps(v, sort_keys=True))
    return x


def test_empalme_equivale_a_regenerar(tmp_path, monkeypatch):
    completo = _run(tmp_path, monkeypatch,
                    {"San-Jose-2026-02.xlsx": _feb(), "San-Jose-2026-03.xlsx": _mar()},
                    tmp_path / "completo.js")
    empalme_out = tmp_path / "empalme.js"
    _run(tmp_path, monkeypatch, {"San-Jose-2026-02.xlsx": _feb()}, empalme_out)
    empalme = _run(tmp_path, monkeypatch, {"San-Jose-2026-03.xlsx": _mar()}, empalme_out)

    assert empalme["MONTHS_BTN"] == ["Feb 2026", "Mar 2026"]
    assert empalme["D"] == completo["D"]
    assert empalme["PROD_FIRST_MONTH"] == completo["PROD_FIRST_MONTH"]
    assert _sin_orden(empalme["GENERAL_DATA"]) == _sin_orden(completo["GENERAL_DATA"])
    # Cliente C es nuevo en marzo; Cliente A no (ya estaba en febrero)
    nuevos = [c["name"] for c in empalme["GENERAL_DATA"]["Mar 26"]["clientesNuevosLista"]]
    assert nuevos == ["Cliente C"]


def test_empalme_aborta_si_reprocesa_un_mes_publicado_sin_historico(tmp_path, monkeypatch):
    out = tmp_path / "data.js"
    _run(tmp_path, monkeypatch,
         {"San-Jose-2026-02.xlsx": _feb(), "San-Jose-2026-03.xlsx": _mar()}, out)
    with pytest.raises(SystemExit):
        _run(tmp_path, monkeypatch, {"San-Jose-2026-03.xlsx": _mar()}, out)
