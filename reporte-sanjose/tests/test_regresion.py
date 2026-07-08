"""
Regresión de SP3a. La base ahora son los números CORREGIDOS (reference_sp3.json):
regla de montos por mes, cajas gratis bien detectadas, valuación de promos estable,
descuentos → revenue neto, y clientes nuevos calculados.

- Guarda contra cambios accidentales futuros (todos los meses, todo GENERAL_DATA/D).
- Verifica que junio trae descuentos → revNeto y clientes nuevos calculados.
"""
import json
import re
from pathlib import Path

import pytest

import generar_data as g

SANJOSE = Path(__file__).resolve().parent.parent
FUENTES = SANJOSE / "fuentes"
REF = json.load(open(Path(__file__).parent / "reference_sp3.json", encoding="utf-8"))


def _parse_data_js(text):
    def grab(name, pat):
        return json.loads(re.search(r"\b" + name + r" = (" + pat + r");", text, re.S).group(1))
    return {"MONTHS": grab("MONTHS", r"\[.*?\]"),
            "D": grab("D", r"\{.*?\}"),
            "GENERAL_DATA": grab("GENERAL_DATA", r"\{.*\}")}


@pytest.mark.skipif(not (FUENTES / "San-Jose-historico-hasta-2026-05.xlsx").exists(),
                    reason="faltan fuentes reales (correr migrar_historico.py + split_excel.py)")
def test_regresion_vs_base_sp3(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "OUTPUT_FILE", tmp_path / "data.js")
    g.main()
    out = _parse_data_js((tmp_path / "data.js").read_text(encoding="utf-8"))
    assert out["MONTHS"] == REF["MONTHS"]
    assert out["GENERAL_DATA"] == REF["GENERAL_DATA"]
    assert out["D"] == REF["D"]


@pytest.mark.skipif(not (FUENTES / "San-Jose-2026-06.xlsx").exists(),
                    reason="falta el fuentes de junio")
def test_junio_tiene_descuentos_y_neto(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "OUTPUT_FILE", tmp_path / "data.js")
    g.main()
    out = _parse_data_js((tmp_path / "data.js").read_text(encoding="utf-8"))
    jun = out["GENERAL_DATA"]["Jun 26"]["summary"]
    assert jun["descuentos"] > 0
    assert jun["revNeto"] == round(jun["rev"] - jun["descuentos"], 2)
    assert jun["clientesNuevos"] >= 0
