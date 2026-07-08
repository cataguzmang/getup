"""
Regresión de SP2 (plomería). Dos garantías distintas:

1. PLOMERÍA FIEL — correr el pipeline nuevo SOLO con el histórico migrado reproduce
   IDÉNTICO el data.js viejo (referencia). Prueba que migración + refactor no cambió nada.

2. CON JUNIO — al sumar junio, los campos NO-promo de los meses históricos quedan
   idénticos y aparece "Jun 26". Los valores `pv` (inversión en promos) de meses
   históricos pueden moverse centavos porque la valuación de cajas gratis usa el
   promedio de precios de TODO el pool cargado (acoplamiento conocido; su revisión es SP3).
"""

import copy
import json
import re
import shutil
from pathlib import Path

import pytest

import generar_data as g

SANJOSE = Path(__file__).resolve().parent.parent
HIST = SANJOSE / "fuentes" / "San-Jose-historico-hasta-2026-05.xlsx"
REF = json.load(open(Path(__file__).parent / "reference_pre_sp2.json", encoding="utf-8"))

HIST_MONTHS = ["Oct 25", "Nov 25", "Dic 25", "Ene 26", "Feb 26", "Mar 26", "Abr 26", "May 26"]


def _parse_data_js(text):
    def grab(name, pat):
        return json.loads(re.search(r"\b" + name + r" = (" + pat + r");", text, re.S).group(1))
    return {
        "MONTHS": grab("MONTHS", r"\[.*?\]"),
        "D": grab("D", r"\{.*?\}"),
        "GENERAL_DATA": grab("GENERAL_DATA", r"\{.*\}"),
    }


def _strip_pv(view):
    """Copia de una vista-mes sin los campos de inversión en promos (pv)."""
    v = copy.deepcopy(view)
    v.get("summary", {}).pop("pv", None)
    for s in v.get("states", {}).values():
        s.pop("pv", None)
    return v


def _run(monkeypatch, tmp_path, sources_dir):
    monkeypatch.setattr(g, "SOURCES_DIR", sources_dir)
    monkeypatch.setattr(g, "OUTPUT_FILE", tmp_path / "data.js")
    g.main()
    return _parse_data_js((tmp_path / "data.js").read_text(encoding="utf-8"))


@pytest.mark.skipif(not HIST.exists(),
                    reason="falta el histórico migrado (correr migrar_historico.py)")
def test_regresion_plomeria_historico_solo(tmp_path, monkeypatch):
    """Solo el histórico → data.js IDÉNTICO a la referencia (plomería fiel)."""
    src = tmp_path / "fuentes"
    src.mkdir()
    shutil.copy(HIST, src / HIST.name)
    out = _run(monkeypatch, tmp_path, src)

    assert out["MONTHS"] == REF["MONTHS"]
    for m in HIST_MONTHS:
        assert out["GENERAL_DATA"][m] == REF["GENERAL_DATA"][m], f"cambió {m}"
    for estado in ("fl", "ny"):
        for metrica, serie_ref in REF["D"][estado].items():
            assert out["D"][estado][metrica] == serie_ref, f"D.{estado}.{metrica} cambió"


@pytest.mark.skipif(not HIST.exists(),
                    reason="falta el histórico migrado (correr migrar_historico.py)")
def test_regresion_con_junio_agrega_mes(tmp_path, monkeypatch):
    """Con el fuentes/ real (histórico + junio): agrega Jun 26; los campos no-pv de los
    meses históricos quedan idénticos; pv histórico solo se mueve centavos (deuda SP3)."""
    out = _run(monkeypatch, tmp_path, g.SOURCES_DIR)  # SOURCES_DIR real por defecto

    assert "Jun 26" in out["MONTHS"]
    assert len(out["D"]["fl"]["rev"]) == 9

    # Campos NO-promo de los meses históricos: idénticos
    for m in HIST_MONTHS:
        assert _strip_pv(out["GENERAL_DATA"][m]) == _strip_pv(REF["GENERAL_DATA"][m]), \
            f"cambió algo no-pv en {m}"
    for estado in ("fl", "ny"):
        for metrica in ("rev", "cajas", "pc", "ord"):
            assert out["D"][estado][metrica][:8] == REF["D"][estado][metrica], \
                f"D.{estado}.{metrica} histórico cambió"

    # pv histórico: solo se mueve por el pool de promedios de promos (acoplamiento SP3)
    for estado in ("fl", "ny"):
        for i in range(8):
            assert abs(out["D"][estado]["pv"][i] - REF["D"][estado]["pv"][i]) < 1.0
