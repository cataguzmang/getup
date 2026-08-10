"""C1/C2 del spec 2026-08-10: Company canónico en fuentes/ y aborto por estado irresoluble."""
import openpyxl

import split_excel as sx
from helpers import HEADER, tx, make_workbook, D


def _company_col(path):
    """Valores de la columna Company del archivo canónico escrito."""
    ws = openpyxl.load_workbook(path).active
    rows = list(ws.iter_rows(values_only=True))
    return [r[sx.COL_COMPANY] for r in rows[1:]]


def test_company_ausente_hoja_ny_escribe_canonico(tmp_path, monkeypatch):
    """Julio: las hojas NY vienen sin columna Company → se escribe LatinFood US Corp."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    hdr = [h for h in HEADER if h != "Company"]          # 10 columnas, sin Company
    row = [D(2026, 7, 6), "S1", "[GET01] Jack Mackerel in Brine",
           "Key Food", "Katherine", 8, 8, 8, 34.2, 273.6]
    make_workbook(entrada / "07. GET Jul 2026.xlsx",
                  {"NY Jul - Jack Mackerel": [hdr, row]})
    rc = sx.main(["--no-build"])
    assert rc == 0
    out = tmp_path / "reporte-sanjose" / "fuentes" / "San-Jose-2026-07.xlsx"
    assert _company_col(out) == ["LatinFood US Corp."]


def test_company_basura_hoja_mia_escribe_canonico(tmp_path, monkeypatch):
    """Julio: MIA trae 'Florida - LF' / 'Florida- LatinFood FL Inc.' → LatinFood Florida."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    make_workbook(entrada / "07. GET Jul 2026.xlsx", {"MIA Jul - Jack Mackerel": [
        HEADER,
        tx(D(2026, 7, 2), "S1", "[GET01] x", "c", "sp", company="Florida - LF"),
        tx(D(2026, 7, 3), "S2", "[GET01] x", "c", "sp",
           company="Florida- LatinFood FL Inc."),
    ]})
    rc = sx.main(["--no-build"])
    assert rc == 0
    out = tmp_path / "reporte-sanjose" / "fuentes" / "San-Jose-2026-07.xlsx"
    assert _company_col(out) == ["LatinFood Florida", "LatinFood Florida"]
