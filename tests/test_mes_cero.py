"""C3 del spec 2026-08-10: marca presente por nombre de hoja pero sin ventas →
archivo canónico solo con encabezado, en el mes dominante del archivo."""
import openpyxl

import split_excel as sx
from helpers import HEADER, tx, make_workbook, D


def test_marca_presente_sin_ventas_escribe_solo_encabezado(tmp_path, monkeypatch):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    make_workbook(entrada / "07. GET Jul 2026.xlsx", {
        "MIA Jul - Jack Mackerel": [
            HEADER,
            tx(D(2026, 7, 2), "S1", "[GET01] x", "c", "sp", company="LatinFood Florida"),
        ],
        "NY Jul - Kombuchacha": [HEADER],   # presente, sin ventas ("No sales reported")
    })
    rc = sx.main(["--no-build"])
    assert rc == 0
    kom = tmp_path / "kombuchacha" / "fuentes" / "KOM-2026-07.xlsx"
    assert kom.exists()
    ws = openpyxl.load_workbook(kom)["KOM"]
    assert list(ws.iter_rows(values_only=True)) == [tuple(sx.HEADER)]
    # ROB no aparece en ninguna hoja → no se inventa archivo
    assert not (tmp_path / "robinson-crusoe" / "fuentes").exists()


def test_archivo_sin_transacciones_no_escribe_nada(tmp_path, monkeypatch):
    """Sin transacciones no hay mes que atribuir: no se escribe nada."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    make_workbook(entrada / "vacio.xlsx", {"NY Jul - Kombuchacha": [HEADER]})
    rc = sx.main(["--no-build"])
    assert rc == 0
    assert not (tmp_path / "kombuchacha" / "fuentes").exists()


def test_brand_from_sheet_name():
    f = sx.brand_from_sheet_name
    assert f("NY Jul - Kombuchacha") == "KOM"
    assert f("MIA Jul - Salmón y  Mejillones") == "ROB"   # acentos y espacios dobles
    assert f("MIA Jul - Jack Mackerel") == "GET"
    assert f("MIA Jun - Jack Mckarel") == "GET"           # el typo de junio
    assert f("MIA Jul - GOA") == "GOA"
    assert f("Hoja Cualquiera") is None
