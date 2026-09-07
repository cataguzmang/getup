"""Archivado de los Excel de entrada tras una corrida limpia.

Motivo (2026-09-07): entrada/ acumulaba los Excel de meses anteriores y cada
corrida los reprocesaba entera. Tras una corrida completa y sin errores, el
dispatcher mueve lo que consumió a entrada/procesados/.
"""
import openpyxl

import split_excel as sx
from helpers import HEADER, tx, make_workbook, D


def _seed(root, nombre="08. GET Ago 2026.xlsx"):
    """entrada/ con un workbook mínimo de una sola marca lista."""
    sheets = {
        "NY Ago - GOA": [
            HEADER,
            tx(D(2026, 8, 12), "S48729", "[GOA01] Green Tea", "Mi Tierra", "Daniela"),
        ],
    }
    entrada = root / "entrada"
    entrada.mkdir(exist_ok=True)
    return make_workbook(entrada / nombre, sheets)


def _seed_generadores(root):
    """Generador falso (exitoso) para cada marca, para simular corrida completa."""
    for brand in sx.BRANDS.values():
        folder = root / brand.folder
        folder.mkdir(parents=True, exist_ok=True)
        (folder / brand.generator).write_text("print('ok')\n", encoding="utf-8")


def test_corrida_limpia_archiva_el_excel(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    src = _seed(tmp_path)
    _seed_generadores(tmp_path)

    assert sx.main([]) == 0

    assert not src.exists(), "el Excel debe salir de entrada/"
    archivado = tmp_path / "entrada" / "procesados" / src.name
    assert archivado.exists(), "el Excel debe quedar en entrada/procesados/"
    assert "procesados" in capsys.readouterr().out


def test_el_archivado_no_toca_el_contenido(tmp_path, monkeypatch):
    """Mover, no reescribir: el .xlsx archivado sigue siendo el original."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    src = _seed(tmp_path)
    original = src.read_bytes()
    _seed_generadores(tmp_path)

    sx.main([])

    assert (tmp_path / "entrada" / "procesados" / src.name).read_bytes() == original


def test_generador_fallido_deja_el_excel_en_entrada(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    src = _seed(tmp_path)
    _seed_generadores(tmp_path)
    goa = sx.BRANDS["GOA"]
    (tmp_path / goa.folder / goa.generator).write_text(
        "import sys; sys.exit(2)\n", encoding="utf-8")

    assert sx.main([]) == 1

    assert src.exists(), "si el build falla, el Excel se queda para reintentar"
    assert not (tmp_path / "entrada" / "procesados" / src.name).exists()
    assert "se quedan en entrada" in capsys.readouterr().out


def test_estado_irresoluble_deja_el_excel_en_entrada(tmp_path, monkeypatch):
    """El abort temprano por estado '?' tampoco puede archivar nada."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    src = make_workbook(entrada / "raro.xlsx", {"XX Ago - GOA": [
        HEADER,
        tx(D(2026, 8, 1), "S1", "[GOA01] x", "c", "sp", company="Otra Cosa"),
    ]})
    _seed_generadores(tmp_path)

    assert sx.main([]) == 1
    assert src.exists()


def test_no_build_no_archiva(tmp_path, monkeypatch, capsys):
    """Sin regenerar los data.js la corrida está a medias: no se archiva."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    src = _seed(tmp_path)

    assert sx.main(["--no-build"]) == 0

    assert src.exists()
    assert "no archivados" in capsys.readouterr().out


def test_only_no_archiva(tmp_path, monkeypatch):
    """Con --only quedan marcas sin procesar: el Excel sigue haciendo falta."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    src = _seed(tmp_path)
    _seed_generadores(tmp_path)

    assert sx.main(["--only", "GOA"]) == 0
    assert src.exists()


def test_keep_input_no_archiva(tmp_path, monkeypatch):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    src = _seed(tmp_path)
    _seed_generadores(tmp_path)

    assert sx.main(["--keep-input"]) == 0
    assert src.exists()


def test_colision_de_nombre_no_pisa_el_archivado(tmp_path, monkeypatch):
    """Reprocesar un mes ya archivado no puede borrar el archivo anterior."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    procesados = tmp_path / "entrada" / "procesados"
    procesados.mkdir(parents=True)
    (procesados / "08. GET Ago 2026.xlsx").write_bytes(b"version-vieja")
    _seed(tmp_path)
    _seed_generadores(tmp_path)

    assert sx.main([]) == 0

    assert (procesados / "08. GET Ago 2026.xlsx").read_bytes() == b"version-vieja"
    assert (procesados / "08. GET Ago 2026-2.xlsx").exists()


def test_iter_input_files_ignora_la_carpeta_procesados(tmp_path):
    """Lo ya archivado no vuelve a entrar al pipeline."""
    (tmp_path / "nuevo.xlsx").write_text("x")
    procesados = tmp_path / "procesados"
    procesados.mkdir()
    (procesados / "viejo.xlsx").write_text("x")

    assert [f.name for f in sx.iter_input_files(tmp_path)] == ["nuevo.xlsx"]
