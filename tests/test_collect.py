import pytest

from helpers import HEADER, make_workbook, D
import split_excel as sx


def test_collect_groups_by_brand_and_month(messy_wb):
    data = sx.collect([messy_wb])
    assert len(data.rows["GOA"]["2026-06"]) == 2
    assert len(data.rows["GET"]["2026-06"]) == 1
    assert len(data.rows["KOM"]["2026-06"]) == 1
    # XYZ no está en BRANDS → no aparece como marca
    assert "XYZ" not in data.rows


def test_collect_unmatched(messy_wb):
    data = sx.collect([messy_wb])
    assert len(data.unmatched) == 1
    assert sx.sku_prefix(data.unmatched[0][sx.COL_VARIANT]) == "XYZ"


def test_collect_incentives_attributed_to_get(messy_wb):
    data = sx.collect([messy_wb])
    segs = data.incentives["GET"]["2026-06"]           # lista de segmentos
    assert isinstance(segs, list) and len(segs) == 1
    seg = segs[0]
    assert seg["state"] == "Florida"                    # bloque de la hoja MIA
    assert seg["rows"][0][1] == "SOLD"                   # ancla del bloque
    # GOA/KOM no traían bloque
    assert "GOA" not in data.incentives
    assert "KOM" not in data.incentives


def test_collect_state_stats(messy_wb):
    data = sx.collect([messy_wb])
    assert data.stats["GET"]["2026-06"]["Florida"] == 1
    assert data.stats["GOA"]["2026-06"]["Nueva York"] == 2


def test_collect_skips_corrupt_file(tmp_path, capsys):
    bad = tmp_path / "roto.xlsx"
    bad.write_text("no soy un xlsx")
    data = sx.collect([bad])
    assert data.rows == {}
    assert data.unmatched == []
    assert "No se pudo leer" in capsys.readouterr().out


def test_collect_returns_plain_dicts_missing_key_raises(messy_wb):
    """La conversión defaultdict→dict debe hacer que una marca ausente lance
    KeyError, no auto-vivificar un {} silencioso."""
    data = sx.collect([messy_wb])
    with pytest.raises(KeyError):
        data.rows["ROB"]
    with pytest.raises(KeyError):
        data.incentives["GOA"]


def test_collect_counter_defaults_to_zero(messy_wb):
    """stats guarda Counters: un estado nunca visto devuelve 0, no KeyError."""
    data = sx.collect([messy_wb])
    assert data.stats["GOA"]["2026-06"]["Florida"] == 0


def test_collect_block_without_transactions_warns(tmp_path, capsys):
    """Bloque de incentivos en una hoja sin filas de pedido: se avisa y no se
    crea ninguna marca en incentives."""
    wb = make_workbook(tmp_path / "solo-bloque.xlsx", {
        "MIA Jun - Jack Mckarel": [
            HEADER,
            [D(2026, 5, 1), "SOLD", "FREE", "Descuentos",
             None, None, None, None, None, None, None],
            ["Alexandra Jimenez", 60, 10, "In Brine", 88.68,
             None, None, None, None, None, None],
        ],
    })
    data = sx.collect([wb])
    assert data.incentives == {}
    assert "sin filas de pedido" in capsys.readouterr().out
