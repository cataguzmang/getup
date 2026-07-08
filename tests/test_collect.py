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
    block = data.incentives["GET"]["2026-06"]
    assert block and block[0][1] == "SOLD"
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
    assert data.rows == {} or all(not v for v in data.rows.values())
    assert "No se pudo leer" in capsys.readouterr().out
