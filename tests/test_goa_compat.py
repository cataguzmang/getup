import importlib.util
from pathlib import Path

import split_excel as sx

ROOT = Path(__file__).resolve().parent.parent


def _load_goa_parser():
    path = ROOT / "goa-inventory" / "parse_excel.py"
    spec = importlib.util.spec_from_file_location("goa_parse", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_emitted_goa_file_is_readable_by_goa_parser(messy_wb, tmp_path):
    data = sx.collect([messy_wb])
    out = tmp_path / "GOA-2026-06.xlsx"
    sx.write_canonical(out, data.rows["GOA"]["2026-06"], [], "GOA")

    goa = _load_goa_parser()
    rows = goa.read_sheet_rows(out)
    lines = goa.parse_transactions(rows)
    assert len(lines) == 2
    assert {ln["sku"] for ln in lines} == {"GOA02", "GOA03"}
    assert all(ln["month"] == "2026-06" for ln in lines)
