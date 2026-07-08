import sys

import split_excel as sx


def test_parse_args_defaults():
    args = sx.parse_args([])
    assert args.input == "entrada"
    assert args.no_build is False
    assert args.only == ""


def test_parse_args_flags():
    args = sx.parse_args(["--input", "otra", "--no-build", "--only", "GOA,KOM"])
    assert args.input == "otra"
    assert args.no_build is True
    assert args.only == "GOA,KOM"


def test_iter_input_files_sorted_and_skips_temp(tmp_path):
    (tmp_path / "b.xlsx").write_text("x")
    (tmp_path / "a.xlsx").write_text("x")
    (tmp_path / "~$a.xlsx").write_text("x")   # temporal de Excel
    (tmp_path / "nota.txt").write_text("x")
    files = sx.iter_input_files(tmp_path)
    assert [f.name for f in files] == ["a.xlsx", "b.xlsx"]


def test_iter_input_files_missing_dir(tmp_path):
    assert sx.iter_input_files(tmp_path / "no-existe") == []


def test_run_generator_absent(tmp_path, monkeypatch):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    brand = sx.BRANDS["GOA"]
    (tmp_path / brand.folder).mkdir(parents=True)
    rc, log = sx.run_generator(brand)
    assert rc is None
    assert "ausente" in log


def test_run_generator_runs_script(tmp_path, monkeypatch):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    brand = sx.BRANDS["GOA"]
    folder = tmp_path / brand.folder
    folder.mkdir(parents=True)
    (folder / brand.generator).write_text("print('ok generador')\n")
    rc, log = sx.run_generator(brand)
    assert rc == 0
    assert "ok generador" in log
