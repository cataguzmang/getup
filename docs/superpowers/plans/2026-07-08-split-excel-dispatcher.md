# SP1 — `split_excel.py` Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Un script en la raíz que separa el Excel "sucio" del distribuidor en un `.xlsx` canónico por marca dentro de `fuentes/`, y regenera el `data.js` de las marcas listas (en SP1: solo GOA).

**Architecture:** Dos capas. Capa 1 = `split_excel.py` (dispatcher): lee `entrada/*.xlsx`, clasifica cada fila de pedido por prefijo de SKU, agrupa por marca y mes, y escribe `<marca>/fuentes/<code>-<YYYY-MM>.xlsx` con passthrough fiel de las celdas. Capa 2 = el generador propio de cada marca (ya existente en GOA), que el dispatcher invoca por `subprocess` al final. El estado se deduce de `Company` solo para el resumen; no se agrega como columna.

**Tech Stack:** Python 3.14, openpyxl 3.1.5, pytest 9.1.1, `argparse`, `subprocess`, `dataclasses`.

**Spec:** [docs/superpowers/specs/2026-07-08-split-excel-dispatcher-design.md](../specs/2026-07-08-split-excel-dispatcher-design.md)

---

## Estructura de archivos

| Archivo | Rol |
|---|---|
| `split_excel.py` (crear) | Dispatcher: config de marcas, helpers puros, `collect`, `write_canonical`, `run_generator`, `main`. |
| `tests/conftest.py` (crear) | `sys.path` para importar `split_excel`; fixture `messy_wb`. |
| `tests/helpers.py` (crear) | Constructores de workbooks sintéticos (`HEADER`, `tx`, `subtotal`, `make_workbook`). |
| `tests/test_helpers.py` (crear) | Tests de `sku_prefix`, `is_transaction_row`, `month_key`, `state_of`. |
| `tests/test_incentives.py` (crear) | Tests de `find_incentive_block`. |
| `tests/test_collect.py` (crear) | Tests de `collect` (agrupación, unmatched, incentivos, stats). |
| `tests/test_write_canonical.py` (crear) | Tests de `write_canonical` (header, hoja, separador, passthrough). |
| `tests/test_cli.py` (crear) | Tests de `parse_args` e `iter_input_files`. |
| `tests/test_main.py` (crear) | Integración de `main` con `ROOT` monkeypatched. |
| `tests/test_goa_compat.py` (crear) | Integración: el `GOA-*.xlsx` emitido lo lee el `parse_excel.py` de GOA. |
| `.gitignore` (crear) | Ignora `*.xlsx`, `entrada/`, `unmatched/`, caches. |
| `README.md` (modificar) | Documenta el flujo `entrada/ → split_excel.py → fuentes/`. |

---

## Task 1: Scaffolding — `.gitignore` y helpers de test

**Files:**
- Create: `.gitignore`
- Create: `tests/helpers.py`
- Create: `tests/conftest.py`

- [ ] **Step 1: Crear `.gitignore` en la raíz**

```gitignore
# Datos comerciales — nunca subir al repo
*.xlsx
*.xls
entrada/
unmatched/

# Python
__pycache__/
*.pyc
.pytest_cache/
```

- [ ] **Step 2: Crear `tests/helpers.py`**

```python
"""Constructores de workbooks sintéticos para los tests de split_excel."""
import datetime as _dt

import openpyxl

D = _dt.datetime

HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson",
    "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]


def _pad(row):
    """Rellena la fila a 11 columnas con None (como hace openpyxl al leer)."""
    return list(row) + [None] * (len(HEADER) - len(row))


def tx(date, order, variant, customer, sp, company="LatinFood US Corp.",
       qd=1, qi=1, qo=1, price=21.85, total=21.85):
    """Fila de pedido real."""
    return [date, order, variant, customer, sp, company, qd, qi, qo, price, total]


def subtotal(label, qd=1, qi=1, qo=1, price=21.85, total=21.85):
    """Fila-subtotal (texto en col A, sin fecha)."""
    return [label, None, None, None, None, None, qd, qi, qo, price, total]


def make_workbook(path, sheets):
    """`sheets`: dict {nombre_hoja: [fila, ...]}. Devuelve `path`."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        for r in rows:
            ws.append(_pad(r))
    wb.save(path)
    return path
```

- [ ] **Step 3: Crear `tests/conftest.py`**

```python
import sys
from pathlib import Path

# Hacer importable split_excel.py (vive en la raíz, padre de tests/)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest

from helpers import HEADER, tx, subtotal, make_workbook, D


@pytest.fixture
def messy_wb(tmp_path):
    """Workbook sucio de ejemplo: 4 marcas (GOA, GET, KOM, XYZ desconocida),
    dos estados, filas-subtotal y un bloque de incentivos en la hoja de MIA."""
    sheets = {
        "NY Jun - GOA": [
            HEADER,
            subtotal("Katherine Osorio Duque (6)", qd=5, qi=5, qo=5, price=25.8, total=109.25),
            tx(D(2026, 6, 24), "S45227", "[GOA02] Pure Chamoline Herbal Tea",
               "Food Bazaar # 14 -West NY", "Katherine Osorio Duque"),
            tx(D(2026, 6, 17), "S44722", "[GOA03] Ginger Lemongrass Herbal Tea",
               "Food Bazaar # 14 -West NY", "Katherine Osorio Duque"),
        ],
        "MIA Jun - Jack Mckarel": [
            HEADER,
            subtotal("[GET01] Jack Mackerel in Brine (8)", qd=68, qi=68, qo=68, price=39.9, total=2964),
            subtotal("    Alexandra J (2)", qd=35, qi=35, qo=35, price=45.6, total=1596),
            tx(D(2026, 6, 22), "S44946", "[GET01] Jack Mackerel in Brine",
               "Key Food 10400", "Alexandra J", company="LatinFood Florida",
               qd=20, qi=20, qo=20, price=45.6, total=912),
            # bloque de incentivos (col A con fecha de mayo, col B == 'SOLD')
            [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None],
            ["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68,
             None, None, None, None, None, None],
        ],
        "NY Jun - KOM": [
            HEADER,
            tx(D(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
               "Convenience store flora llc", "Mireya Fernandez",
               qd=0.5, qi=0.5, qo=0.5, price=50.64, total=25.32),
        ],
        "NY Jun - XYZ": [
            HEADER,
            tx(D(2026, 6, 10), "S99999", "[XYZ01] Mystery Product", "Some Store", "Someone"),
        ],
    }
    return make_workbook(tmp_path / "GET Jun 2026.xlsx", sheets)
```

- [ ] **Step 4: Verificar que los helpers importan sin error**

Run: `python -c "import sys; sys.path.insert(0,'tests'); import helpers; print(len(helpers.HEADER))"`
Expected: imprime `11`

- [ ] **Step 5: Commit**

```bash
git add .gitignore tests/helpers.py tests/conftest.py
git commit -m "chore: scaffolding de tests y .gitignore para split_excel"
```

---

## Task 2: Helpers puros — clasificación, mes y estado

**Files:**
- Create: `split_excel.py`
- Test: `tests/test_helpers.py`

- [ ] **Step 1: Escribir los tests que fallan**

`tests/test_helpers.py`:
```python
import datetime

from helpers import tx, subtotal, HEADER
import split_excel as sx

D = datetime.datetime


def test_sku_prefix():
    assert sx.sku_prefix("[GOA02] Pure Chamoline Herbal Tea") == "GOA"
    assert sx.sku_prefix("[GET01] Jack Mackerel in Brine") == "GET"
    assert sx.sku_prefix("  [ROB05] Mussel en Brine") == "ROB"
    assert sx.sku_prefix("Katherine Osorio Duque (6)") is None
    assert sx.sku_prefix(None) is None


def test_is_transaction_row():
    good = tx(D(2026, 6, 24), "S45227", "[GOA02] x", "cliente", "sp")
    assert sx.is_transaction_row(good) is True
    # subtotal por vendedor
    assert sx.is_transaction_row(subtotal("Katherine (6)")) is False
    # subtotal por producto de MIA (texto con SKU en col A, sin fecha)
    assert sx.is_transaction_row(subtotal("[GET01] Jack Mackerel (8)")) is False
    # encabezado del bloque de incentivos
    inc = [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None]
    assert sx.is_transaction_row(inc) is False


def test_month_key():
    assert sx.month_key(D(2026, 6, 24, 11, 4)) == "2026-06"
    assert sx.month_key(D(2026, 12, 1)) == "2026-12"


def test_state_of():
    fl = tx(D(2026, 6, 1), "S1", "[GET01] x", "c", "sp", company="LatinFood Florida")
    ny = tx(D(2026, 6, 1), "S2", "[GET01] x", "c", "sp", company="LatinFood US Corp.")
    assert sx.state_of(fl, "MIA Jun - Jack Mckarel") == "Florida"
    assert sx.state_of(ny, "NY Jun - GOA") == "Nueva York"
    # Company vacío → respaldo por prefijo de hoja
    blank = tx(D(2026, 6, 1), "S3", "[GET01] x", "c", "sp", company=None)
    assert sx.state_of(blank, "MIA Jun - Jack Mckarel") == "Florida"
    # desconocido
    unk = tx(D(2026, 6, 1), "S4", "[GET01] x", "c", "sp", company="Otra Cosa")
    assert sx.state_of(unk, "?? Jun") == "?"
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest tests/test_helpers.py -v`
Expected: FAIL con `ModuleNotFoundError: No module named 'split_excel'`

- [ ] **Step 3: Crear `split_excel.py` con config y helpers puros**

```python
"""
split_excel.py — Dispatcher de reparto por marca (SP1)

Lee el/los Excel "sucios" del distribuidor en entrada/ (una hoja por estado+marca),
separa cada fila de pedido por marca según el prefijo del SKU, y deposita un .xlsx
canónico y limpio en fuentes/ de cada marca. Luego regenera el data.js de las marcas
cuyo generador ya entiende el formato canónico (por ahora: GOA).

Uso:
    python split_excel.py [--input entrada] [--no-build] [--only GOA,KOM]
"""

import argparse
import datetime
import re
import subprocess
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent


# ── Configuración de marcas ──────────────────────────────────────────────
@dataclass(frozen=True)
class Brand:
    name: str        # nombre legible
    code: str        # prefijo del nombre de archivo canónico
    folder: str      # carpeta destino (relativa a ROOT)
    generator: str   # script generador dentro de la carpeta
    sheet: str       # nombre de la hoja del archivo canónico
    sp1_ready: bool  # ¿su generador ya lee el formato canónico?


BRANDS = {
    "GET": Brand("San José", "San-Jose", "reporte-sanjose", "generar_data.py", "San Jose", False),
    "GOA": Brand("Garden of the Andes", "GOA", "goa-inventory", "parse_excel.py", "GOA", True),
    "KOM": Brand("Kombuchacha", "KOM", "kombuchacha", "build_data.py", "KOM", False),
    "ROB": Brand("Robinson Crusoe", "ROB", "robinson-crusoe", "build_data.py", "ROB", False),
}

HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson",
    "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]

# Índices de columna (0-based)
COL_DATE, COL_ORDER, COL_VARIANT, COL_CUSTOMER, COL_SP, COL_COMPANY = 0, 1, 2, 3, 4, 5

SKU_PREFIX_RE = re.compile(r"^\s*\[([A-Z]{2,4})\d+\]")

STATE_BY_COMPANY = {
    "latinfood florida": "Florida",
    "latinfood us corp.": "Nueva York",
    "latinfood us corp": "Nueva York",
}
STATE_BY_SHEET_PREFIX = {"MIA": "Florida", "NY": "Nueva York"}


# ── Helpers puros ─────────────────────────────────────────────────────────
def sku_prefix(variant):
    """'[GOA02] Pure Chamoline' -> 'GOA'. None si no matchea o no es str."""
    if not isinstance(variant, str):
        return None
    m = SKU_PREFIX_RE.match(variant)
    return m.group(1) if m else None


def is_transaction_row(row):
    """Fila de pedido real: datetime en col A y SKU [XXX##] en col C."""
    return (
        isinstance(row[COL_DATE], datetime.datetime)
        and sku_prefix(row[COL_VARIANT]) is not None
    )


def month_key(date):
    """datetime -> 'YYYY-MM'."""
    return f"{date.year:04d}-{date.month:02d}"


def state_of(row, sheet_name):
    """Estado desde Company; respaldo por prefijo de hoja; '?' si desconocido."""
    company = row[COL_COMPANY]
    if isinstance(company, str):
        key = re.sub(r"\s+", " ", company).strip().lower()
        if key in STATE_BY_COMPANY:
            return STATE_BY_COMPANY[key]
    parts = str(sheet_name).split()
    prefix = parts[0].upper() if parts else ""
    return STATE_BY_SHEET_PREFIX.get(prefix, "?")
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest tests/test_helpers.py -v`
Expected: PASS (5 tests... 4 funciones de test, todas verde)

- [ ] **Step 5: Commit**

```bash
git add split_excel.py tests/test_helpers.py
git commit -m "feat: helpers de clasificación, mes y estado en split_excel"
```

---

## Task 3: `find_incentive_block`

**Files:**
- Modify: `split_excel.py` (agregar `find_incentive_block`)
- Test: `tests/test_incentives.py`

- [ ] **Step 1: Escribir el test que falla**

`tests/test_incentives.py`:
```python
import datetime

import split_excel as sx

D = datetime.datetime


def _rows_with_block():
    return [
        list(sx.HEADER),
        [D(2026, 6, 22), "S44946", "[GET01] x", "c", "Alexandra J",
         "LatinFood Florida", 20, 20, 20, 45.6, 912],
        [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None],
        ["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68,
         None, None, None, None, None, None],
    ]


def test_find_incentive_block_returns_from_sold():
    block = sx.find_incentive_block(_rows_with_block())
    assert len(block) == 2
    assert block[0][1] == "SOLD"
    assert block[1][0] == "Alexandra Jimenez"


def test_find_incentive_block_absent():
    rows = [
        list(sx.HEADER),
        [D(2026, 6, 22), "S44946", "[GET01] x", "c", "sp",
         "LatinFood Florida", 20, 20, 20, 45.6, 912],
    ]
    assert sx.find_incentive_block(rows) == []
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest tests/test_incentives.py -v`
Expected: FAIL con `AttributeError: module 'split_excel' has no attribute 'find_incentive_block'`

- [ ] **Step 3: Agregar `find_incentive_block` a `split_excel.py`** (tras `state_of`)

```python
def find_incentive_block(rows):
    """Filas del bloque de incentivos (desde el ancla 'SOLD' al final), o [].

    Ancla: alguna celda de la fila == 'SOLD' (case-insensitive).
    """
    for i, row in enumerate(rows):
        if any(isinstance(c, str) and c.strip().upper() == "SOLD" for c in row):
            return [list(r) for r in rows[i:]]
    return []
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest tests/test_incentives.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add split_excel.py tests/test_incentives.py
git commit -m "feat: detección del bloque de incentivos"
```

---

## Task 4: `collect` — recolección y agrupación

**Files:**
- Modify: `split_excel.py` (agregar `Collected` y `collect`)
- Test: `tests/test_collect.py`

- [ ] **Step 1: Escribir el test que falla**

`tests/test_collect.py`:
```python
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
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest tests/test_collect.py -v`
Expected: FAIL con `AttributeError: module 'split_excel' has no attribute 'collect'`

- [ ] **Step 3: Agregar `Collected` y `collect` a `split_excel.py`** (tras `find_incentive_block`)

```python
@dataclass
class Collected:
    rows: dict         # code -> {month -> [row, ...]}
    incentives: dict   # code -> {month -> [block_row, ...]}
    unmatched: list    # [row, ...]
    stats: dict        # code -> {month -> Counter(state)}


def collect(input_files):
    """Recorre los workbooks y separa filas por marca/mes, guardando además los
    SKUs no mapeados, los bloques de incentivos y los conteos por estado."""
    rows = defaultdict(lambda: defaultdict(list))
    incentives = defaultdict(dict)
    unmatched = []
    stats = defaultdict(lambda: defaultdict(Counter))

    for path in input_files:
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:  # archivo corrupto / no es xlsx
            print(f"  ⚠ No se pudo leer {Path(path).name}: {e} (se omite)")
            continue

        for ws in wb.worksheets:
            sheet_rows = list(ws.iter_rows(values_only=True))
            tx_codes, tx_months = [], []

            for row in sheet_rows:
                if not is_transaction_row(row):
                    continue
                prefix = sku_prefix(row[COL_VARIANT])
                mk = month_key(row[COL_DATE])
                if prefix not in BRANDS:
                    unmatched.append(list(row))
                    continue
                rows[prefix][mk].append(list(row))
                stats[prefix][mk][state_of(row, ws.title)] += 1
                tx_codes.append(prefix)
                tx_months.append(mk)

            block = find_incentive_block(sheet_rows)
            if block and tx_codes:
                dom_code = Counter(tx_codes).most_common(1)[0][0]
                dom_month = Counter(tx_months).most_common(1)[0][0]
                prev = incentives[dom_code].get(dom_month, [])
                incentives[dom_code][dom_month] = prev + block
            elif block:
                print(f"  ⚠ Bloque de incentivos en '{ws.title}' sin filas de pedido; se omite")

    # Convertir defaultdicts a dicts normales para asserts predecibles
    return Collected(
        rows={c: dict(m) for c, m in rows.items()},
        incentives={c: dict(m) for c, m in incentives.items()},
        unmatched=unmatched,
        stats={c: {mk: cnt for mk, cnt in m.items()} for c, m in stats.items()},
    )
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest tests/test_collect.py -v`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add split_excel.py tests/test_collect.py
git commit -m "feat: collect — separa filas por marca/mes con unmatched e incentivos"
```

---

## Task 5: `write_canonical` — escritura del archivo canónico

**Files:**
- Modify: `split_excel.py` (agregar `write_canonical`)
- Test: `tests/test_write_canonical.py`

- [ ] **Step 1: Escribir el test que falla**

`tests/test_write_canonical.py`:
```python
import datetime

import openpyxl

import split_excel as sx

D = datetime.datetime


def test_write_canonical_header_rows_and_sheet(tmp_path):
    rows = [[D(2026, 6, 24), "S45227", "[GOA02] x", "c", "sp",
             "LatinFood US Corp.", 1, 1, 1, 21.85, 21.85]]
    out = tmp_path / "GOA-2026-06.xlsx"
    sx.write_canonical(out, rows, [], "GOA")

    wb = openpyxl.load_workbook(out)
    ws = wb["GOA"]
    grid = list(ws.iter_rows(values_only=True))
    assert list(grid[0]) == sx.HEADER
    assert grid[1][1] == "S45227"
    assert isinstance(grid[1][0], datetime.datetime)


def test_write_canonical_preserves_fractional_qty(tmp_path):
    rows = [[D(2026, 6, 18), "S44804", "[KOM01] x", "c", "sp",
             "LatinFood US Corp.", 0.5, 0.5, 0.5, 50.64, 25.32]]
    out = tmp_path / "KOM-2026-06.xlsx"
    sx.write_canonical(out, rows, [], "KOM")

    ws = openpyxl.load_workbook(out).active
    assert list(ws.iter_rows(values_only=True))[1][6] == 0.5


def test_write_canonical_appends_incentive_block(tmp_path):
    rows = [[D(2026, 6, 22), "S44946", "[GET01] x", "c", "sp",
             "LatinFood Florida", 20, 20, 20, 45.6, 912]]
    block = [["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68,
              None, None, None, None, None, None]]
    out = tmp_path / "San-Jose-2026-06.xlsx"
    sx.write_canonical(out, rows, block, "San Jose")

    grid = list(openpyxl.load_workbook(out).active.iter_rows(values_only=True))
    # header + 1 fila + separador en blanco + 1 fila de bloque
    assert len(grid) == 4
    assert all(c is None for c in grid[2])            # separador
    assert grid[3][0] == "Alexandra Jimenez"


def test_write_canonical_creates_missing_dirs(tmp_path):
    out = tmp_path / "nueva" / "fuentes" / "GOA-2026-06.xlsx"
    sx.write_canonical(out, [], [], "GOA")
    assert out.exists()
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest tests/test_write_canonical.py -v`
Expected: FAIL con `AttributeError: module 'split_excel' has no attribute 'write_canonical'`

- [ ] **Step 3: Agregar `write_canonical` a `split_excel.py`** (tras `collect`)

```python
def write_canonical(path, rows, incentive_block, sheet_name):
    """Escribe el .xlsx canónico: encabezado + filas de pedido (passthrough fiel)
    + (separador en blanco + bloque de incentivos, si lo hay)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(HEADER)
    for row in rows:
        ws.append(list(row))
    if incentive_block:
        ws.append([None] * len(HEADER))
        for row in incentive_block:
            ws.append(list(row))
    wb.save(path)
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest tests/test_write_canonical.py -v`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add split_excel.py tests/test_write_canonical.py
git commit -m "feat: write_canonical — archivo por marca con passthrough e incentivos"
```

---

## Task 6: CLI — `parse_args`, `iter_input_files`, `run_generator`

**Files:**
- Modify: `split_excel.py` (agregar las tres funciones)
- Test: `tests/test_cli.py`

- [ ] **Step 1: Escribir el test que falla**

`tests/test_cli.py`:
```python
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
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest tests/test_cli.py -v`
Expected: FAIL con `AttributeError: module 'split_excel' has no attribute 'parse_args'`

- [ ] **Step 3: Agregar las funciones a `split_excel.py`** (tras `write_canonical`)

```python
def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Reparte el Excel del distribuidor por marca.")
    p.add_argument("--input", default="entrada", help="Carpeta con los .xlsx sucios.")
    p.add_argument("--no-build", action="store_true", help="No regenerar ningún data.js.")
    p.add_argument("--only", default="", help="Limitar a códigos de marca (ej. GOA,KOM).")
    return p.parse_args(argv)


def iter_input_files(input_dir):
    """Todos los .xlsx de la carpeta (sin temporales ~$), en orden de nombre."""
    d = Path(input_dir)
    if not d.is_absolute():
        d = ROOT / d
    if not d.exists():
        return []
    return sorted(p for p in d.glob("*.xlsx") if not p.name.startswith("~$"))


def run_generator(brand):
    """Corre el generador de la marca. Devuelve (returncode|None, log)."""
    folder = ROOT / brand.folder
    script = folder / brand.generator
    if not script.exists():
        return None, f"generador ausente ({brand.generator})"
    proc = subprocess.run(
        [sys.executable, brand.generator],
        cwd=folder, capture_output=True, text=True,
    )
    return proc.returncode, proc.stdout + proc.stderr
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest tests/test_cli.py -v`
Expected: PASS (6 tests)

- [ ] **Step 5: Commit**

```bash
git add split_excel.py tests/test_cli.py
git commit -m "feat: CLI — parse_args, iter_input_files y run_generator"
```

---

## Task 7: `main` — orquestación y resumen

**Files:**
- Modify: `split_excel.py` (agregar `main` y el guard `__main__`)
- Test: `tests/test_main.py`

- [ ] **Step 1: Escribir el test que falla**

`tests/test_main.py`:
```python
import openpyxl

import split_excel as sx
from helpers import HEADER, tx, subtotal, make_workbook, D


def _seed_entrada(root):
    """Crea entrada/ con el workbook sucio bajo una ROOT temporal."""
    sheets = {
        "NY Jun - GOA": [
            HEADER,
            subtotal("Katherine (6)", qd=5),
            tx(D(2026, 6, 24), "S45227", "[GOA02] x", "Food Bazaar", "Katherine"),
            tx(D(2026, 6, 17), "S44722", "[GOA03] y", "Food Bazaar", "Katherine"),
        ],
        "MIA Jun - Jack Mckarel": [
            HEADER,
            tx(D(2026, 6, 22), "S44946", "[GET01] z", "Key Food", "Alexandra J",
               company="LatinFood Florida", qd=20, price=45.6, total=912),
            [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None],
            ["Alexandra Jimenez", 60, 10, "In Brine", 88.68, None, None, None, None, None, None],
        ],
        "NY Jun - XYZ": [
            HEADER,
            tx(D(2026, 6, 10), "S99999", "[XYZ01] mystery", "Store", "Someone"),
        ],
    }
    entrada = root / "entrada"
    entrada.mkdir()
    make_workbook(entrada / "GET Jun 2026.xlsx", sheets)


def test_main_writes_canonical_files(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    _seed_entrada(tmp_path)

    rc = sx.main(["--no-build"])
    assert rc == 0

    goa = tmp_path / "goa-inventory" / "fuentes" / "GOA-2026-06.xlsx"
    get = tmp_path / "reporte-sanjose" / "fuentes" / "San-Jose-2026-06.xlsx"
    assert goa.exists() and get.exists()

    ws = openpyxl.load_workbook(goa)["GOA"]
    assert list(ws.iter_rows(values_only=True))[0] == tuple(sx.HEADER)
    assert ws.max_row == 3  # header + 2 líneas

    out = capsys.readouterr().out
    assert "GOA" in out and "San José" in out
    assert "pendiente" in out  # GET no se buildea en SP1


def test_main_writes_unmatched(tmp_path, monkeypatch):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    _seed_entrada(tmp_path)
    sx.main(["--no-build"])
    assert (tmp_path / "unmatched" / "unmatched.xlsx").exists()


def test_main_no_input(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    (tmp_path / "entrada").mkdir()
    rc = sx.main(["--no-build"])
    assert rc == 1
    assert "No hay .xlsx" in capsys.readouterr().out


def test_main_only_filters_brands(tmp_path, monkeypatch):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    _seed_entrada(tmp_path)
    sx.main(["--no-build", "--only", "GOA"])
    assert (tmp_path / "goa-inventory" / "fuentes" / "GOA-2026-06.xlsx").exists()
    assert not (tmp_path / "reporte-sanjose" / "fuentes" / "San-Jose-2026-06.xlsx").exists()
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest tests/test_main.py -v`
Expected: FAIL con `AttributeError: module 'split_excel' has no attribute 'main'`

- [ ] **Step 3: Agregar `main` y el guard a `split_excel.py`** (al final del archivo)

```python
def main(argv=None):
    args = parse_args(argv)
    only = {c.strip().upper() for c in args.only.split(",") if c.strip()}

    files = iter_input_files(args.input)
    if not files:
        print(f"✗ No hay .xlsx en {args.input}/. Pon ahí el Excel del distribuidor.")
        return 1

    data = collect(files)
    print(f"✓ split_excel.py — {', '.join(f.name for f in files)}\n")

    for code, brand in BRANDS.items():
        if only and code not in only:
            continue
        months = data.rows.get(code, {})
        if not months:
            continue

        print(f"  {code}  ({brand.name})")
        for mk in sorted(months):
            month_rows = months[mk]
            block = data.incentives.get(code, {}).get(mk, [])
            out = ROOT / brand.folder / "fuentes" / f"{brand.code}-{mk}.xlsx"
            write_canonical(out, month_rows, block, brand.sheet)
            st = data.stats.get(code, {}).get(mk, {})
            st_txt = " ".join(f"{k}:{v}" for k, v in st.items())
            inc_txt = " · +incentivos" if block else ""
            rel = out.relative_to(ROOT).as_posix()
            print(f"    → {rel}   ({len(month_rows)} líneas · {st_txt}{inc_txt})")

        if args.no_build:
            print("    ⏸ build omitido (--no-build)")
        elif not brand.sp1_ready:
            print("    ⏸ data.js NO regenerado (pendiente su sub-proyecto)")
        else:
            rc, log = run_generator(brand)
            if rc == 0:
                print("    ↻ data.js regenerado ✓")
            else:
                print(f"    ✗ generador falló (rc={rc}):\n{log}")

    if data.unmatched:
        out = ROOT / "unmatched" / "unmatched.xlsx"
        write_canonical(out, data.unmatched, [], "unmatched")
        prefixes = sorted({sku_prefix(r[COL_VARIANT]) or "?" for r in data.unmatched})
        print(f"\n  ⚠ {len(data.unmatched)} líneas sin marca → unmatched/unmatched.xlsx  "
              f"(prefijos: {', '.join(prefixes)})")
    else:
        print("\n  ⚠ SKUs no mapeados: (ninguno)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest tests/test_main.py -v`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add split_excel.py tests/test_main.py
git commit -m "feat: main — orquestación, reparto y resumen"
```

---

## Task 8: Integración — el archivo GOA emitido lo lee el parser de GOA

**Files:**
- Test: `tests/test_goa_compat.py`

- [ ] **Step 1: Escribir el test**

`tests/test_goa_compat.py`:
```python
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
```

- [ ] **Step 2: Correr y verificar que pasa**

Run: `python -m pytest tests/test_goa_compat.py -v`
Expected: PASS (1 test). Si falla por el nombre de hoja, confirmar que `write_canonical` usó `"GOA"` (está en `BRANDS["GOA"].sheet`).

- [ ] **Step 3: Correr toda la batería**

Run: `python -m pytest -v`
Expected: PASS (todos los archivos de test en verde)

- [ ] **Step 4: Commit**

```bash
git add tests/test_goa_compat.py
git commit -m "test: integración — el GOA canónico lo lee parse_excel.py de GOA"
```

---

## Task 9: Aceptación con el Excel real de junio + regenerar GOA

**Files:**
- Move: `GET Jun 2026.xlsx` → `entrada/GET Jun 2026.xlsx`
- Verify: salidas reales + `goa-inventory/data.js`

- [ ] **Step 1: Mover el Excel real a `entrada/`**

```bash
mkdir -p entrada
git mv "GET Jun 2026.xlsx" "entrada/GET Jun 2026.xlsx" 2>/dev/null || mv "GET Jun 2026.xlsx" "entrada/GET Jun 2026.xlsx"
```
(Está sin trackear, así que `mv` normal basta; queda ignorado por el `.gitignore`.)

- [ ] **Step 2: Correr el split de verdad**

Run: `python split_excel.py`
Expected (resumen): GOA con 8 líneas y `↻ data.js regenerado ✓`; San José (`GET`) con 3 líneas y `+incentivos` y `⏸ pendiente`; KOM 2 líneas ⏸; ROB 1 línea ⏸; `SKUs no mapeados: (ninguno)`.

- [ ] **Step 3: Verificar los archivos canónicos creados**

Run: `python -c "import openpyxl,glob;[print(f, openpyxl.load_workbook(f).active.max_row) for f in sorted(glob.glob('*/fuentes/*-2026-06.xlsx'))]"`
Expected: lista con `goa-inventory/fuentes/GOA-2026-06.xlsx`, `kombuchacha/fuentes/KOM-2026-06.xlsx`, `reporte-sanjose/fuentes/San-Jose-2026-06.xlsx`, `robinson-crusoe/fuentes/ROB-2026-06.xlsx`.

- [ ] **Step 4: Verificar que GOA data.js incluye junio 2026**

Run: `python -c "import re;t=open('goa-inventory/data.js',encoding='utf-8').read();print('2026-06' in t, 'Junio 2026' in t)"`
Expected: `True True`

- [ ] **Step 5: Verificar que git NO trackea ningún .xlsx**

Run: `git status --porcelain --untracked-files=all | grep -i '.xlsx' || echo "OK: ningún xlsx en git"`
Expected: `OK: ningún xlsx en git`

- [ ] **Step 6: Commit (script, tests y data.js de GOA)**

```bash
git add split_excel.py goa-inventory/data.js
git commit -m "feat: split_excel operativo — GOA actualizado con junio 2026"
```

---

## Task 10: README raíz — documentar el flujo

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Reemplazar la sección "Cómo actualizar un dashboard" del `README.md`**

Buscar el bloque actual:
```
## Cómo actualizar un dashboard

Cada dashboard lee sus datos desde un archivo `data.js` generado por un script Python a partir de un Excel.

```
Excel → script Python → data.js → git push → GitHub Pages
```

Ver el README de cada carpeta para instrucciones específicas.
```

y sustituirlo por:
```
## Cómo actualizar los dashboards

El distribuidor manda un solo Excel mensual (`GET <Mes> <Año>.xlsx`) con una hoja por
estado + marca. El flujo es un comando:

1. Deja ese Excel en `entrada/` (carpeta privada, no se sube al repo).
2. Ejecuta el dispatcher:
   ```
   python split_excel.py
   ```
   - Separa cada marca a su `fuentes/<Marca>-<AAAA-MM>.xlsx` (formato canónico limpio).
   - Regenera el `data.js` de las marcas listas (hoy: GOA).
   - Los SKUs que no reconozca van a `unmatched/` con aviso.
3. Sube los `data.js` cambiados:
   ```
   git add */data.js && git commit -m "data: actualizar reportes" && git push
   ```

Mapeo de marcas (por prefijo de SKU): `GET` → San José · `GOA` → Garden of the Andes ·
`KOM` → Kombuchacha · `ROB` → Robinson Crusoe.

Cada dashboard conserva su propio generador (`fuentes/ → data.js`); `split_excel.py` los
llama al final. Ver el README de cada carpeta para el detalle de cada reporte.

> **Requisitos:** `pip install openpyxl`
```

- [ ] **Step 2: Verificar que el README renderiza sin bloques rotos**

Run: `python -c "t=open('README.md',encoding='utf-8').read();print('split_excel.py' in t, t.count('```')%2==0)"`
Expected: `True True`

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: README raíz con el flujo entrada/ → split_excel.py"
```

---

## Verificación final

- [ ] `python -m pytest -v` → todos los tests en verde.
- [ ] `python split_excel.py` sobre el junio real → resumen sin `unmatched`, GOA regenerado.
- [ ] `git status` limpio salvo los `data.js` esperados; ningún `.xlsx` trackeado.
- [ ] Los cuatro `fuentes/<Marca>-2026-06.xlsx` existen.

## Fuera de alcance (recordatorio)

- SP2: refactor de San José a `fuentes/` + migración del histórico.
- SP3: dashboards de KOM y ROB + `report_core.py`.
- SP4: migración a GitHub de GetUp + subdominio privado.
