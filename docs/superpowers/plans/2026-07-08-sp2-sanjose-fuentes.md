# SP2 — San José al modelo `fuentes/` — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Migrar el histórico de San José al formato canónico `fuentes/` y refactorizar `generar_data.py` para leerlo, **conservando exactamente los números actuales** y sumando junio 2026, y engancharlo al `split_excel.py`.

**Architecture:** Un migrador único (`migrar_historico.py`) reordena el Excel histórico viejo (13 columnas) a un solo `.xlsx` canónico (11 columnas). El `generar_data.py` cambia **solo `load_rows()`** para leer todos los `fuentes/*.xlsx`, filtrar filas de pedido reales, derivar estado desde `Company` y sacar el producto del `[SKU]` — conservando la regla de montos y toda la agregación/serialización. Se marca `GET` como `sp1_ready=True` en `split_excel.py`. Una prueba de regresión garantiza que los meses históricos no cambian.

**Tech Stack:** Python 3.14, openpyxl 3.1.5, pytest 9.1.1.

**Spec:** [docs/superpowers/specs/2026-07-08-sp2-sanjose-fuentes-design.md](../specs/2026-07-08-sp2-sanjose-fuentes-design.md)

---

## Estructura de archivos

| Archivo | Rol |
|---|---|
| `reporte-sanjose/migrar_historico.py` (crear) | Migración única: Historico (13 col) → `fuentes/San-Jose-historico-hasta-2026-05.xlsx` (11 col canónico). |
| `reporte-sanjose/generar_data.py` (modificar) | `load_rows()` nuevo (lee `fuentes/`, estado desde `Company`) + constantes/mensajes; el resto igual. |
| `split_excel.py` (modificar) | `BRANDS["GET"].sp1_ready = True`. |
| `reporte-sanjose/tests/conftest.py` (crear) | `sys.path` para importar `generar_data`/`migrar_historico`. |
| `reporte-sanjose/tests/sjhelpers.py` (crear) | Constructores de workbooks sintéticos (histórico y canónico). Nombre único para no chocar con `tests/helpers.py` de la raíz. |
| `reporte-sanjose/tests/reference_pre_sp2.json` (crear) | Snapshot de `MONTHS/D/GENERAL_DATA` del `data.js` actual (referencia de regresión). |
| `reporte-sanjose/tests/test_migrar_historico.py` (crear) | Prueba del migrador. |
| `reporte-sanjose/tests/test_load_rows.py` (crear) | Prueba de `load_rows` sobre canónico. |
| `reporte-sanjose/tests/test_regresion.py` (crear) | Reproduce idénticos los meses históricos (skip si falta el histórico migrado). |
| `tests/test_main.py` (modificar) | Ajustar el test de SP1 tras marcar `GET` como listo. |
| `reporte-sanjose/README.md` (modificar) | Documentar el flujo `fuentes/`. |
| `reporte-sanjose/data.js` (regenerado) | Ahora Oct 2025 – Jun 2026. |

**Nota de pytest:** cada carpeta de tests tiene su `conftest.py`. Los helpers de San José van en `sjhelpers.py` (nombre único) para evitar la colisión de `import file mismatch` que causaría un segundo `helpers.py`.

---

## Task 1: Scaffolding de tests + capturar la referencia de regresión

**Files:**
- Create: `reporte-sanjose/tests/conftest.py`
- Create: `reporte-sanjose/tests/sjhelpers.py`
- Create: `reporte-sanjose/tests/reference_pre_sp2.json`

- [ ] **Step 1: Crear `reporte-sanjose/tests/conftest.py`**

```python
import sys
from pathlib import Path

# Hace importable generar_data.py y migrar_historico.py (viven en el padre de tests/)
SANJOSE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SANJOSE))
```

- [ ] **Step 2: Crear `reporte-sanjose/tests/sjhelpers.py`**

```python
"""Constructores de workbooks sintéticos para los tests de San José."""
import datetime as _dt

import openpyxl

D = _dt.datetime

# Esquema viejo del Historico (13 columnas)
HIST_HEADER = [
    "Company", "State", "Customer", "Order", "Order Date", "Cod Prod", "Product",
    "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Salesperson", "Total", "Unit Price",
]

# Formato canónico (11 columnas)
CANON_HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson", "Company",
    "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]


def _mkwb(path, header, rows, sheet):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(header)
    for r in rows:
        ws.append(list(r) + [None] * (len(header) - len(r)))
    wb.save(path)
    return path


def make_historico(path, rows, sheet="Historico"):
    return _mkwb(path, HIST_HEADER, rows, sheet)


def make_canonical(path, rows, sheet="San Jose"):
    return _mkwb(path, CANON_HEADER, rows, sheet)


def hist_row(company, state, customer, order, date, cod, product,
             qd, salesperson, total, unit_price, qi=None, qo=None):
    qi = qd if qi is None else qi
    qo = qd if qo is None else qo
    return [company, state, customer, order, date, cod, product,
            qd, qi, qo, salesperson, total, unit_price]


def canon_row(date, order, variant, customer, salesperson, company,
              qd, unit_price, total, qi=None, qo=None):
    qi = qd if qi is None else qi
    qo = qd if qo is None else qo
    return [date, order, variant, customer, salesperson, company,
            qd, qi, qo, unit_price, total]
```

- [ ] **Step 3: Capturar la referencia desde el `data.js` ACTUAL (antes de tocar nada)**

Run (una sola línea):
```bash
python -c "import re,json; t=open('reporte-sanjose/data.js',encoding='utf-8').read(); g=lambda n,pat: json.loads(re.search(r'\b'+n+r' = ('+pat+r');',t,re.S).group(1)); ref={'MONTHS':g('MONTHS',r'\[.*?\]'),'D':g('D',r'\{.*?\}'),'GENERAL_DATA':g('GENERAL_DATA',r'\{.*\}')}; open('reporte-sanjose/tests/reference_pre_sp2.json','w',encoding='utf-8').write(json.dumps(ref,ensure_ascii=False,indent=2)); print('meses ref:', ref['MONTHS'])"
```
Expected: imprime `meses ref: ['Oct 25', 'Nov 25', 'Dic 25', 'Ene 26', 'Feb 26', 'Mar 26', 'Abr 26', 'May 26']` y crea el fixture.

- [ ] **Step 4: Verificar el fixture**

Run: `python -c "import json; d=json.load(open('reporte-sanjose/tests/reference_pre_sp2.json',encoding='utf-8')); print(len(d['MONTHS']),'meses;','ALL' in d['GENERAL_DATA'], len(d['D']['fl']))"`
Expected: `8 meses; True 8`

- [ ] **Step 5: Commit**

```bash
git add reporte-sanjose/tests/conftest.py reporte-sanjose/tests/sjhelpers.py reporte-sanjose/tests/reference_pre_sp2.json
git commit -m "test: scaffolding San José + referencia de regresión pre-SP2"
```

---

## Task 2: Migrador único `migrar_historico.py`

**Files:**
- Create: `reporte-sanjose/migrar_historico.py`
- Test: `reporte-sanjose/tests/test_migrar_historico.py`

- [ ] **Step 1: Escribir el test que falla**

`reporte-sanjose/tests/test_migrar_historico.py`:
```python
import openpyxl

from sjhelpers import make_historico, hist_row, D, CANON_HEADER
import migrar_historico as mig


def test_migrate_reorders_to_canonical(tmp_path):
    src = make_historico(tmp_path / "hist.xlsx", [
        hist_row("LatinFood Florida", "Florida", "Cliente A", "S100",
                 D(2025, 10, 9), "[GET01]", "Jack Mackerel in Brine",
                 4, "Alexandra J", 182.4, 45.6),
        hist_row("LatinFood US Corp", "Nueva York", "Cliente B", "S200",
                 D(2026, 1, 5), "[GET02]", "Jack Mackerel in Tomato Sauce",
                 1, "Katherine", 0, 0),  # promo
    ])
    out = tmp_path / "fuentes" / "San-Jose-historico.xlsx"
    n = mig.migrate(src=src, out=out)
    assert n == 0  # ninguna fila sin Cod Prod

    ws = openpyxl.load_workbook(out).active
    assert ws.title == "San Jose"
    grid = list(ws.iter_rows(values_only=True))
    assert list(grid[0]) == CANON_HEADER
    # fila 1: Product Variant reconstruido, Company preservado, sin columna State
    r = grid[1]
    assert r[0] == D(2025, 10, 9)                 # Order Date
    assert r[1] == "S100"                          # Order
    assert r[2] == "[GET01] Jack Mackerel in Brine"  # Product Variant
    assert r[3] == "Cliente A"                      # Customer
    assert r[4] == "Alexandra J"                    # Salesperson
    assert r[5] == "LatinFood Florida"             # Company
    assert r[9] == 45.6                             # Unit Price
    assert r[10] == 182.4                           # Total
    assert len(grid) == 3                           # header + 2 filas


def test_migrate_counts_missing_cod(tmp_path):
    src = make_historico(tmp_path / "hist.xlsx", [
        hist_row("LatinFood Florida", "Florida", "C", "S1",
                 D(2025, 10, 9), None, "Producto sin cod",
                 1, "SP", 45.6, 45.6),
    ])
    out = tmp_path / "fuentes" / "h.xlsx"
    assert mig.migrate(src=src, out=out) == 1
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest reporte-sanjose/tests/test_migrar_historico.py -v`
Expected: FAIL con `ModuleNotFoundError: No module named 'migrar_historico'`

- [ ] **Step 3: Crear `reporte-sanjose/migrar_historico.py`**

```python
"""
migrar_historico.py — Migración ÚNICA del histórico de San José al formato canónico.

Lee 'Historico Ventas Latin Food - San Jose.xlsx' (esquema viejo, 13 columnas) y
escribe 'fuentes/San-Jose-historico-hasta-2026-05.xlsx' (formato canónico, 11 columnas),
para que el generar_data.py nuevo lo lea igual que cualquier otro mes.

Fiel: los montos se copian tal cual (mayo 2026 viene con Total/Unit Price intercambiados;
eso lo resuelve el parser, no este migrador). Se descarta la columna State (redundante
con Company). El Historico original queda como respaldo.

Uso: python migrar_historico.py
"""

from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
SRC = HERE / "Historico Ventas Latin Food - San Jose.xlsx"
OUT = HERE / "fuentes" / "San-Jose-historico-hasta-2026-05.xlsx"
SHEET = "San Jose"

CANON_HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson", "Company",
    "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]


def migrate(src=SRC, out=OUT):
    """Reordena el Historico al formato canónico. Devuelve # de filas sin Cod Prod."""
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb["Historico"]
    out_wb = openpyxl.Workbook()
    ows = out_wb.active
    ows.title = SHEET
    ows.append(CANON_HEADER)

    missing_cod = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        (company, state, customer, order, date, cod_prod, product,
         qd, qi, qo, salesperson, total, unit_price) = row
        if date is None:
            continue
        if not cod_prod:
            missing_cod += 1
        variant = f"{cod_prod or ''} {product or ''}".strip()
        ows.append([date, order, variant, customer, salesperson, company,
                    qd, qi, qo, unit_price, total])

    Path(out).parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(out)
    return missing_cod


def main():
    n = migrate()
    print(f"✓ histórico migrado → {OUT.relative_to(HERE)}")
    if n:
        print(f"  ⚠ {n} filas sin Cod Prod (Product Variant quedó sin SKU)")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest reporte-sanjose/tests/test_migrar_historico.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add reporte-sanjose/migrar_historico.py reporte-sanjose/tests/test_migrar_historico.py
git commit -m "feat: migrar_historico — Historico viejo → fuentes/ canónico"
```

---

## Task 3: Refactor de `load_rows()` en `generar_data.py`

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (constantes, helpers, `load_rows`, mensajes de `main`)
- Test: `reporte-sanjose/tests/test_load_rows.py`

- [ ] **Step 1: Escribir el test que falla**

`reporte-sanjose/tests/test_load_rows.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def _build_fuentes(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", [
        # venta real FL
        canon_row(D(2026, 6, 22), "S1", "[GET01] Jack Mackerel in Brine ",
                  "Key Food", "Alexandra J", "LatinFood Florida", 20, 45.6, 912),
        # venta real NY
        canon_row(D(2026, 6, 18), "S2", "[GET01] Jack Mackerel in Brine",
                  "Food Fair", "Katherine", "LatinFood US Corp.", 1, 21.85, 21.85),
        # caja promo (ambos montos 0)
        canon_row(D(2026, 6, 3), "S3", "[GET02] Jack Mackerel in Tomato Sauce",
                  "Price Choice", "Raquel L", "LatinFood Florida", 3, 0, 0),
        # entregado 0 → se omite
        canon_row(D(2026, 6, 3), "S4", "[GET02] Jack Mackerel in Tomato Sauce",
                  "Price Choice", "Raquel L", "LatinFood Florida", 0, 45.6, 0),
        # bloque de incentivos (se ignora): fila SOLD + fila con nombre suelto
        [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None],
        ["Alexandra J", None, None, None, None, None, None, None, None, None, None],
    ])
    return fuentes


def test_load_rows_filters_and_derives(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", _build_fuentes(tmp_path))
    rows = g.load_rows()
    # 3 filas reales (las dos ventas + la promo); qty0 y el bloque se descartan
    assert len(rows) == 3

    fl = [r for r in rows if r["state"] == "Florida"]
    ny = [r for r in rows if r["state"] == "Nueva York"]
    assert len(ny) == 1 and ny[0]["company"] == "LatinFood US Corp."
    assert len(fl) == 2

    # nombre de producto sin [SKU] y sin espacios sobrantes
    assert {r["product"] for r in rows} == {"Jack Mackerel in Brine",
                                            "Jack Mackerel in Tomato Sauce"}
    # promo detectada (ambos montos 0)
    promo = [r for r in rows if r["is_promo"]]
    assert len(promo) == 1 and promo[0]["qty"] == 3
    # regla de montos actual (max) conservada
    venta = [r for r in rows if r["order"] == "S1"][0]
    assert venta["line_rev"] == 912 and venta["box_price"] == 912 / 20
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest reporte-sanjose/tests/test_load_rows.py -v`
Expected: FAIL (por ahora `load_rows` no acepta `SOURCES_DIR` / lee el Excel viejo) — típicamente `AttributeError: module 'generar_data' has no attribute 'SOURCES_DIR'`.

- [ ] **Step 3: Refactorizar `generar_data.py`**

Reemplazar el bloque desde los imports hasta el final de `load_rows()` (líneas 9–79 actuales) por:

```python
import re
import json
from pathlib import Path
from collections import defaultdict
from datetime import datetime

import openpyxl

HERE = Path(__file__).parent
SOURCES_DIR = HERE / "fuentes"
OUTPUT_FILE = HERE / "data.js"

MES_CORTO = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
             7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
MES_LARGO = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',
             6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',
             10:'Octubre',11:'Noviembre',12:'Diciembre'}

# Formato canónico: Order Date · Order · Product Variant · Customer · Salesperson ·
# Company · Qty Delivered · Qty Invoiced · Qty Ordered · Unit Price · Total
SKU_RE = re.compile(r"^\s*\[[A-Z]{2,4}\d+\]\s*")
STATE_BY_COMPANY = {
    "latinfood florida": "Florida",
    "latinfood us corp.": "Nueva York",
    "latinfood us corp": "Nueva York",
}


def month_key(date):
    return (date.year, date.month)

def month_label_short(ym):
    y, m = ym
    return f"{MES_CORTO[m]} {str(y)[2:]}"

def month_label_long(ym):
    y, m = ym
    return f"{MES_LARGO[m]} {y}"

def month_label_btn(ym):
    y, m = ym
    return f"{MES_CORTO[m]} {y}"


def _source_files():
    if not SOURCES_DIR.exists():
        return []
    return sorted(p for p in SOURCES_DIR.glob("*.xlsx") if not p.name.startswith("~$"))


def _state_from_company(company):
    if not isinstance(company, str):
        return None
    return STATE_BY_COMPANY.get(re.sub(r"\s+", " ", company).strip().lower())


def _is_tx(row):
    """Fila de pedido real: datetime en col A y SKU [XXX##] en col C (Product Variant)."""
    return (
        isinstance(row[0], datetime)
        and isinstance(row[2], str)
        and SKU_RE.match(row[2]) is not None
    )


def _product_name(variant):
    """'[GET01] Jack Mackerel in Brine ' -> 'Jack Mackerel in Brine'."""
    return SKU_RE.sub("", str(variant)).strip()


def load_rows():
    """Lee todos los fuentes/*.xlsx (formato canónico) y devuelve las filas de pedido
    limpias. Ignora el bloque de incentivos, subtotales y separadores."""
    rows = []
    for path in _source_files():
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        for row in ws.iter_rows(values_only=True):
            if not _is_tx(row):
                continue
            date, order, variant, customer, salesperson, company = row[:6]
            qty_del, unit_price, total = row[6], row[9], row[10]

            state = _state_from_company(company)
            if not state:
                continue
            qty = qty_del or 0
            if qty <= 0:                       # entregado 0 → no es venta ni promo real
                continue

            # Regla de montos ACTUAL (se conserva en SP2; su arreglo es SP3):
            # el Excel mezcla dos orientaciones (mayo 2026 viene intercambiado),
            # line_total = max(Total, Unit Price).
            a = total or 0
            b = unit_price or 0
            is_promo = (a == 0 and b == 0)
            line_rev = max(a, b)
            box_price = line_rev / qty if qty > 0 else 0.0

            rows.append({
                'company': company or '',
                'state': state,
                'customer': (customer or '').strip(),
                'order': order or '',
                'date': date,
                'product': _product_name(variant),
                'qty': qty,
                'line_rev': line_rev,
                'box_price': box_price,
                'is_promo': is_promo,
                'salesperson': (salesperson or '').strip(),
            })
    return rows
```

- [ ] **Step 4: Actualizar los mensajes de `main()` que nombraban el Excel viejo**

En `generar_data.py`, dentro de `main()`, reemplazar:
```python
    print(f"Leyendo {EXCEL_FILE}...")
```
por:
```python
    print("Leyendo fuentes/...")
```

Y en el template del `js`, reemplazar la línea de comentario:
```python
// Fuente: {EXCEL_FILE}
```
por:
```python
// Fuente: fuentes/*.xlsx
```
(El resto de `main()` y todas las funciones `get_order_unit_prices`, `promo_box_price`, `build_monthly_state_data`, `build_period_data`, `find_product_first_months` quedan **sin cambios**.)

- [ ] **Step 5: Correr y verificar que pasa**

Run: `python -m pytest reporte-sanjose/tests/test_load_rows.py -v`
Expected: PASS. Además `python -m pytest -q` (suite completa) sin regresiones nuevas.

- [ ] **Step 6: Commit**

```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_load_rows.py
git commit -m "feat: generar_data lee fuentes/ canónico (estado desde Company)"
```

---

## Task 4: Correr la migración real + prueba de regresión

**Files:**
- Test: `reporte-sanjose/tests/test_regresion.py`
- (Operacional) genera `reporte-sanjose/fuentes/San-Jose-historico-hasta-2026-05.xlsx`

- [ ] **Step 1: Correr la migración real**

Run: `python reporte-sanjose/migrar_historico.py`
Expected: `✓ histórico migrado → fuentes/San-Jose-historico-hasta-2026-05.xlsx` (sin aviso de Cod Prod).

- [ ] **Step 2: Verificar el archivo migrado**

Run: `python -c "import openpyxl; ws=openpyxl.load_workbook('reporte-sanjose/fuentes/San-Jose-historico-hasta-2026-05.xlsx').active; print('hoja',ws.title,'filas',ws.max_row)"`
Expected: `hoja San Jose filas 212` (encabezado + 211 filas de dato).

- [ ] **Step 3: Escribir la prueba de regresión**

`reporte-sanjose/tests/test_regresion.py`:
```python
import json
import re
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


@pytest.mark.skipif(not HIST.exists(),
                    reason="falta el histórico migrado (correr migrar_historico.py)")
def test_regresion_reproduce_historico(tmp_path, monkeypatch):
    # Genera data.js en un temporal, leyendo el fuentes/ REAL (histórico + junio)
    monkeypatch.setattr(g, "OUTPUT_FILE", tmp_path / "data.js")
    g.main()
    out = _parse_data_js((tmp_path / "data.js").read_text(encoding="utf-8"))

    # Los meses históricos deben reproducirse IDÉNTICOS
    for m in HIST_MONTHS:
        assert out["GENERAL_DATA"][m] == REF["GENERAL_DATA"][m], f"cambió {m}"

    # D["fl"]/D["ny"] son dicts de arrays por métrica (rev, cajas, pv, pc, ord),
    # una entrada por mes. Los primeros 8 valores (meses históricos) deben ser idénticos.
    for estado in ("fl", "ny"):
        for metrica, serie_ref in REF["D"][estado].items():
            assert out["D"][estado][metrica][:8] == serie_ref, \
                f"D.{estado}.{metrica} histórico cambió"

    # Y junio 2026 debe aparecer como mes nuevo (novena entrada en cada serie)
    assert "Jun 26" in out["MONTHS"]
    assert len(out["D"]["fl"]["rev"]) == 9
```

- [ ] **Step 4: Correr la regresión**

Run: `python -m pytest reporte-sanjose/tests/test_regresion.py -v`
Expected: PASS. Si falla en algún mes histórico, es bug de migración/refactor — investigar (NO cambiar la referencia).

- [ ] **Step 5: Commit**

```bash
git add reporte-sanjose/tests/test_regresion.py
git commit -m "test: regresión — meses históricos idénticos + junio agregado"
```

---

## Task 5: Enganchar San José al `split_excel.py`

**Files:**
- Modify: `split_excel.py:40` (`BRANDS["GET"]` → `sp1_ready=True`)
- Modify: `tests/test_main.py` (ajustar el test de SP1)

- [ ] **Step 1: Marcar GET como listo en `split_excel.py`**

Reemplazar la línea de `BRANDS`:
```python
    "GET": Brand("San José", "San-Jose", "reporte-sanjose", "generar_data.py", "San Jose", False),
```
por:
```python
    "GET": Brand("San José", "San-Jose", "reporte-sanjose", "generar_data.py", "San Jose", True),
```

- [ ] **Step 2: Ajustar `tests/test_main.py`** — al marcar GET como listo, ya no imprime "pendiente"; se agrega una hoja KOM (que sigue sin estar lista) para conservar esa aserción, y se afirma que GET quedó `sp1_ready`.

En `tests/test_main.py`, dentro de `_seed_entrada`, agregar una hoja KOM al dict `sheets` (KOM sigue con `sp1_ready=False`, así que mantiene viva la aserción "pendiente"):
```python
        "NY Jun - KOM": [
            HEADER,
            tx(D(2026, 6, 18), "S44804", "[KOM01] x", "Store", "Mireya",
               qd=1, price=50.64, total=50.64),
        ],
```

Y agregar un test nuevo al final del archivo:
```python
def test_get_is_wired_as_ready():
    assert sx.BRANDS["GET"].sp1_ready is True
```

- [ ] **Step 3: Correr los tests afectados**

Run: `python -m pytest tests/test_main.py -v`
Expected: PASS (incluye `test_main_writes_canonical_files` con "pendiente" ahora proveniente de KOM, y `test_get_is_wired_as_ready`).

- [ ] **Step 4: Suite completa**

Run: `python -m pytest -q`
Expected: todo verde.

- [ ] **Step 5: Commit**

```bash
git add split_excel.py tests/test_main.py
git commit -m "feat: enganchar San José al split (GET sp1_ready=True)"
```

---

## Task 6: Aceptación — regenerar San José con el pipeline real

**Files:**
- Verify + commit: `reporte-sanjose/data.js`

- [ ] **Step 1: Regenerar todo con el split real**

Run: `python split_excel.py`
Expected (resumen): ahora **GET** imprime `↻ data.js regenerado ✓` (además de GOA). KOM/ROB siguen ⏸. Sin SKUs no mapeados.

- [ ] **Step 2: Verificar que San José abarca Oct 2025 – Jun 2026**

Run: `python -c "import re,json; t=open('reporte-sanjose/data.js',encoding='utf-8').read(); m=json.loads(re.search(r'MONTHS = (\[.*?\]);',t,re.S).group(1)); print(m)"`
Expected: `['Oct 25', 'Nov 25', 'Dic 25', 'Ene 26', 'Feb 26', 'Mar 26', 'Abr 26', 'May 26', 'Jun 26']`

- [ ] **Step 3: Verificar que el diff de `data.js` NO cambia los meses históricos**

Run: `python -m pytest reporte-sanjose/tests/test_regresion.py -q`
Expected: PASS (los históricos idénticos, junio agregado).

- [ ] **Step 4: Verificar que git no trackea ningún .xlsx nuevo**

Run: `git status --porcelain --untracked-files=all | grep -i '.xlsx' || echo "OK: ningún xlsx en git"`
Expected: `OK: ningún xlsx en git`

- [ ] **Step 5: Commit del `data.js` regenerado**

```bash
git add reporte-sanjose/data.js
git commit -m "data: San José regenerado desde fuentes/ — Oct 2025 a Jun 2026"
```

---

## Task 7: README de San José + suite final

**Files:**
- Modify: `reporte-sanjose/README.md`

- [ ] **Step 1: Reemplazar la sección "Uso mensual" del `reporte-sanjose/README.md`**

Buscar el bloque actual:
```
## Uso mensual

1. Agregar los datos al Excel (`Historico Ventas Latin Food - San Jose.xlsx`).
2. Ejecutar el generador de datos:
   ```bash
   python generar_data.py
   ```
3. Subir el `data.js` actualizado al repositorio.
4. El sitio se actualiza automáticamente en GitHub Pages.
```

y sustituirlo por:
```
## Uso mensual

San José ahora se alimenta del pipeline central: el Excel mensual del distribuidor se
separa por marca con `split_excel.py` (en la raíz), que deja
`fuentes/San-Jose-<AAAA-MM>.xlsx` (formato canónico) y regenera este `data.js`.

1. Desde la raíz del repo, deja el Excel del distribuidor en `entrada/` y corre
   `python split_excel.py` (regenera San José y las demás marcas listas).
2. O, para regenerar solo San José desde sus `fuentes/`: `python generar_data.py`.
3. Sube el `data.js` actualizado y el sitio se actualiza en GitHub Pages.

El histórico Oct 2025 – May 2026 vive migrado en
`fuentes/San-Jose-historico-hasta-2026-05.xlsx` (generado una vez con
`migrar_historico.py`). El Excel `Historico Ventas Latin Food - San Jose.xlsx` original
queda solo como respaldo.

> El estado (Florida / Nueva York) se deriva de la columna `Company`.
```

Y en la tabla de "Archivos", agregar filas para `migrar_historico.py` y `fuentes/`, y marcar el Historico como respaldo. (Editar la tabla existente de forma coherente con su formato.)

- [ ] **Step 2: Verificar el README**

Run: `python -c "t=open('reporte-sanjose/README.md',encoding='utf-8').read(); print('split_excel' in t, t.count('\`\`\`')%2==0)"`
Expected: `True True`

- [ ] **Step 3: Suite completa final**

Run: `python -m pytest -q`
Expected: todo verde.

- [ ] **Step 4: Commit**

```bash
git add reporte-sanjose/README.md
git commit -m "docs: README de San José con el flujo fuentes/ + split_excel"
```

---

## Verificación final

- [ ] `python -m pytest -q` → todo verde (incluye regresión de San José).
- [ ] `reporte-sanjose/data.js` abarca Oct 2025 – Jun 2026; meses históricos idénticos a la referencia.
- [ ] `python split_excel.py` regenera GOA **y** San José; KOM/ROB siguen pendientes.
- [ ] Ningún `.xlsx` trackeado por git; el `Historico...xlsx` sigue de respaldo (gitignoreado).

## Fuera de alcance (recordatorio)

- **SP3:** arreglar la regla `max(Total, Unit Price)`, incorporar el bloque de incentivos
  (SOLD/FREE/Descuentos/NEW CUSTOMERS), revisar promos in-line, rediseñar el dashboard.
- **SP4:** migración del repo al GitHub de GetUp + subdominio privado.
