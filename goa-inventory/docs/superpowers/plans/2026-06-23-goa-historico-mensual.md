# Histórico mensual GOA — Plan de Implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ingerir el formato nuevo de Excel (multi-mes, hoja `Sheet1`, sin incentivos), acumular todos los Excel recibidos en un histórico a prueba de pérdidas, y mostrar en el dashboard un total general más un desglose por mes con un selector arriba.

**Architecture:** Se mantiene el pipeline `Excel → parse_excel.py → data.js → dashboard.html`. `parse_excel.py` pasa de leer un archivo fijo a leer todos los `*.xlsx` de la carpeta `fuentes/`, dedup por número de orden, etiquetar cada línea con su mes, y producir una vista general + `months[]`. El carry-forward conserva del `data.js` anterior cualquier mes que ya no esté en las fuentes. `dashboard.html` agrega un selector de mes y un `render(view)` que redibuja KPIs/gráficos/tablas.

**Tech Stack:** Python 3.14, openpyxl 3.1.5, pytest 9.1.1; HTML + Chart.js 4 (sin build).

**Spec:** `docs/superpowers/specs/2026-06-23-goa-historico-mensual-design.md`

**Directorio de trabajo:** todos los comandos se ejecutan desde el directorio del
proyecto `goa-inventory/` (la raíz de esta sesión). Los `git` funcionan igual desde
ahí porque encuentran la raíz del repo automáticamente.

---

## Estructura de archivos

| Archivo | Cambio | Responsabilidad |
|---|---|---|
| `fuentes/` | Crear (carpeta) | Contiene todos los Excel recibidos (local, gitignored). |
| `parse_excel.py` | Modificar | Ingesta multi-archivo, dedup, meses, carry-forward, agregación, `data.js`. |
| `tests/test_parse_excel.py` | Crear | Pruebas unitarias de la lógica nueva. |
| `tests/conftest.py` | Crear | Helper para construir `.xlsx` de prueba en disco temporal. |
| `data.js` | Regenerar | Salida; gana `months[]` y `meta.sources/carriedForward`. |
| `dashboard.html` | Modificar | Selector de mes + `render(view)` + ciclo de vida de Chart.js. |
| `README.md` | Modificar | Flujo nuevo (carpeta `fuentes/`). |

**Notas de diseño que se preservan del código actual:**
- `parse_transactions`, `aggregate_products/salespeople/customers/daily`, `parse_incentives` se **reutilizan**. Solo se añade el campo `month` a cada línea y se envuelven en funciones nuevas.
- Forma de `totals/products/salespeople/customers/daily/incentives` **idéntica** en la vista general y en cada mes, para que el dashboard use un solo `render`.

---

## Task 1: Carpeta de fuentes y andamiaje de pruebas

**Files:**
- Create: `fuentes/` (carpeta; mover los Excel actuales adentro)
- Create: `tests/conftest.py`
- Modify: `.gitignore` (ignorar `tests/__pycache__/` y `.pytest_cache/`)

- [ ] **Step 1: Crear la carpeta y mover los Excel existentes**

Run (Bash):
```bash
mkdir -p fuentes
mv "GOA clientes recurrentes.xlsx" "Latin-GOA-MAY.xlsx" fuentes/ 2>/dev/null || true
ls fuentes/
```
Expected: lista que incluye `GOA clientes recurrentes.xlsx` y `Latin-GOA-MAY.xlsx`. (Ignorar el archivo de bloqueo `~$...` si aparece.)

- [ ] **Step 2: Crear el helper de fixtures `tests/conftest.py`**

```python
import datetime
import sys
from pathlib import Path

# Hacer importable parse_excel.py (vive en el directorio padre de tests/)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
import pytest


def make_xlsx(path: Path, rows, sheet_name="GOA"):
    """Crea un .xlsx con `rows` (lista de listas) en la hoja indicada."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


HEADER = [
    "Order Date", "Order", "Product Variant", "Customer", "Salesperson",
    "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total",
]


def line(date, order, sku_name, customer, sp, qty=1, price=21.85, total=21.85):
    return [date, order, sku_name, customer, sp, "LatinFood US Corp",
            qty, qty, qty, price, total]


@pytest.fixture
def fixtures():
    return {"make_xlsx": make_xlsx, "HEADER": HEADER, "line": line, "dt": datetime.datetime}
```

- [ ] **Step 3: Ignorar artefactos de pruebas en git**

Añadir al final de `goa-inventory/.gitignore`:
```
.pytest_cache/
tests/__pycache__/
```

- [ ] **Step 4: Commit**

```bash
git add .gitignore tests/conftest.py
git commit -m "test: carpeta fuentes/ y andamiaje de pytest para GOA"
```
(La carpeta `fuentes/` y los `.xlsx` no se commitean: están en `.gitignore`.)

---

## Task 2: Descubrir fuentes y detectar la hoja

**Files:**
- Modify: `parse_excel.py` (config + funciones nuevas `list_source_files`, `read_sheet_rows`)
- Test: `tests/test_parse_excel.py`

- [ ] **Step 1: Escribir la prueba que falla**

Crear `tests/test_parse_excel.py`:
```python
import datetime
from pathlib import Path

import parse_excel as pe
from conftest import make_xlsx, HEADER, line


def test_list_source_files_ignora_temporales(tmp_path):
    make_xlsx(tmp_path / "a.xlsx", [HEADER])
    make_xlsx(tmp_path / "b.xlsx", [HEADER])
    (tmp_path / "~$a.xlsx").write_text("lock")  # archivo de bloqueo de Excel
    files = pe.list_source_files(tmp_path)
    names = sorted(f.name for f in files)
    assert names == ["a.xlsx", "b.xlsx"]


def test_read_sheet_rows_prefiere_GOA_luego_Sheet1(tmp_path):
    dt = datetime.datetime
    p = make_xlsx(tmp_path / "s.xlsx",
                  [HEADER, line(dt(2026, 2, 3), "S1", "[GOA01] Green Tea", "Tienda", "Ana")],
                  sheet_name="Sheet1")
    rows = pe.read_sheet_rows(p)
    assert rows[0][0] == "Order Date"
    assert rows[1][1] == "S1"
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: FAIL (`AttributeError: module 'parse_excel' has no attribute 'list_source_files'`).

- [ ] **Step 3: Implementar en `parse_excel.py`**

Reemplazar la sección de configuración:
```python
HERE = Path(__file__).parent
EXCEL_FILE = HERE / "Latin-GOA-MAY.xlsx"
OUTPUT_FILE = HERE / "data.js"
SHEET_NAME = "GOA"
```
por:
```python
HERE = Path(__file__).parent
SOURCES_DIR = HERE / "fuentes"
OUTPUT_FILE = HERE / "data.js"
SHEET_CANDIDATES = ("GOA", "Sheet1")  # orden de preferencia; si no, la primera hoja
```

Añadir estas funciones (después de los helpers de limpieza):
```python
def list_source_files(folder):
    """Todos los .xlsx de la carpeta, en orden de nombre, sin temporales (~$)."""
    folder = Path(folder)
    if not folder.exists():
        return []
    return sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))


def read_sheet_rows(path):
    """Filas (values_only) de la hoja preferida: GOA → Sheet1 → primera."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = None
    for name in SHEET_CANDIDATES:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        ws = wb.active
    return list(ws.iter_rows(values_only=True))
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS (2 passed).

- [ ] **Step 5: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py
git commit -m "feat: descubrir fuentes y detectar hoja (GOA/Sheet1/primera)"
```

---

## Task 3: Etiquetar cada línea con su mes

**Files:**
- Modify: `parse_excel.py` (`month_key`, `month_label`; añadir `month` en `parse_transactions`)
- Test: `tests/test_parse_excel.py`

- [ ] **Step 1: Escribir la prueba que falla**

Añadir a `tests/test_parse_excel.py`:
```python
def test_parse_transactions_agrega_mes():
    dt = datetime.datetime
    rows = [HEADER,
            line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana"),
            line(dt(2026, 5, 3), "S2", "[GOA02] Pure Chamoline", "Tienda B", "Luis")]
    lines = pe.parse_transactions(rows)
    assert lines[0]["month"] == "2026-02"
    assert lines[1]["month"] == "2026-05"


def test_month_label_en_espanol():
    assert pe.month_label("2026-02") == "Febrero 2026"
    assert pe.month_label("2026-05") == "Mayo 2026"
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -k "mes or label" -v`
Expected: FAIL (`KeyError: 'month'` y `AttributeError: month_label`).

- [ ] **Step 3: Implementar**

Añadir helpers:
```python
def month_key(date_iso):
    """'2026-02-24' -> '2026-02'."""
    return date_iso[:7]


def month_label(key):
    """'2026-02' -> 'Febrero 2026'."""
    y, m = key.split("-")
    return f"{MESES_ES[int(m)].capitalize()} {y}"
```

En `parse_transactions`, dentro del `lines.append({...})`, añadir el campo `month` (justo después de `"date"`):
```python
        lines.append({
            "date": date.date().isoformat(),
            "month": month_key(date.date().isoformat()),
            "order": clean_str(row[COL_ORDER]),
            # ... resto igual ...
        })
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py
git commit -m "feat: etiquetar cada línea de pedido con su mes (YYYY-MM)"
```

---

## Task 4: Cargar múltiples fuentes con dedup por número de orden

**Files:**
- Modify: `parse_excel.py` (`empty_incentives`, `merge_incentives`, `dominant_month`, `load_all_sources`)
- Test: `tests/test_parse_excel.py`

- [ ] **Step 1: Escribir la prueba que falla**

```python
def test_load_all_sources_dedup_por_orden(tmp_path):
    dt = datetime.datetime
    # archivo 1: orden S1 (feb)
    make_xlsx(tmp_path / "1.xlsx", [HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana")])
    # archivo 2: repite S1 y agrega S2 (mar)
    make_xlsx(tmp_path / "2.xlsx", [HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana"),
        line(dt(2026, 3, 1), "S2", "[GOA02] Pure Chamoline", "Tienda B", "Luis")])
    lines, inc_by_month, sources, dups = pe.load_all_sources(tmp_path)
    orders = sorted({ln["order"] for ln in lines})
    assert orders == ["S1", "S2"]          # S1 no se cuenta dos veces
    assert len(lines) == 2
    assert "S1" in dups                      # se reporta el duplicado
    assert sorted(sources) == ["1.xlsx", "2.xlsx"]
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -k dedup -v`
Expected: FAIL (`AttributeError: load_all_sources`).

- [ ] **Step 3: Implementar**

```python
def empty_incentives():
    return {"freeUnitsBySalesperson": [], "costItems": [], "totalDescuentos": 0.0}


def merge_incentives(incs):
    """Combina varios bloques de incentivos en uno."""
    free, cost, total = {}, [], 0.0
    for inc in incs:
        for f in inc.get("freeUnitsBySalesperson", []):
            free[f["name"]] = free.get(f["name"], 0) + f["free"]
        cost.extend(inc.get("costItems", []))
        total += inc.get("totalDescuentos", 0.0)
    return {
        "freeUnitsBySalesperson": [{"name": n, "free": v} for n, v in free.items()],
        "costItems": cost,
        "totalDescuentos": round(total, 2),
    }


def dominant_month(lines):
    """Mes más frecuente entre las líneas (para atribuir incentivos del archivo)."""
    counts = {}
    for ln in lines:
        counts[ln["month"]] = counts.get(ln["month"], 0) + 1
    return max(counts, key=counts.get)


def load_all_sources(folder):
    """Lee todas las fuentes, dedup por orden entre archivos, e incentivos por mes.

    Devuelve: (lines, incentives_by_month, source_names, dup_orders).
    La PRIMERA aparición de un número de orden gana; las demás se reportan.
    """
    seen_orders = set()
    lines = []
    incentives_by_month = {}
    dup_orders = []
    sources = []

    for path in list_source_files(folder):
        sources.append(path.name)
        try:
            rows = read_sheet_rows(path)
        except Exception as e:                       # archivo corrupto / sin hoja
            print(f"  ⚠ No se pudo leer {path.name}: {e} (se omite)")
            continue
        file_lines = parse_transactions(rows)
        kept = []
        for ln in file_lines:
            if ln["order"] in seen_orders:
                dup_orders.append(ln["order"])
                continue
            kept.append(ln)
        for ln in kept:
            seen_orders.add(ln["order"])
        lines.extend(kept)

        inc = parse_incentives(rows)
        has_inc = inc["freeUnitsBySalesperson"] or inc["costItems"] or inc["totalDescuentos"]
        if has_inc and kept:
            m = dominant_month(kept)
            incentives_by_month[m] = merge_incentives(
                [incentives_by_month.get(m, empty_incentives()), inc])

    return lines, incentives_by_month, sources, sorted(set(dup_orders))
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py
git commit -m "feat: cargar múltiples fuentes con dedup por orden e incentivos por mes"
```

---

## Task 5: Construir una vista agregada reutilizable

**Files:**
- Modify: `parse_excel.py` (`build_view` envolviendo las agregaciones + inyección de cajas gratis)
- Test: `tests/test_parse_excel.py`

- [ ] **Step 1: Escribir la prueba que falla**

```python
def test_build_view_totales_e_incentivos():
    dt = datetime.datetime
    lines = pe.parse_transactions([HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana", qty=2, total=43.70),
        line(dt(2026, 2, 25), "S2", "[GOA01] Green Tea", "Tienda B", "Ana", qty=1, total=21.85),
        line(dt(2026, 2, 25), "S3", "[GOA02] Pure Chamoline", "Tienda A", "Luis", qty=1, total=21.85)])
    inc = {"freeUnitsBySalesperson": [{"name": "Ana", "free": 5}],
           "costItems": [], "totalDescuentos": 70.95}
    v = pe.build_view(lines, inc)
    assert v["totals"]["units"] == 4
    assert v["totals"]["revenue"] == 87.40
    assert v["totals"]["orders"] == 3
    assert v["totals"]["customers"] == 2          # Tienda A y Tienda B
    ana = next(s for s in v["salespeople"] if s["name"] == "Ana")
    assert ana["freeUnits"] == 5                   # inyectado desde incentivos
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -k build_view -v`
Expected: FAIL (`AttributeError: build_view`).

- [ ] **Step 3: Implementar**

```python
def build_view(lines, incentives):
    """Vista agregada (misma forma para la general y para cada mes)."""
    products = aggregate_products(lines)
    salespeople = aggregate_salespeople(lines)
    customers = aggregate_customers(lines)
    daily = aggregate_daily(lines)

    free_map = {f["name"].lower(): f["free"] for f in incentives["freeUnitsBySalesperson"]}
    for s in salespeople:
        s["freeUnits"] = next(
            (v for k, v in free_map.items()
             if k in s["name"].lower() or s["name"].lower().startswith(k[:8])),
            0,
        )

    totals = {
        "units": sum(p["units"] for p in products),
        "revenue": round(sum(p["revenue"] for p in products), 2),
        "orders": len({ln["order"] for ln in lines}),
        "customers": len({ln["customer"] for ln in lines}),
        "salespeople": len(salespeople),
        "freeUnits": sum(f["free"] for f in incentives["freeUnitsBySalesperson"]),
    }
    return {
        "totals": totals, "products": products, "salespeople": salespeople,
        "customers": customers, "daily": daily, "incentives": incentives,
    }
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py
git commit -m "feat: build_view reutilizable (agregaciones + cajas gratis)"
```

---

## Task 6: Construir el desglose por mes

**Files:**
- Modify: `parse_excel.py` (`build_months`)
- Test: `tests/test_parse_excel.py`

- [ ] **Step 1: Escribir la prueba que falla**

```python
def test_build_months_separa_por_mes():
    dt = datetime.datetime
    lines = pe.parse_transactions([HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana"),
        line(dt(2026, 3, 1), "S2", "[GOA02] Pure Chamoline", "Tienda B", "Luis"),
        line(dt(2026, 3, 2), "S3", "[GOA01] Green Tea", "Tienda C", "Luis")])
    months = pe.build_months(lines, {})
    keys = [m["key"] for m in months]
    assert keys == ["2026-02", "2026-03"]                 # cronológico
    assert months[0]["label"] == "Febrero 2026"
    assert months[1]["totals"]["orders"] == 2             # S2 y S3
    assert months[0]["periodStart"] == "2026-02-24"
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -k build_months -v`
Expected: FAIL (`AttributeError: build_months`).

- [ ] **Step 3: Implementar**

```python
def build_months(lines, incentives_by_month):
    """Una vista por mes, en orden cronológico, con metadatos de periodo."""
    by_month = {}
    for ln in lines:
        by_month.setdefault(ln["month"], []).append(ln)

    out = []
    for key in sorted(by_month):
        month_lines = by_month[key]
        inc = incentives_by_month.get(key, empty_incentives())
        view = build_view(month_lines, inc)
        dates = sorted(ln["date"] for ln in month_lines)
        out.append({
            "key": key,
            "label": month_label(key),
            "periodStart": dates[0],
            "periodEnd": dates[-1],
            **view,
        })
    return out
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py
git commit -m "feat: desglose por mes (months[]) con metadatos de periodo"
```

---

## Task 7: Combinar vistas (merge_views) para el carry-forward

**Files:**
- Modify: `parse_excel.py` (`merge_views`)
- Test: `tests/test_parse_excel.py`

**Contexto:** `merge_views` combina varias vistas-mes en una sola. Solo se usa para la vista general cuando hay meses conservados del histórico (carry-forward). `units/revenue/orders` se suman (las órdenes no cruzan meses → exacto); `customers` y `salespeople` se cuentan por **unión de nombres** (exacto). El conteo de `customers` por vendedor se **suma** (puede sobrecontar a una tienda atendida por el mismo vendedor en dos meses — caveat documentado, solo afecta meses conservados).

- [ ] **Step 1: Escribir la prueba que falla**

```python
def test_merge_views_suma_y_une():
    dt = datetime.datetime
    feb = pe.build_view(pe.parse_transactions([HEADER,
        line(dt(2026, 2, 1), "S1", "[GOA01] Green Tea", "Tienda A", "Ana")]),
        pe.empty_incentives())
    mar = pe.build_view(pe.parse_transactions([HEADER,
        line(dt(2026, 3, 1), "S2", "[GOA01] Green Tea", "Tienda A", "Ana"),
        line(dt(2026, 3, 2), "S3", "[GOA02] Pure Chamoline", "Tienda B", "Luis")]),
        pe.empty_incentives())
    g = pe.merge_views([feb, mar])
    assert g["totals"]["units"] == 3
    assert g["totals"]["orders"] == 3                 # suma (órdenes únicas por mes)
    assert g["totals"]["customers"] == 2              # unión: Tienda A (feb+mar) + Tienda B
    green = next(p for p in g["products"] if p["sku"] == "GOA01")
    assert green["units"] == 2                         # GOA01 en feb y mar
    assert len(g["daily"]) == 3                        # concatenado
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -k merge_views -v`
Expected: FAIL (`AttributeError: merge_views`).

- [ ] **Step 3: Implementar**

```python
def merge_views(views):
    """Combina varias vistas (general + meses conservados) en una sola."""
    # Productos por SKU (con desglose por vendedor)
    prods = {}
    for v in views:
        for p in v["products"]:
            t = prods.setdefault(p["sku"], {
                "sku": p["sku"], "name": p["name"],
                "units": 0, "revenue": 0.0, "orders": 0, "_sp": {}})
            t["units"] += p["units"]; t["revenue"] += p["revenue"]; t["orders"] += p["orders"]
            for sp in p["salespeople"]:
                s = t["_sp"].setdefault(sp["name"], {"units": 0, "revenue": 0.0})
                s["units"] += sp["units"]; s["revenue"] += sp["revenue"]
    products = []
    for t in prods.values():
        sps = [{"name": n, "units": s["units"], "revenue": round(s["revenue"], 2)}
               for n, s in t["_sp"].items()]
        sps.sort(key=lambda s: (-s["units"], -s["revenue"]))
        products.append({"sku": t["sku"], "name": t["name"], "units": t["units"],
                         "revenue": round(t["revenue"], 2), "orders": t["orders"],
                         "salespeople": sps})
    products.sort(key=lambda p: (-p["units"], -p["revenue"]))

    # Vendedores por nombre
    sps = {}
    for v in views:
        for s in v["salespeople"]:
            t = sps.setdefault(s["name"], {"name": s["name"], "units": 0, "revenue": 0.0,
                                           "orders": 0, "customers": 0, "freeUnits": 0})
            t["units"] += s["units"]; t["revenue"] += s["revenue"]; t["orders"] += s["orders"]
            t["customers"] += s["customers"]; t["freeUnits"] += s.get("freeUnits", 0)
    salespeople = [{**t, "revenue": round(t["revenue"], 2)} for t in sps.values()]
    salespeople.sort(key=lambda s: (-s["revenue"], -s["units"]))

    # Clientes por nombre
    cs = {}
    for v in views:
        for c in v["customers"]:
            t = cs.setdefault(c["name"], {"name": c["name"], "units": 0, "revenue": 0.0, "orders": 0})
            t["units"] += c["units"]; t["revenue"] += c["revenue"]; t["orders"] += c["orders"]
    customers = [{**t, "revenue": round(t["revenue"], 2)} for t in cs.values()]
    customers.sort(key=lambda c: (-c["revenue"], -c["units"]))

    daily = sorted((d for v in views for d in v["daily"]), key=lambda d: d["date"])
    incentives = merge_incentives([v["incentives"] for v in views])

    totals = {
        "units": sum(p["units"] for p in products),
        "revenue": round(sum(p["revenue"] for p in products), 2),
        "orders": sum(v["totals"]["orders"] for v in views),
        "customers": len(cs),
        "salespeople": len(sps),
        "freeUnits": sum(f["free"] for f in incentives["freeUnitsBySalesperson"]),
    }
    return {"totals": totals, "products": products, "salespeople": salespeople,
            "customers": customers, "daily": daily, "incentives": incentives}
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py
git commit -m "feat: merge_views para combinar vistas-mes (carry-forward)"
```

---

## Task 8: Cargar el data.js anterior y conservar meses (carry-forward)

**Files:**
- Modify: `parse_excel.py` (`load_previous_report`, `carry_forward`)
- Test: `tests/test_parse_excel.py`

- [ ] **Step 1: Escribir la prueba que falla**

```python
def test_carry_forward_conserva_mes_ausente(tmp_path):
    prev = {"months": [
        {"key": "2026-01", "label": "Enero 2026", "periodStart": "2026-01-05",
         "periodEnd": "2026-01-20", "totals": {"units": 9}, "products": [],
         "salespeople": [], "customers": [], "daily": [], "incentives": pe.empty_incentives()},
        {"key": "2026-02", "label": "Febrero 2026", "totals": {"units": 1}, "products": [],
         "salespeople": [], "customers": [], "daily": [], "incentives": pe.empty_incentives()},
    ]}
    months = [{"key": "2026-02", "label": "Febrero 2026", "totals": {"units": 1},
               "products": [], "salespeople": [], "customers": [], "daily": [],
               "incentives": pe.empty_incentives()}]
    merged, carried = pe.carry_forward(months, prev)
    keys = [m["key"] for m in merged]
    assert keys == ["2026-01", "2026-02"]    # enero conservado, ordenado
    assert carried == ["2026-01"]


def test_load_previous_report_parsea_datajs(tmp_path):
    import json
    p = tmp_path / "data.js"
    p.write_text("const reportData = " + json.dumps({"months": []}) + ";\n", encoding="utf-8")
    rep = pe.load_previous_report(p)
    assert rep == {"months": []}
    assert pe.load_previous_report(tmp_path / "noexiste.js") is None
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -k "carry or previous" -v`
Expected: FAIL (`AttributeError`).

- [ ] **Step 3: Implementar**

```python
def load_previous_report(path):
    """Parsea el `reportData` del data.js anterior; None si no existe o no se puede."""
    path = Path(path)
    if not path.exists():
        return None
    txt = path.read_text(encoding="utf-8")
    m = re.search(r"const reportData = (.*);\s*$", txt, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except Exception:
        return None


def carry_forward(months, prev_report):
    """Conserva del data.js anterior los meses que ya no están en las fuentes.

    Devuelve (months_ordenados, carried_keys).
    """
    if not prev_report or "months" not in prev_report:
        return months, []
    present = {m["key"] for m in months}
    carried = []
    for pm in prev_report["months"]:
        if pm["key"] not in present:
            months.append(pm)
            carried.append(pm["key"])
    months.sort(key=lambda m: m["key"])
    return months, sorted(carried)
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py
git commit -m "feat: carry-forward de meses ausentes desde el data.js anterior"
```

---

## Task 9: Orquestación `build_report` + `main`

**Files:**
- Modify: `parse_excel.py` (reescribir `build_report` y `main`)
- Test: `tests/test_parse_excel.py` (integración con fuentes temporales)

- [ ] **Step 1: Escribir la prueba que falla**

```python
def test_build_report_integra_general_y_meses(tmp_path, monkeypatch):
    dt = datetime.datetime
    folder = tmp_path / "fuentes"
    folder.mkdir()
    make_xlsx(folder / "feb.xlsx", [HEADER,
        line(dt(2026, 2, 24), "S1", "[GOA01] Green Tea", "Tienda A", "Ana")])
    make_xlsx(folder / "mar.xlsx", [HEADER,
        line(dt(2026, 3, 1), "S2", "[GOA02] Pure Chamoline", "Tienda B", "Luis"),
        line(dt(2026, 3, 2), "S3", "[GOA01] Green Tea", "Tienda A", "Luis")])
    monkeypatch.setattr(pe, "SOURCES_DIR", folder)
    monkeypatch.setattr(pe, "OUTPUT_FILE", tmp_path / "data.js")

    report = pe.build_report()
    assert [m["key"] for m in report["months"]] == ["2026-02", "2026-03"]
    # La suma de los meses cuadra con el total general
    assert report["totals"]["units"] == sum(m["totals"]["units"] for m in report["months"])
    assert report["totals"]["orders"] == 3
    assert report["meta"]["sources"] == ["feb.xlsx", "mar.xlsx"]
    assert report["meta"]["carriedForward"] == []
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_parse_excel.py -k build_report -v`
Expected: FAIL (la `build_report` actual usa `EXCEL_FILE`/`SHEET_NAME`, ya inexistentes → `AttributeError`).

- [ ] **Step 3: Implementar**

Reemplazar la función `build_report` completa por:
```python
def build_report():
    lines, inc_by_month, sources, dup_orders = load_all_sources(SOURCES_DIR)
    if not lines:
        raise RuntimeError("No se detectaron líneas en fuentes/. ¿Pusiste los Excel?")

    months = build_months(lines, inc_by_month)
    prev = load_previous_report(OUTPUT_FILE)
    months, carried = carry_forward(months, prev)

    # Vista general: desde las líneas presentes; si hay meses conservados, se combinan.
    general = build_view(lines, merge_incentives(list(inc_by_month.values())))
    if carried:
        carried_views = [
            {k: m[k] for k in ("totals", "products", "salespeople",
                               "customers", "daily", "incentives")}
            for m in months if m["key"] in carried
        ]
        general = merge_views([general] + carried_views)

    period_start = min(m["periodStart"] for m in months)
    period_end = max(m["periodEnd"] for m in months)
    period_label = (months[0]["label"] if len(months) == 1
                    else f"{months[0]['label']} – {months[-1]['label']}")

    return {
        "meta": {
            "brand": BRAND,
            "distributor": DISTRIBUTOR,
            "periodStart": period_start,
            "periodEnd": period_end,
            "periodLabel": period_label,
            "lineItems": len(lines),
            "sources": sources,
            "carriedForward": carried,
            "duplicateOrders": dup_orders,
        },
        **general,
        "months": months,
    }
```

Reemplazar `main` por:
```python
def main():
    report = build_report()
    js = "const reportData = " + json.dumps(report, indent=2, ensure_ascii=False) + ";\n"
    OUTPUT_FILE.write_text(js, encoding="utf-8")

    m, t = report["meta"], report["totals"]
    print("✓ data.js generado")
    print(f"  Periodo:    {m['periodLabel']} ({m['periodStart']} → {m['periodEnd']})")
    print(f"  Fuentes:    {', '.join(m['sources']) or '(ninguna)'}")
    print(f"  Meses:      {', '.join(mo['key'] for mo in report['months'])}")
    print(f"  Líneas:     {m['lineItems']}")
    print(f"  Ingresos:   ${t['revenue']:,.2f}   Unidades: {t['units']}   Órdenes: {t['orders']}")
    if m["duplicateOrders"]:
        print(f"  ⚠ Órdenes duplicadas omitidas: {', '.join(m['duplicateOrders'])}")
    if m["carriedForward"]:
        print(f"  ⚠ Meses conservados del histórico (sin Excel en fuentes/): "
              f"{', '.join(m['carriedForward'])}")
```

- [ ] **Step 4: Correr y ver que pasa**

Run: `python -m pytest tests/test_parse_excel.py -v`
Expected: PASS (toda la suite).

- [ ] **Step 5: Generar el `data.js` real y revisar la salida**

Run: `python parse_excel.py`
Expected: imprime Periodo `Febrero 2026 – Mayo 2026`, fuentes con los dos archivos, meses `2026-02, 2026-03, 2026-04, 2026-05`, sin advertencias de duplicados/conservados.

- [ ] **Step 6: Verificar que la suma de meses cuadra (chequeo manual)**

Run:
```bash
python -c "
import json,re
d=json.loads(re.search(r'const reportData = (.*);',open('data.js',encoding='utf-8').read(),re.S).group(1))
su=sum(m['totals']['units'] for m in d['months']); sr=round(sum(m['totals']['revenue'] for m in d['months']),2)
so=sum(m['totals']['orders'] for m in d['months'])
print('general:',d['totals']['units'],d['totals']['revenue'],d['totals']['orders'])
print('suma meses:',su,sr,so)
assert d['totals']['units']==su and d['totals']['orders']==so, 'NO CUADRA'
print('OK: unidades y órdenes cuadran')
"
```
Expected: `OK: unidades y órdenes cuadran`. (Nota: la moneda puede diferir por centavos de redondeo; unidades y órdenes deben cuadrar exacto.)

- [ ] **Step 7: Commit**

```bash
git add parse_excel.py tests/test_parse_excel.py data.js
git commit -m "feat: build_report con vista general + meses + meta (histórico GOA)"
```

---

## Task 10: Selector de mes y `render(view)` en el dashboard

**Files:**
- Modify: `dashboard.html` (CSS del selector; markup del selector; reescritura del `<script>` en una función `render`)

**Contexto:** Hoy el `<script>` dibuja todo de forma imperativa leyendo `d` (la raíz). Se envuelve en `render(view, header)` y se añade un selector que alterna entre la vista general (`d`) y cada mes (`d.months[i]`). Los gráficos se guardan en `charts` y se destruyen antes de recrearlos.

- [ ] **Step 1: Añadir el CSS del selector**

En `dashboard.html`, dentro de `<style>`, después del bloque `.section-title { ... }` (≈ línea 76), añadir:
```css
  .month-selector {
    display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 22px;
  }
  .month-selector button {
    padding: 7px 16px; border-radius: 22px; border: 1px solid var(--border);
    background: var(--white); font-size: .82rem; font-weight: 600; cursor: pointer;
    color: var(--text-muted); transition: all .15s; box-shadow: var(--shadow);
  }
  .month-selector button:hover { border-color: var(--green-mid); color: var(--green-dark); }
  .month-selector button.active {
    background: var(--green-dark); color: #fff; border-color: var(--green-dark);
  }
```

- [ ] **Step 2: Añadir el markup del selector**

En `<main>`, justo antes de `<div class="kpi-grid" id="kpi-grid"></div>` (≈ línea 227), insertar:
```html
  <div class="month-selector" id="month-selector"></div>
```

- [ ] **Step 3: Reescribir el `<script>` final en una función `render`**

Reemplazar el **bloque `<script>` final completo** (el que empieza con `const d = reportData;` ≈ línea 341 y termina en su `</script>` ≈ línea 578; **no** tocar el `<script src="data.js">` de la línea 210) por este bloque (incluye sus propias etiquetas `<script>`/`</script>`):
```html
<script>
const d = reportData;

// ── Helpers ───────────────────────────────────────────────────────────
const usdFmt  = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
const usdFmt0 = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
const numFmt  = new Intl.NumberFormat('en-US');
const money  = n => usdFmt.format(n || 0);
const money0 = n => usdFmt0.format(n || 0);
const qty    = n => numFmt.format(n || 0);
const fmtDate = iso => new Date(iso + 'T00:00:00').toLocaleDateString('es-ES', { day: 'numeric', month: 'long', year: 'numeric' });
const fmtDay = iso => new Date(iso + 'T00:00:00').toLocaleDateString('es-ES', { day: '2-digit', month: 'short' });
const shortName = name => name.replace(' Herbal Tea', '').replace('Pure ', '');
const firstName = name => name.split(' ')[0];
const palette = ['#2d5a27','#3d7a32','#4a8c3f','#5ea050','#7ab648','#94c96a','#aad98a','#c0e8a8','#d4f0c2'];

// Plugin inline: dibuja el valor al final de cada barra
const barValueLabel = {
  id: 'barValueLabel',
  afterDatasetsDraw(chart, _args, opts) {
    const fmt = opts.formatter || (v => v);
    const { ctx, chartArea } = chart;
    ctx.save();
    ctx.font = '600 11px "Segoe UI", system-ui, sans-serif';
    ctx.fillStyle = '#1e2e1a';
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'left';
    chart.getDatasetMeta(0).data.forEach((bar, i) => {
      const text = fmt(chart.data.datasets[0].data[i]);
      let x = bar.x + 6;
      const w = ctx.measureText(text).width;
      if (x + w > chartArea.right) x = chartArea.right - w - 2;
      ctx.fillText(text, x, bar.y);
    });
    ctx.restore();
  }
};

let charts = {};

function horizontalBarChart(canvasId, items, fmt, tooltipLabel, extraTooltip) {
  const sorted = [...items].sort((a, b) => b.value - a.value);
  const labels = sorted.map(i => i.label);
  const data = sorted.map(i => i.value);
  const callbacks = { label: tooltipLabel };
  if (extraTooltip) callbacks.afterLabel = ctx => extraTooltip(sorted[ctx.dataIndex].meta);
  charts[canvasId] = new Chart(document.getElementById(canvasId), {
    type: 'bar',
    plugins: [barValueLabel],
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: labels.map((_, i) => palette[i % palette.length]),
        borderRadius: 5, borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks },
        barValueLabel: { formatter: fmt },
      },
      scales: {
        x: { grid: { color: '#e8f5e0' }, suggestedMax: Math.max(...data, 0) * 1.18, ticks: { font: { size: 11 }, callback: fmt } },
        y: { ticks: { font: { size: 11 } } }
      }
    }
  });
}

// ── Render de una vista (general o un mes) ────────────────────────────
function render(view, header) {
  const t = view.totals;

  // Encabezado
  document.getElementById('report-period').textContent = header.periodLabel;
  document.getElementById('report-range').textContent = `${fmtDay(header.periodStart)} – ${fmtDay(header.periodEnd)}`;

  // KPIs
  const topSp = view.salespeople[0] || { name: '—', revenue: 0 };
  const kpis = [
    { label: 'Ingresos totales', value: money0(t.revenue), sub: header.periodLabel, gold: true },
    { label: 'Unidades vendidas', value: qty(t.units), sub: `${qty(view.products.length)} SKUs` },
    { label: 'Clientes activos', value: qty(t.customers), sub: `${qty(t.orders)} órdenes` },
    { label: 'Vendedor top', value: topSp.name, sub: `${money0(topSp.revenue)} vendidos`, gold: true, small: true },
  ];
  document.getElementById('kpi-grid').innerHTML = kpis.map(k => `
    <div class="kpi-card ${k.gold ? 'gold' : ''}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value ${k.small ? 'sm' : ''}">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');

  // Destruir gráficos previos
  Object.values(charts).forEach(c => c.destroy());
  charts = {};

  // Ventas por SKU — Ingresos
  horizontalBarChart('revenueChart',
    view.products.map(p => ({ label: shortName(p.name), value: p.revenue })),
    money0, ctx => ` ${money(ctx.parsed.x)}`);

  // Ventas por SKU — Unidades
  horizontalBarChart('unitsChart',
    view.products.map(p => ({ label: shortName(p.name), value: p.units })),
    qty, ctx => ` ${qty(ctx.parsed.x)} unidades`);

  // Ranking de vendedores por ingresos
  horizontalBarChart('spChart',
    view.salespeople.map(s => ({ label: firstName(s.name), value: s.revenue, meta: s })),
    money0,
    ctx => ` ${money(ctx.parsed.x)} (${qty(Math.round(ctx.parsed.x / (t.revenue || 1) * 100))}%)`,
    s => `${qty(s.units)} u. · ${qty(s.orders)} órdenes · ${qty(s.customers)} clientes`);

  // Tendencia
  charts.trendChart = new Chart(document.getElementById('trendChart'), {
    type: 'line',
    data: {
      labels: view.daily.map(x => fmtDay(x.date)),
      datasets: [{
        label: 'Ingresos', data: view.daily.map(x => x.revenue),
        borderColor: '#2d5a27', backgroundColor: 'rgba(122,182,72,.18)',
        fill: true, tension: .3, pointRadius: 4, pointBackgroundColor: '#c8a84b', borderWidth: 2,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: {
          title: items => fmtDate(view.daily[items[0].dataIndex].date),
          label: ctx => ` ${money(ctx.parsed.y)} · ${qty(view.daily[ctx.dataIndex].units)} u. · ${qty(view.daily[ctx.dataIndex].orders)} órdenes`
        } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 } } },
        y: { grid: { color: '#e8f5e0' }, ticks: { font: { size: 10 }, callback: v => money0(v) } }
      }
    }
  });

  // Tabla de productos (expandible por vendedor)
  const pbody = document.getElementById('product-tbody');
  pbody.innerHTML = '';
  view.products.forEach((p, idx) => {
    const tr = document.createElement('tr');
    tr.className = 'product-row';
    tr.setAttribute('onclick', `toggleSp(${idx}, this)`);
    tr.innerHTML = `
      <td><span class="sku-badge">${p.sku}</span>${p.name}</td>
      <td class="num">${qty(p.units)}</td>
      <td class="num"><strong>${money(p.revenue)}</strong></td>
      <td class="num">${qty(p.orders)}</td>
      <td><span class="toggle-btn"><span class="arrow">▶</span> ${p.salespeople.length} vend.</span></td>`;
    pbody.appendChild(tr);
    p.salespeople.forEach(sp => {
      const spTr = document.createElement('tr');
      spTr.className = `sp-row sp-group-${idx}`;
      spTr.innerHTML = `
        <td>${sp.name}</td>
        <td class="num">${qty(sp.units)}</td>
        <td class="num">${money(sp.revenue)}</td>
        <td class="num"></td>
        <td></td>`;
      pbody.appendChild(spTr);
    });
  });
  const ptotal = document.createElement('tr');
  ptotal.className = 'total-row';
  ptotal.innerHTML = `<td>TOTAL</td><td class="num">${qty(t.units)}</td><td class="num">${money(t.revenue)}</td><td class="num">${qty(t.orders)}</td><td></td>`;
  pbody.appendChild(ptotal);

  // Tabla de vendedores
  const sbody = document.getElementById('sp-tbody');
  sbody.innerHTML = '';
  view.salespeople.forEach((sp, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="sp-rank">${i + 1}</td>
      <td>${sp.name}</td>
      <td class="num">${qty(sp.units)}</td>
      <td class="num"><strong>${money(sp.revenue)}</strong></td>
      <td class="num">${qty(sp.orders)}</td>
      <td class="num">${qty(sp.customers)}</td>
      <td class="num">${sp.freeUnits ? `<span class="free-badge">${qty(sp.freeUnits)}</span>` : '—'}</td>`;
    sbody.appendChild(tr);
  });

  // Tabla de clientes
  const cbody = document.getElementById('cust-tbody');
  cbody.innerHTML = '';
  view.customers.forEach((c, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="sp-rank">${i + 1}</td>
      <td>${c.name}</td>
      <td class="num">${qty(c.orders)}</td>
      <td class="num">${qty(c.units)}</td>
      <td class="num"><strong>${money(c.revenue)}</strong></td>`;
    cbody.appendChild(tr);
  });

  // Incentivos
  const fbody = document.getElementById('free-tbody');
  const freeRows = view.incentives.freeUnitsBySalesperson;
  if (freeRows.length) {
    fbody.innerHTML = freeRows.map(f =>
      `<tr><td>${f.name}</td><td class="num"><span class="free-badge">${qty(f.free)}</span></td></tr>`).join('');
  } else {
    fbody.innerHTML = `<tr><td colspan="2" style="color:var(--text-muted)">Sin datos</td></tr>`;
  }
  const clist = document.getElementById('cost-list');
  clist.innerHTML = view.incentives.costItems.map(ci =>
    `<li><span>${ci.label}</span><span class="amount">${money(ci.amount)}</span></li>`
  ).join('') || `<li style="color:var(--text-muted)">Sin ítems de costo</li>`;
  document.getElementById('cost-total').textContent = money(view.incentives.totalDescuentos);
}

function toggleSp(idx, rowEl) {
  const rows = document.querySelectorAll(`.sp-group-${idx}`);
  const arrow = rowEl.querySelector('.arrow');
  const open = arrow.classList.toggle('open');
  rows.forEach(r => r.classList.toggle('visible', open));
}

// ── Selector de mes ───────────────────────────────────────────────────
const VIEWS = [
  { id: 'all', label: 'Todos', view: d,
    header: { periodLabel: d.meta.periodLabel, periodStart: d.meta.periodStart, periodEnd: d.meta.periodEnd } },
  ...d.months.map(m => ({
    id: m.key, label: m.label.split(' ')[0], view: m,
    header: { periodLabel: m.label, periodStart: m.periodStart, periodEnd: m.periodEnd } }))
];

const selector = document.getElementById('month-selector');
selector.innerHTML = VIEWS.map((v, i) =>
  `<button data-i="${i}" class="${i === 0 ? 'active' : ''}">${v.label}</button>`).join('');
selector.addEventListener('click', e => {
  const btn = e.target.closest('button');
  if (!btn) return;
  selector.querySelectorAll('button').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  const v = VIEWS[+btn.dataset.i];
  render(v.view, v.header);
});

// Render inicial: vista general
render(VIEWS[0].view, VIEWS[0].header);
</script>
```

- [ ] **Step 4: Verificación en navegador**

Run (Bash, sirve y captura):
```bash
python -m http.server 8765 >/dev/null 2>&1 &
sleep 1
```
Abrir `http://localhost:8765/dashboard.html`. Confirmar visualmente:
- Aparece la barra de botones `Todos · Febrero · Marzo · Abril · Mayo`.
- `Todos` muestra el periodo `Febrero 2026 – Mayo 2026` y los totales combinados.
- Al hacer clic en cada mes: KPIs, los 4 gráficos y las 3 tablas se redibujan; el encabezado cambia al mes; los incentivos solo aparecen en `Mayo`.
- Alternar varias veces no rompe los gráficos (sin parpadeo infinito).

Detener el servidor: `kill %1` (o cerrar la terminal de fondo).

- [ ] **Step 5: Commit**

```bash
git add dashboard.html
git commit -m "feat: selector de mes y render(view) en el dashboard GOA"
```

---

## Task 11: Actualizar el README

**Files:**
- Modify: `README.md` (secciones "Formato del Excel", "Actualizar datos", "Archivos")

- [ ] **Step 1: Reemplazar la sección "Actualizar datos"**

Sustituir el bloque actual (pasos 1–4 bajo `## Actualizar datos`) por:
```markdown
## Actualizar datos

1. Deja el/los Excel que recibas en la carpeta `fuentes/` (no se suben al repo;
   los `*.xlsx` están en `.gitignore` por privacidad). Un mismo archivo puede traer
   varios meses; se separan solos.
2. Ejecuta: `python parse_excel.py`
   - Lee **todos** los Excel de `fuentes/`, descarta órdenes duplicadas y reconstruye
     el histórico completo (todos los meses) en `data.js`.
   - Si un mes que ya estaba publicado ya no tiene Excel en la carpeta, se conserva
     del `data.js` anterior (sale un aviso).
3. Sube `data.js` al repo:
   ```
   git add data.js
   git commit -m "data: actualizar reporte GOA"
   git push
   ```
4. El dashboard se actualiza automáticamente en GitHub Pages, con el mes nuevo
   disponible en el selector de arriba.
```

- [ ] **Step 2: Actualizar la sección "Formato del Excel" y la tabla "Archivos"**

En "Formato del Excel (definitivo)", añadir al final un párrafo:
```markdown
El distribuidor puede entregar un mes suelto (hoja `GOA`) o un export con varios
meses (hoja `Sheet1`); el script detecta la hoja y separa por mes según la fecha de
cada línea. El bloque de incentivos es opcional.
```

En la tabla "Archivos", añadir una fila:
```markdown
| `fuentes/` | Carpeta local con todos los Excel recibidos (no se sube al repo) |
```

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: README con el flujo de carpeta fuentes/ e histórico mensual"
```

---

## Task 12: Verificación final

**Files:** (ninguno — solo verificación)

- [ ] **Step 1: Suite de pruebas completa en verde**

Run: `python -m pytest tests/ -v`
Expected: todos los tests PASS.

- [ ] **Step 2: Regenerar data.js desde cero y revisar advertencias**

Run: `python parse_excel.py`
Expected: Periodo `Febrero 2026 – Mayo 2026`, 4 meses, sin advertencias de duplicados ni conservados (con ambos Excel en `fuentes/`).

- [ ] **Step 3: Smoke test de carry-forward (manual, reversible)**

```bash
mv "fuentes/Latin-GOA-MAY.xlsx" /tmp/   # quitar mayo temporalmente
python parse_excel.py                    # debe AVISAR que conserva 2026-05
python -c "import json,re; d=json.loads(re.search(r'const reportData = (.*);',open('data.js',encoding='utf-8').read(),re.S).group(1)); print('meses:', [m['key'] for m in d['months']]); assert '2026-05' in [m['key'] for m in d['months']], 'mayo no se conservó'; print('OK carry-forward')"
mv "/tmp/Latin-GOA-MAY.xlsx" fuentes/    # restaurar
python parse_excel.py                    # vuelve a 4 meses sin avisos
```
Expected: el paso del medio imprime el aviso de mes conservado y `OK carry-forward`; tras restaurar, vuelve a generarse limpio.

- [ ] **Step 4: Commit del data.js final**

```bash
git add data.js
git commit -m "data: histórico GOA feb–may 2026 (general + desglose por mes)"
```

- [ ] **Step 5: Cierre**

Usar la skill `superpowers:finishing-a-development-branch` para decidir cómo integrar la rama `feat/goa-historico-mensual` (merge a `main`, PR, o limpieza) y si se hace push.

---

## Notas de verificación cruzada (spec → plan)

- Ingesta formato nuevo (Sheet1, multi-mes, sin incentivos) → Tasks 2, 3, 4.
- Acumulación en carpeta `fuentes/` con dedup → Tasks 1, 4.
- Carry-forward a prueba de pérdidas → Tasks 7, 8, 9 (+ smoke test Task 12).
- Total general + `months[]` → Tasks 6, 9.
- Selector de mes + `render(view)` → Task 10.
- README → Task 11.
- Privacidad (Excel gitignored, solo `data.js` público) → preservado (Task 1 no commitea `fuentes/`).
```
