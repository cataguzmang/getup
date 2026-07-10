# Parsing header-aware — Plan de Implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que `report_core.py` y `split_excel.py` ubiquen las columnas por nombre de encabezado (tolerando columnas insertadas/quitadas/reordenadas), con fallback posicional y warnings por faltantes.

**Architecture:** Módulo nuevo `column_resolver.py` en la raíz (find_header_row / resolve_columns / get / CANONICAL_POSITIONS). `report_core.load_all_sources` resuelve columnas **por archivo**; `split_excel.collect` resuelve **por hoja** y rearma cada fila al orden canónico antes de todo lo demás, de modo que el resto del código (y el canónico escrito en `fuentes/`) no cambia.

**Tech Stack:** Python 3.14 + openpyxl + pytest. Sin cambios de formato de salida ni de `reportData`.

**Regla de seguridad:** con encabezado presente se confía solo en los nombres; sin encabezado, posiciones canónicas (comportamiento actual). Los archivos canónicos existentes mapean 1:1 → regresión idéntica (72 tests actuales deben pasar sin tocar).

---

## File Structure

- **Crear** `column_resolver.py` (raíz) — resolvedor compartido.
- **Crear** `tests/test_column_resolver.py`.
- **Modificar** `report_core.py` — `is_transaction_row`, `parse_transactions`, `load_all_sources`.
- **Modificar** `split_excel.py` — import, helper `canonical_row`, cuerpo de `collect`.
- **Ampliar** `tests/test_report_core.py` y crear `tests/test_header_aware.py` (split_excel).
- **Regenerar** `kombuchacha/data.js` con la fuente real nueva (validación end-to-end).

---

## Task 1: `column_resolver.py`

**Files:**
- Create: `column_resolver.py`
- Test: `tests/test_column_resolver.py`

- [ ] **Step 1: Escribir los tests que fallan**

Create `tests/test_column_resolver.py`:

```python
import column_resolver as cr

HEADER = ["Order Date", "Order", "Product Variant", "Customer", "Salesperson",
          "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total"]
# Layout real recibido el 2026-07-10: Address insertada, sin Company
HEADER_ADDR = ["Order Date", "Order", "Product Variant", "Customer", "Address",
               "Salesperson", "Qty Delivered", "Qty Invoiced", "Qty Ordered",
               "Unit Price", "Total"]


def test_find_header_row_primera_fila():
    assert cr.find_header_row([HEADER, ["x"] * 11]) == 0


def test_find_header_row_en_medio():
    rows = [["basura"], [None, None], HEADER, ["x"] * 11]
    assert cr.find_header_row(rows) == 2


def test_find_header_row_ausente():
    assert cr.find_header_row([["a", "b", "c"], [1, 2, 3]]) is None


def test_find_header_row_respeta_max_scan():
    rows = [["x"]] * 12 + [HEADER]
    assert cr.find_header_row(rows, max_scan=10) is None


def test_resolve_canonico_mapea_identico():
    assert cr.resolve_columns(HEADER) == cr.CANONICAL_POSITIONS


def test_resolve_con_address_insertada_y_sin_company(capsys):
    cols = cr.resolve_columns(HEADER_ADDR, label="Sales Kombuchacha.xlsx")
    assert cols["customer"] == 3
    assert cols["salesperson"] == 5      # corrida por Address
    assert cols["qty_delivered"] == 6
    assert "company" not in cols          # ausente -> no aparece
    out = capsys.readouterr().out
    assert "company" in out.lower() and "Sales Kombuchacha.xlsx" in out


def test_resolve_normaliza_mayusculas_y_espacios():
    header = ["  order DATE ", "ORDER", "product   variant", "Customer",
              "SalesPerson", "company", "qty delivered", "qty invoiced",
              "qty ordered", "unit price", "total"]
    cols = cr.resolve_columns(header)
    assert cols == cr.CANONICAL_POSITIONS


def test_get_tolerante():
    row = ["2026-06-01", "S1", "[KOM01] x"]
    cols = {"date": 0, "variant": 2, "total": 10}
    assert cr.get(row, cols, "date") == "2026-06-01"
    assert cr.get(row, cols, "total") is None            # índice fuera de rango
    assert cr.get(row, cols, "company", "") == ""        # campo no resuelto
```

- [ ] **Step 2: Verificar que fallan**

Run: `python -m pytest tests/test_column_resolver.py -v`
Expected: FAIL con `ModuleNotFoundError: No module named 'column_resolver'`.

- [ ] **Step 3: Implementar `column_resolver.py`**

Create `column_resolver.py`:

```python
"""
column_resolver.py — Resolución de columnas por nombre de encabezado.

El distribuidor cambia el layout de sus Excel sin aviso (columnas insertadas,
quitadas o reordenadas). Este módulo ubica cada campo lógico por el NOMBRE de su
encabezado en vez de por posición fija, con fallback posicional cuando no hay
encabezado reconocible. Lo comparten `report_core.py` y `split_excel.py`.

Caso real (2026-07-10): 'Sales Kombuchacha.xlsx' llegó con una columna Address
insertada y sin Company; el mapeo posicional leyó direcciones como vendedores.
"""

import re

# Campo lógico -> encabezado esperado (normalizado: minúsculas, espacios simples)
FIELD_HEADERS = {
    "date": "order date",
    "order": "order",
    "variant": "product variant",
    "customer": "customer",
    "salesperson": "salesperson",
    "company": "company",
    "qty_delivered": "qty delivered",
    "qty_invoiced": "qty invoiced",
    "qty_ordered": "qty ordered",
    "unit_price": "unit price",
    "total": "total",
}

# Posiciones del formato canónico (fallback cuando no hay encabezado)
CANONICAL_POSITIONS = {
    "date": 0, "order": 1, "variant": 2, "customer": 3, "salesperson": 4,
    "company": 5, "qty_delivered": 6, "qty_invoiced": 7, "qty_ordered": 8,
    "unit_price": 9, "total": 10,
}

# Columnas imprescindibles para reconocer una fila como encabezado
_REQUIRED = ("order date", "product variant")


def _norm(cell):
    """Normaliza una celda para comparar: str, trim, espacios simples, minúsculas."""
    if cell is None:
        return ""
    return re.sub(r"\s+", " ", str(cell)).strip().lower()


def find_header_row(rows, max_scan=10):
    """Índice de la primera fila (entre las primeras `max_scan`) que contiene los
    encabezados imprescindibles. None si no aparece."""
    for i, row in enumerate(rows[:max_scan]):
        normed = {_norm(c) for c in row}
        if all(r in normed for r in _REQUIRED):
            return i
    return None


def resolve_columns(header_row, label=""):
    """{campo_lógico: índice} para cada encabezado reconocido en la fila.

    Los campos no encontrados NO aparecen en el dict; por cada faltante se
    imprime un aviso (una vez por llamada). `label` identifica archivo/hoja."""
    normed = [_norm(c) for c in header_row]
    cols = {}
    for field, header in FIELD_HEADERS.items():
        if header in normed:
            cols[field] = normed.index(header)
        else:
            where = f"{label}: " if label else ""
            print(f"  ⚠ {where}columna '{header}' no encontrada en el encabezado; "
                  f"sus valores quedarán vacíos")
    return cols


def get(row, cols, field, default=None):
    """row[cols[field]] con tolerancia: campo sin resolver o fila corta -> default."""
    i = cols.get(field)
    if i is None or i >= len(row):
        return default
    return row[i]
```

- [ ] **Step 4: Verificar que pasan**

Run: `python -m pytest tests/test_column_resolver.py -v`
Expected: PASS (9 tests).

- [ ] **Step 5: Commit**

```bash
git add column_resolver.py tests/test_column_resolver.py
git commit -m "feat: column_resolver — resolución de columnas por encabezado"
```

---

## Task 2: `report_core.py` header-aware

**Files:**
- Modify: `report_core.py`
- Test: `tests/test_report_core.py` (ampliar)

- [ ] **Step 1: Escribir los tests que fallan**

Append a `tests/test_report_core.py`:

```python
HEADER_ADDR = ["Order Date", "Order", "Product Variant", "Customer", "Address",
               "Salesperson", "Qty Delivered", "Qty Invoiced", "Qty Ordered",
               "Unit Price", "Total"]


def _line_addr(date, order, variant, customer, address, sp,
               qty=1, price=21.85, total=21.85):
    """Fila con el layout nuevo: Address insertada, sin Company."""
    return [date, order, variant, customer, address, sp, qty, qty, qty, price, total]


def test_layout_con_address_lee_vendedor_correcto(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "Sales Kombuchacha.xlsx", [HEADER_ADDR,
        _line_addr(dt(2026, 5, 12), "S42061", "[KOM01] Kombuchacha Zero Blueberry",
                   "Tropical Sprmkt - Dunellen", "446 North Ave, Dunellen, NJ 08812",
                   "Mireya Fernandez", qty=0.5, price=50.64, total=25.32),
    ])
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    names = [s["name"] for s in rep["salespeople"]]
    assert names == ["Mireya Fernandez"]        # persona, NO la dirección
    assert rep["totals"]["revenue"] == 25.32


def test_mezcla_canonico_y_layout_nuevo(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    # archivo canónico (11 columnas estándar)
    _xlsx(fuentes / "KOM-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
              "Convenience store flora llc", "Mireya Fernandez",
              qty=0.5, price=50.64, total=25.32),
    ], sheet="KOM")
    # archivo con layout nuevo (Address, sin Company)
    _xlsx(fuentes / "Sales Kombuchacha.xlsx", [HEADER_ADDR,
        _line_addr(dt(2026, 5, 4), "S41495", "[KOM01] Kombuchacha Zero Blueberry",
                   "Castellanos Grocery", "2 Orchard St, Morristown, NJ 07860",
                   "Katherine Osorio Duque", qty=1, price=54.64, total=54.64),
    ])
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    names = sorted(s["name"] for s in rep["salespeople"])
    assert names == ["Katherine Osorio Duque", "Mireya Fernandez"]
    assert rep["totals"]["revenue"] == 79.96      # 25.32 + 54.64
    assert len(rep["months"]) == 2                 # 2026-05 y 2026-06


def test_sin_encabezado_usa_posiciones_canonicas(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    # sin fila de encabezado: debe caer al mapeo posicional canónico
    _xlsx(fuentes / "raw.xlsx", [
        _line(dt(2026, 6, 1), "S1", "[KOM01] Kombuchacha Zero Blueberry",
              "Tienda", "Mireya Fernandez", qty=1, price=50.64, total=50.64),
    ])
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp",
                          "KOM", ("KOM", "Sheet1"))
    assert rep["salespeople"][0]["name"] == "Mireya Fernandez"
    assert rep["totals"]["revenue"] == 50.64
```

- [ ] **Step 2: Verificar que fallan**

Run: `python -m pytest tests/test_report_core.py -v`
Expected: los 3 tests nuevos FAIL (`test_layout_con_address...` lee la dirección como vendedor); los 5 previos PASS.

- [ ] **Step 3: Modificar `report_core.py`**

(a) En los imports (después de `import openpyxl`), agregar:

```python
from column_resolver import (CANONICAL_POSITIONS, find_header_row, get,
                             resolve_columns)
```

(b) Reemplazar `is_transaction_row`:

```python
def is_transaction_row(row, pattern):
    return (
        isinstance(row[COL_DATE], datetime.datetime)
        and isinstance(row[COL_VARIANT], str)
        and pattern.match(row[COL_VARIANT])
    )
```

por:

```python
def is_transaction_row(row, pattern, cols):
    date = get(row, cols, "date")
    variant = get(row, cols, "variant")
    return (
        isinstance(date, datetime.datetime)
        and isinstance(variant, str)
        and pattern.match(variant)
    )
```

(c) Reemplazar `parse_transactions` completo:

```python
def parse_transactions(rows, pattern):
    """Devuelve la lista de líneas de pedido limpias (dicts)."""
    lines = []
    for row in rows:
        if not is_transaction_row(row, pattern):
            continue
        sku, name = split_sku(row[COL_VARIANT], pattern)
        if sku is None:
            continue
        date = row[COL_DATE]
        lines.append({
            "date": date.date().isoformat(),
            "month": month_key(date.date().isoformat()),
            "order": clean_str(row[COL_ORDER]),
            "sku": sku,
            "product": name,
            "customer": clean_str(row[COL_CUSTOMER]),
            "salesperson": clean_str(row[COL_SALESPERSON]),
            "qtyOrdered": q(row[COL_QTY_ORDERED]),
            "qtyDelivered": q(row[COL_QTY_DELIVERED]),
            "qtyInvoiced": q(row[COL_QTY_INVOICED]),
            "unitPrice": round(num(row[COL_UNIT_PRICE]), 2),
            "total": round(num(row[COL_TOTAL]), 2),
        })
    return lines
```

por:

```python
def parse_transactions(rows, pattern, cols=None):
    """Devuelve la lista de líneas de pedido limpias (dicts).

    `cols` mapea campo lógico -> índice (ver column_resolver). Si es None, se
    resuelve desde el encabezado de `rows`; sin encabezado, posiciones canónicas.
    """
    if cols is None:
        hdr_i = find_header_row(rows)
        cols = (resolve_columns(rows[hdr_i]) if hdr_i is not None
                else dict(CANONICAL_POSITIONS))
    lines = []
    for row in rows:
        if not is_transaction_row(row, pattern, cols):
            continue
        sku, name = split_sku(get(row, cols, "variant"), pattern)
        if sku is None:
            continue
        date = get(row, cols, "date")
        lines.append({
            "date": date.date().isoformat(),
            "month": month_key(date.date().isoformat()),
            "order": clean_str(get(row, cols, "order")),
            "sku": sku,
            "product": name,
            "customer": clean_str(get(row, cols, "customer")),
            "salesperson": clean_str(get(row, cols, "salesperson")),
            "qtyOrdered": q(get(row, cols, "qty_ordered")),
            "qtyDelivered": q(get(row, cols, "qty_delivered")),
            "qtyInvoiced": q(get(row, cols, "qty_invoiced")),
            "unitPrice": round(num(get(row, cols, "unit_price")), 2),
            "total": round(num(get(row, cols, "total")), 2),
        })
    return lines
```

(d) En `load_all_sources`, reemplazar:

```python
        sources.append(path.name)
        file_lines = parse_transactions(rows, pattern)
```

por:

```python
        sources.append(path.name)
        hdr_i = find_header_row(rows)
        cols = (resolve_columns(rows[hdr_i], label=path.name)
                if hdr_i is not None else dict(CANONICAL_POSITIONS))
        file_lines = parse_transactions(rows, pattern, cols)
```

Nota: los índices `COL_*` de `report_core.py` se conservan — `parse_incentives`
los sigue usando (el bloque de incentivos no es tabular y no cambia).

- [ ] **Step 4: Verificar que pasan**

Run: `python -m pytest tests/test_report_core.py -v`
Expected: PASS (8 tests: 5 previos + 3 nuevos).

- [ ] **Step 5: Regresión de la raíz**

Run: `python -m pytest tests/ -q`
Expected: todo PASS.

- [ ] **Step 6: Commit**

```bash
git add report_core.py tests/test_report_core.py
git commit -m "feat: report_core lee columnas por encabezado (por archivo)"
```

---

## Task 3: `split_excel.py` header-aware

**Files:**
- Modify: `split_excel.py`
- Test: `tests/test_header_aware.py` (nuevo)

- [ ] **Step 1: Escribir los tests que fallan**

Create `tests/test_header_aware.py`:

```python
import openpyxl

import split_excel as sx
from helpers import D, make_workbook

HEADER_ADDR = ["Order Date", "Order", "Product Variant", "Customer", "Address",
               "Salesperson", "Qty Delivered", "Qty Invoiced", "Qty Ordered",
               "Unit Price", "Total"]


def _wb_addr(tmp_path):
    """Workbook de entrada con el layout nuevo (Address insertada, sin Company)."""
    sheets = {
        "NY Jun - KOM": [
            HEADER_ADDR,
            [D(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
             "Convenience store flora llc", "132-134 Bloomfield Ave, Newark, NJ",
             "Mireya Fernandez", 0.5, 0.5, 0.5, 50.64, 25.32],
        ],
    }
    return make_workbook(tmp_path / "GET Jun 2026.xlsx", sheets)


def test_collect_reordena_a_canonico(tmp_path):
    data = sx.collect([_wb_addr(tmp_path)])
    row = data.rows["KOM"]["2026-06"][0]
    assert row[sx.COL_SP] == "Mireya Fernandez"     # persona en la col canónica 4
    assert row[sx.COL_COMPANY] is None              # Company ausente -> vacía
    assert row[sx.COL_CUSTOMER] == "Convenience store flora llc"
    assert len(row) == len(sx.HEADER)               # siempre 11 columnas


def test_estado_cae_al_prefijo_de_hoja_sin_company(tmp_path):
    data = sx.collect([_wb_addr(tmp_path)])
    # sin Company, el estado sale del prefijo de la hoja ('NY ...')
    assert data.stats["KOM"]["2026-06"]["Nueva York"] == 1


def test_write_canonical_desde_layout_nuevo(tmp_path):
    data = sx.collect([_wb_addr(tmp_path)])
    out = tmp_path / "KOM-2026-06.xlsx"
    sx.write_canonical(out, data.rows["KOM"]["2026-06"], [], "KOM")
    ws = openpyxl.load_workbook(out)["KOM"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows[0] == tuple(sx.HEADER)
    assert rows[1][4] == "Mireya Fernandez"         # Salesperson en su lugar
    assert rows[1][6] == 0.5                        # Qty Delivered en su lugar


def test_canonico_sin_cambios_sigue_identico(messy_wb):
    """El workbook canónico de siempre produce exactamente lo mismo (regresión)."""
    data = sx.collect([messy_wb])
    row = data.rows["KOM"]["2026-06"][0]
    assert row[sx.COL_SP] == "Mireya Fernandez"
    assert row[sx.COL_COMPANY] == "LatinFood US Corp."
```

- [ ] **Step 2: Verificar que fallan**

Run: `python -m pytest tests/test_header_aware.py -v`
Expected: `test_collect_reordena_a_canonico`, `test_estado_cae_al_prefijo_de_hoja_sin_company` y `test_write_canonical_desde_layout_nuevo` FAIL (hoy la dirección queda en la col 4); `test_canonico_sin_cambios_sigue_identico` PASS.

- [ ] **Step 3: Modificar `split_excel.py`**

(a) En los imports (después de `import openpyxl`), agregar:

```python
from column_resolver import (CANONICAL_POSITIONS, find_header_row, get,
                             resolve_columns)
```

(b) Después de la definición de `HEADER` (tras la línea de `COL_DATE, ... = 0, 1, 2, 3, 4, 5`), agregar:

```python
# Orden de campos lógicos del formato canónico (para rearmar filas)
CANONICAL_FIELDS = ("date", "order", "variant", "customer", "salesperson",
                    "company", "qty_delivered", "qty_invoiced", "qty_ordered",
                    "unit_price", "total")


def canonical_row(row, cols):
    """Rearma la fila en el orden canónico de HEADER según el mapa de columnas."""
    return [get(row, cols, f) for f in CANONICAL_FIELDS]
```

(c) En `collect`, reemplazar:

```python
        for ws in wb.worksheets:
            sheet_rows = list(ws.iter_rows(values_only=True))
            tx_codes, tx_months, tx_states = [], [], []

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
                tx_states.append(state_of(row, ws.title))
```

por:

```python
        for ws in wb.worksheets:
            sheet_rows = list(ws.iter_rows(values_only=True))
            hdr_i = find_header_row(sheet_rows)
            cols = (resolve_columns(sheet_rows[hdr_i],
                                    label=f"{Path(path).name}/{ws.title}")
                    if hdr_i is not None else dict(CANONICAL_POSITIONS))
            tx_codes, tx_months, tx_states = [], [], []

            for row in sheet_rows:
                # Rearmar al orden canónico ANTES de procesar: el resto del
                # pipeline (y el .xlsx que se escribe) siempre ve el layout canónico.
                crow = canonical_row(row, cols)
                if not is_transaction_row(crow):
                    continue
                prefix = sku_prefix(crow[COL_VARIANT])
                mk = month_key(crow[COL_DATE])
                if prefix not in BRANDS:
                    unmatched.append(crow)
                    continue
                rows[prefix][mk].append(crow)
                stats[prefix][mk][state_of(crow, ws.title)] += 1
                tx_codes.append(prefix)
                tx_months.append(mk)
                tx_states.append(state_of(crow, ws.title))
```

`is_transaction_row`, `state_of`, `find_incentive_block` y `write_canonical` NO
cambian: operan sobre filas ya canónicas (o sobre el bloque de incentivos crudo,
que no es tabular).

- [ ] **Step 4: Verificar que pasan**

Run: `python -m pytest tests/test_header_aware.py -v`
Expected: PASS (4 tests).

- [ ] **Step 5: Regresión de la raíz**

Run: `python -m pytest tests/ -q`
Expected: todo PASS (los fixtures canónicos existentes mapean idéntico).

- [ ] **Step 6: Commit**

```bash
git add split_excel.py tests/test_header_aware.py
git commit -m "feat: split_excel rearma filas al canónico por encabezado (por hoja)"
```

---

## Task 4: Regresión total + validación end-to-end con el archivo real

**Files:**
- Modify: `kombuchacha/data.js` (regenerado)

- [ ] **Step 1: Suite completa**

Run: `python -m pytest -q`
Expected: todo PASS (raíz + goa-inventory + reporte-sanjose). Ninguna regresión.

- [ ] **Step 2: Regenerar KOM con la fuente real nueva**

`kombuchacha/fuentes/` contiene `KOM-2026-06.xlsx` (canónico) y
`Sales Kombuchacha.xlsx` (layout con Address, sin Company). Run (Git Bash):

```bash
PYTHONUTF8=1 python kombuchacha/build_data.py
```

Expected en la salida:
- `⚠ Sales Kombuchacha.xlsx: columna 'company' no encontrada...` (warning esperado)
- Periodo: Mayo 2026 – Junio 2026 · Meses: 2026-05, 2026-06
- Líneas: 22 · Ingresos: $641.00 · Unidades: 12.5 · Órdenes: 9
- `⚠ Órdenes duplicadas omitidas: S44804`

- [ ] **Step 3: Verificar vendedores-personas (el bug original)**

Run:

```bash
PYTHONUTF8=1 python -c "
import re, json
d = json.loads(re.search(r'const reportData = (.*);', open('kombuchacha/data.js', encoding='utf-8').read(), re.S).group(1))
names = [s['name'] for s in d['salespeople']]
print('vendedores:', names)
assert all(',' not in n for n in names), 'hay direcciones en el ranking!'
assert 'Mireya Fernandez' in names and 'Katherine Osorio Duque' in names and 'Luisa Fernanda' in names
assert d['totals']['revenue'] == 641.0 and d['totals']['units'] == 12.5
assert len(d['months']) == 2
print('OK: vendedores correctos, totales correctos, 2 meses')
"
```

Expected: `OK: vendedores correctos, totales correctos, 2 meses`.

- [ ] **Step 4: Verificación visual del dashboard KOM**

Con el server de preview `kombuchacha-dashboard` (puerto 8138), abrir
`/dashboard.html` y verificar que con 2 meses el dashboard "creció":
- Aparece el selector de mes (Todos / Mayo / Junio) y desaparece la nota de datos iniciales.
- Se dibujan los 4 gráficos (≥2 SKUs, ≥2 vendedores, ≥2 días) — incluida la tendencia.
- Ranking de vendedores: personas, ninguna dirección.
- Sin errores de consola.

- [ ] **Step 5: Commit**

```bash
git add kombuchacha/data.js
git commit -m "data: KOM regenerado con Sales Kombuchacha.xlsx (mayo+junio, header-aware)"
```

---

## Cierre

Tras completar todas las tareas:
- Anunciar: "I'm using the finishing-a-development-branch skill to complete this work."
- **REQUIRED SUB-SKILL:** superpowers:finishing-a-development-branch (merge a `main` → GitHub Pages despliega el KOM actualizado con mayo+junio).

Nota: `robinson-crusoe/data.js`, GOA y San José no cambian. El `.xlsx` nuevo está
gitignoreado como todos los datos fuente.
