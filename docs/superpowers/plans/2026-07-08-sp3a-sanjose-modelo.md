# SP3a — San José v2: modelo de datos — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Corregir y enriquecer el modelo de datos de San José: regla de montos por mes, valuación de promos estable, parseo del bloque de incentivos por estado (descuentos, cross-checks), clientes nuevos calculados, y un `data.js` enriquecido — sin tocar la UI (SP3b).

**Architecture:** `split_excel.py` etiqueta cada bloque de incentivos anexado con su estado de origen (`__INCENTIVOS__`). `generar_data.py` gana: (a) orientación de montos por mes en `load_rows`, (b) valuación de promos con fallback orden→producto-mes→mes (sin pool global), (c) `load_incentivos()` que lee los bloques marcados → descuentos/cross-checks por (mes,estado), (d) cómputo de clientes nuevos por primera aparición, y (e) builders que emiten `revNeto`, `descuentos`, `clientesNuevos`, etc. La regresión se re-basea porque los números cambian a propósito.

**Tech Stack:** Python 3.14, openpyxl 3.1.5, pytest 9.1.1.

**Spec:** [docs/superpowers/specs/2026-07-08-sp3a-sanjose-modelo-design.md](../specs/2026-07-08-sp3a-sanjose-modelo-design.md)

## Estructuras de datos (contrato entre tareas)

- `load_rows()` → lista de dicts de línea (igual que hoy, con `line_rev`/`is_promo`/`box_price` ya corregidos por la orientación por mes).
- `load_incentivos()` → `{ (year,month): { estado: {"descuentos":[{"product","cases","amount"}], "totalDescuentos":float, "freeReportado":{sp:int}, "soldReportado":{sp:int}, "vendedoresNuevos":[str]} } }`.
- `new_customers_by_month(rows)` → `{ (year,month): [ {"name","state","salesperson"} ] }` (primera aparición).
- `build_period_data(...)` y `build_monthly_state_data(...)` reciben `incentivos` y `new_customers` además de lo actual.

---

## Task 1: `split_excel.py` — etiquetar bloques de incentivos con su estado

**Files:**
- Modify: `split_excel.py` (`collect` guarda estado por bloque; `main` antepone marcador)
- Modify: `tests/test_collect.py` (estructura de `incentives` cambia)

- [ ] **Step 1: Actualizar el test de incentivos en `tests/test_collect.py`**

Reemplazar `test_collect_incentives_attributed_to_get` por:
```python
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
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest tests/test_collect.py::test_collect_incentives_attributed_to_get -v`
Expected: FAIL (hoy `incentives[...]` es lista de filas, no de segmentos con `state`).

- [ ] **Step 3: En `split_excel.py`, cambiar `collect()` para guardar el estado dominante del bloque**

Localizar en `collect()` el bloque que arma incentivos:
```python
            block = find_incentive_block(sheet_rows)
            if block and tx_codes:
                dom_code = Counter(tx_codes).most_common(1)[0][0]
                dom_month = Counter(tx_months).most_common(1)[0][0]
                prev = incentives[dom_code].get(dom_month, [])
                incentives[dom_code][dom_month] = prev + block
            elif block:
                print(f"  ⚠ Bloque de incentivos en '{ws.title}' sin filas de pedido "
                      f"de marca conocida; se omite")
```
y reemplazarlo por (se agrega el estado dominante y se guarda como segmento):
```python
            block = find_incentive_block(sheet_rows)
            if block and tx_codes:
                dom_code = Counter(tx_codes).most_common(1)[0][0]
                dom_month = Counter(tx_months).most_common(1)[0][0]
                dom_state = Counter(tx_states).most_common(1)[0][0]
                seg = {"state": dom_state, "rows": block}
                incentives[dom_code].setdefault(dom_month, []).append(seg)
            elif block:
                print(f"  ⚠ Bloque de incentivos en '{ws.title}' sin filas de pedido "
                      f"de marca conocida; se omite")
```
Y en el bucle de filas de esa misma hoja, junto a `tx_codes`/`tx_months`, acumular el estado por fila. Localizar:
```python
            tx_codes, tx_months = [], []
```
reemplazar por:
```python
            tx_codes, tx_months, tx_states = [], [], []
```
y donde se hace `tx_codes.append(prefix)` / `tx_months.append(mk)`, agregar debajo:
```python
                tx_states.append(state_of(row, ws.title))
```

- [ ] **Step 4: En `collect()`, la conversión final de `incentives` ya funciona** (sigue siendo `{code: dict(month)}`, ahora con listas de segmentos). No cambia. Verificar que la línea de retorno de incentives es:
```python
        incentives={c: dict(m) for c, m in incentives.items()},
```

- [ ] **Step 5: En `main()`, anteponer el marcador `__INCENTIVOS__` al escribir cada bloque**

Localizar en `main()`:
```python
            block = data.incentives.get(code, {}).get(mk, [])
            out = ROOT / brand.folder / "fuentes" / f"{brand.code}-{mk}.xlsx"
            write_canonical(out, month_rows, block, brand.sheet)
```
y reemplazarlo por:
```python
            segs = data.incentives.get(code, {}).get(mk, [])
            block = []
            for seg in segs:
                block.append(["__INCENTIVOS__", seg["state"]] + [None] * (len(HEADER) - 2))
                block.extend(seg["rows"])
                block.append([None] * len(HEADER))   # separador entre segmentos
            out = ROOT / brand.folder / "fuentes" / f"{brand.code}-{mk}.xlsx"
            write_canonical(out, month_rows, block, brand.sheet)
```
Y la línea del resumen que marca incentivos:
```python
            inc_txt = " · +incentivos" if block else ""
```
cambiarla a:
```python
            inc_txt = " · +incentivos" if segs else ""
```

- [ ] **Step 6: Correr y verificar que pasa**

Run: `python -m pytest tests/test_collect.py tests/test_main.py -v`
Expected: PASS (incluido el test actualizado y los de `main`, cuyo archivo San José sigue conteniendo `SOLD`/`Alexandra Jimenez` más ahora `__INCENTIVOS__`).

- [ ] **Step 7: Commit**

```bash
git add split_excel.py tests/test_collect.py
git commit -m "feat: split etiqueta bloques de incentivos con su estado (__INCENTIVOS__)"
```

---

## Task 2: `generar_data.py` — regla de montos por mes

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (`load_rows` + helper `_detect_orientation`)
- Test: `reporte-sanjose/tests/test_orientacion.py`

- [ ] **Step 1: Escribir el test que falla** — `reporte-sanjose/tests/test_orientacion.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def _load(tmp_path, monkeypatch, rows):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", rows)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    return g.load_rows()


def test_mes_normal(tmp_path, monkeypatch):
    # Abril normal: Total = qty*UnitPrice
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 4, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 45.6, 182.4),
    ])
    r = rows[0]
    assert r["line_rev"] == 182.4 and r["box_price"] == 45.6 and not r["is_promo"]


def test_mes_swap(tmp_path, monkeypatch):
    # Mayo swap: UnitPrice = qty*Total (intercambiado)
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 5, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 5, 228, 45.6),   # UP=228=5*45.6
        canon_row(D(2026, 5, 3), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 182.4, 45.6),
    ])
    r = [x for x in rows if x["order"] == "S1"][0]
    assert r["line_rev"] == 228 and r["box_price"] == 45.6


def test_caja_gratis_total_cero(tmp_path, monkeypatch):
    # Mes normal, caja gratis con Total=0 pero UnitPrice != 0 → debe ser promo
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 3, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 5, 45.6, 228),   # ancla normal
        canon_row(D(2026, 3, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 3, 45.6, 0),     # Total=0 → gratis
    ])
    gratis = [x for x in rows if x["order"] == "S2"][0]
    assert gratis["is_promo"] is True and gratis["qty"] == 3


def test_venta_con_descuento_inline(tmp_path, monkeypatch):
    # Mes normal, qty1 con Total (41.04) < UnitPrice (45.6): revenue = Total, no max
    rows = _load(tmp_path, monkeypatch, [
        canon_row(D(2026, 2, 1), "S0", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 45.6, 182.4),  # ancla normal
        canon_row(D(2026, 2, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 1, 45.6, 41.04),
    ])
    r = [x for x in rows if x["order"] == "S1"][0]
    assert r["line_rev"] == 41.04 and not r["is_promo"]
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest reporte-sanjose/tests/test_orientacion.py -v`
Expected: FAIL (la regla `max()` actual da line_rev=228→ok en swap por casualidad, pero falla en `test_caja_gratis_total_cero` (max=45.6, no promo) y `test_venta_con_descuento_inline` (max=45.6, no 41.04)).

- [ ] **Step 3: Reescribir `load_rows()` con orientación por mes** en `reporte-sanjose/generar_data.py`.

Agregar este helper antes de `load_rows`:
```python
def _detect_orientation(raw):
    """Dado un grupo de filas crudas del mismo mes (tuplas (qty, total, unit_price)),
    decide si viene 'normal' (Total = qty*UnitPrice) o 'swap' (UnitPrice = qty*Total).
    Usa solo filas inequívocas (qty>1 y ambos montos != 0). Ante la duda, 'normal'."""
    normal = swap = 0
    for qty, total, unit in raw:
        if qty and qty > 1 and total and unit:
            if abs(total - qty * unit) < 0.01:
                normal += 1
            elif abs(unit - qty * total) < 0.01:
                swap += 1
    return "swap" if swap > normal else "normal"
```

Reemplazar el cuerpo de `load_rows()` por una versión en dos pasadas (agrupar por mes → orientación → construir filas):
```python
def load_rows():
    """Lee todos los fuentes/*.xlsx (canónico), detecta la orientación de montos por
    mes y devuelve las filas de pedido limpias. Ignora incentivos/separadores."""
    raw_by_month = defaultdict(list)   # (y,m) -> [(row_tuple, path)]
    unmapped = set()
    for path in _source_files():
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        for row in ws.iter_rows(values_only=True):
            if not _is_tx(row):
                continue
            raw_by_month[(row[0].year, row[0].month)].append(row)

    rows = []
    for ym, month_rows in raw_by_month.items():
        orient = _detect_orientation(
            [(r[6], r[10], r[9]) for r in month_rows])  # (qty_del, total, unit_price)
        for row in month_rows:
            date, order, variant, customer, salesperson, company = row[:6]
            qty_del, unit_price, total = row[6], row[9], row[10]
            state = _state_from_company(company)
            if not state:
                label = company.strip() if isinstance(company, str) and company.strip() else "(sin Company)"
                unmapped.add(label)
                continue
            qty = qty_del or 0
            if qty <= 0:
                continue
            a = total or 0
            b = unit_price or 0
            if orient == "swap":
                line_rev = b            # Unit Price es el total de línea
                box_price = a           # Total es el precio por caja
            else:
                line_rev = a            # normal
                box_price = b
            is_promo = (line_rev == 0)  # caja gratis: total de línea 0 (entrega > 0)
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
    if unmapped:
        print(f"  ⚠ Company sin estado mapeado (filas omitidas): {', '.join(sorted(unmapped))}")
    return rows
```
(Nota: `box_price` ya no se deriva de `line_rev/qty`; ahora es el precio por caja de la columna correcta. Para cajas gratis `box_price` puede venir 0 → la valuación estable de la Tarea 3 aplica el fallback.)

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest reporte-sanjose/tests/test_orientacion.py -v`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_orientacion.py
git commit -m "feat: regla de montos por orientación de mes (arregla descuentos y cajas gratis)"
```

---

## Task 3: `generar_data.py` — valuación de promos estable (fallback por producto-mes)

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (`get_order_unit_prices`, `promo_box_price`, y las llamadas en los builders)
- Test: `reporte-sanjose/tests/test_valuacion_estable.py`

- [ ] **Step 1: Escribir el test que falla** — `reporte-sanjose/tests/test_valuacion_estable.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def _fuentes(tmp_path, extra_junio=False):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    rows_may = [
        # venta pagada mayo (ancla precio del producto en mayo)
        canon_row(D(2026, 5, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 4, 45.6, 182.4),
        # caja gratis mayo, sin precio en su fila (box_price=0) → usa fallback producto-mes
        canon_row(D(2026, 5, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 2, 0, 0),
    ]
    make_canonical(fuentes / "may.xlsx", rows_may)
    if extra_junio:
        make_canonical(fuentes / "jun.xlsx", [
            canon_row(D(2026, 6, 1), "S9", "[GET01] Jack Mackerel in Brine",
                      "C", "SP", "LatinFood Florida", 10, 60.0, 600.0),  # precio distinto
        ])
    return fuentes


def _pv_mayo(fuentes, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    D_ = g.build_monthly_state_data(rows, months, op, pa, ga)
    i = months.index((2026, 5))
    return D_["fl"]["pv"][i]


def test_pv_estable_al_agregar_mes(tmp_path, monkeypatch):
    pv_solo = _pv_mayo(_fuentes(tmp_path / "a", extra_junio=False), monkeypatch)
    pv_con_junio = _pv_mayo(_fuentes(tmp_path / "b", extra_junio=True), monkeypatch)
    # La caja gratis de mayo se valúa con el precio de mayo, no con el pool (que incluiría junio)
    assert pv_solo == pv_con_junio
    assert pv_solo == round(2 * 45.6, 2)   # 2 cajas * precio producto-mes de mayo
```
(Los dos `_fuentes` usan subcarpetas distintas `a`/`b` para aislar.)

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest reporte-sanjose/tests/test_valuacion_estable.py -v`
Expected: FAIL (hoy `product_avg` es global sobre todo el pool → junio cambia el pv de mayo).

- [ ] **Step 3: Cambiar la valuación a por-producto-por-mes** en `reporte-sanjose/generar_data.py`.

Reemplazar `get_order_unit_prices` y `promo_box_price` por:
```python
def get_order_unit_prices(rows):
    """Precios para valuar cajas gratis, SIN pool global (estable al agregar meses):
      - order_prices[order]         : precio de caja de una línea pagada del mismo orden
      - product_month_avg[(prod,ym)]: promedio de precio de caja del producto en ese mes
      - month_avg[ym]               : promedio del mes (último recurso)
    """
    order_prices = defaultdict(float)
    pm = defaultdict(list)
    m = defaultdict(list)
    for r in rows:
        if not r['is_promo'] and r['box_price'] > 0:
            order_prices[r['order']] = r['box_price']
            ym = month_key(r['date'])
            pm[(r['product'], ym)].append(r['box_price'])
            m[ym].append(r['box_price'])
    product_month_avg = {k: sum(v) / len(v) for k, v in pm.items()}
    month_avg = {k: sum(v) / len(v) for k, v in m.items()}
    return order_prices, product_month_avg, month_avg


def promo_box_price(r, order_prices, product_month_avg, month_avg):
    """Precio por caja para valuar una caja gratis: orden → producto-mes → mes → 0."""
    if r['box_price'] > 0:
        return r['box_price']
    ym = month_key(r['date'])
    return (order_prices.get(r['order'])
            or product_month_avg.get((r['product'], ym))
            or month_avg.get(ym)
            or 0.0)
```
Las firmas de `build_monthly_state_data(rows, months, order_prices, product_month_avg, month_avg)` y `build_period_data(...)` no cambian de forma (siguen recibiendo 3 objetos de precios); solo cambian los nombres internos. **No hace falta editar sus cuerpos** salvo que usen los parámetros por nombre — verificar que las llamadas internas a `promo_box_price(r, order_prices, product_avg, global_avg)` siguen pasando los 3 en orden (los renombres son transparentes posicionalmente). Renombrar los parámetros en las firmas para claridad es opcional.

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest reporte-sanjose/tests/test_valuacion_estable.py -v`
Expected: PASS (2 asserts).

- [ ] **Step 5: Commit**

```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_valuacion_estable.py
git commit -m "feat: valuación de promos estable (fallback orden→producto-mes→mes, sin pool global)"
```

---

## Task 4: `generar_data.py` — `load_incentivos()` (parseo del bloque por estado)

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (nuevo `load_incentivos` + helpers)
- Test: `reporte-sanjose/tests/test_incentivos.py`

- [ ] **Step 1: Escribir el test que falla** — `reporte-sanjose/tests/test_incentivos.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g

HEADER_LEN = 11


def _pad(row):
    return list(row) + [None] * (HEADER_LEN - len(row))


def _fuentes_con_bloque(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    # Una venta real de junio + el bloque de incentivos marcado con estado Florida
    rows = [
        canon_row(D(2026, 6, 22), "S1", "[GET01] Jack Mackerel in Brine",
                  "Key Food", "Alexandra J", "LatinFood Florida", 20, 45.6, 912),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 5, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68]),
        _pad(["Angela Quimbay", 5, None, "Tomate == > 18 cases", 540]),
        _pad(["Lina Raquel", 50, 11, None, 628.68]),
        _pad([None, None, 21, "Clientes Nuevos", 15]),
        _pad(["NEW CUSTOMERS"]),
        _pad(["Alexandra J"]),
        _pad(["Angela Q"]),
        _pad(["Raquel L", 1]),
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    return fuentes


def test_load_incentivos(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", _fuentes_con_bloque(tmp_path))
    inc = g.load_incentivos()
    fl = inc[(2026, 6)]["Florida"]
    # descuentos por producto
    productos = {d["product"]: d for d in fl["descuentos"]}
    assert productos["In Brine"]["cases"] == 3 and productos["In Brine"]["amount"] == 88.68
    assert productos["Tomate"]["cases"] == 18 and productos["Tomate"]["amount"] == 540
    # total de descuentos (suma de montos de descuento por producto)
    assert round(fl["totalDescuentos"], 2) == 628.68
    # cross-check free/sold reportados
    assert fl["freeReportado"]["Alexandra Jimenez"] == 10
    assert fl["soldReportado"]["Lina Raquel"] == 50
    # vendedores con clientes nuevos
    assert "Alexandra J" in fl["vendedoresNuevos"]


def test_load_incentivos_vacio(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 6, 1), "S1", "[GET01] x", "C", "SP", "LatinFood Florida", 1, 45.6, 45.6),
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    assert g.load_incentivos() == {}
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest reporte-sanjose/tests/test_incentivos.py -v`
Expected: FAIL con `AttributeError: module 'generar_data' has no attribute 'load_incentivos'`.

- [ ] **Step 3: Agregar `load_incentivos` y helpers** a `reporte-sanjose/generar_data.py` (tras `load_rows`):
```python
def _parse_discount_label(text):
    """'In Brine == > 3 cases' -> ('In Brine', 3). None si no reconoce producto."""
    t = str(text).lower()
    product = None
    if "in brine" in t:
        product = "In Brine"
    elif "tomate" in t or "tomato" in t:
        product = "Tomate"
    if product is None:
        return None
    mnum = re.search(r"(\d+)\s*cases", t)
    cases = int(mnum.group(1)) if mnum else 0
    return product, cases


def _last_number(row):
    for cell in reversed(list(row)):
        if isinstance(cell, (int, float)) and not isinstance(cell, bool) and cell == cell:
            return float(cell)
    return None


def load_incentivos():
    """Lee los bloques `__INCENTIVOS__` de los fuentes/ y los parsea por (mes, estado).
    Defensivo: el bloque es manual e irregular; ante lo que no reconoce, deja vacío."""
    out = {}
    for path in _source_files():
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        allrows = list(ws.iter_rows(values_only=True))
        # mes dominante de las filas de pedido de este archivo
        tx_months = [(r[0].year, r[0].month) for r in allrows if _is_tx(r)]
        if not tx_months:
            continue
        dom_month = max(set(tx_months), key=tx_months.count)
        # ubicar segmentos marcados
        i = 0
        while i < len(allrows):
            row = allrows[i]
            if isinstance(row[0], str) and row[0].strip() == "__INCENTIVOS__":
                state = row[1]
                seg = []
                i += 1
                while i < len(allrows) and not (
                        isinstance(allrows[i][0], str) and allrows[i][0].strip() == "__INCENTIVOS__"):
                    seg.append(allrows[i])
                    i += 1
                parsed = _parse_incentive_segment(seg)
                out.setdefault(dom_month, {})[state] = parsed
            else:
                i += 1
    return out


def _parse_incentive_segment(seg):
    """Parsea un segmento de incentivos en descuentos / free / sold / vendedores nuevos."""
    descuentos = []
    free_rep, sold_rep = {}, {}
    vendedores_nuevos = []
    in_new_customers = False
    for row in seg:
        col_a = row[0].strip() if isinstance(row[0], str) else row[0]
        if isinstance(col_a, str) and col_a.upper() == "NEW CUSTOMERS":
            in_new_customers = True
            continue
        if in_new_customers:
            if isinstance(col_a, str) and col_a:
                vendedores_nuevos.append(col_a)
            continue
        # fila de vendedor con SOLD/FREE (col B, col C numéricos) + posible descuento
        if isinstance(col_a, str) and col_a and col_a.upper() not in ("SOLD",):
            sold = row[1] if isinstance(row[1], (int, float)) and not isinstance(row[1], bool) else None
            free = row[2] if isinstance(row[2], (int, float)) and not isinstance(row[2], bool) else None
            if sold is not None:
                sold_rep[col_a] = int(sold)
            if free is not None:
                free_rep[col_a] = int(free)
        # descuento por producto: buscar etiqueta reconocible en cualquier celda + monto al final
        for cell in row:
            lab = _parse_discount_label(cell) if isinstance(cell, str) else None
            if lab:
                amount = _last_number(row) or 0.0
                descuentos.append({"product": lab[0], "cases": lab[1], "amount": round(amount, 2)})
                break
    total = round(sum(d["amount"] for d in descuentos), 2)
    return {
        "descuentos": descuentos,
        "totalDescuentos": total,
        "freeReportado": free_rep,
        "soldReportado": sold_rep,
        "vendedoresNuevos": vendedores_nuevos,
    }
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest reporte-sanjose/tests/test_incentivos.py -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_incentivos.py
git commit -m "feat: load_incentivos — descuentos por producto/estado + cross-checks del bloque"
```

---

## Task 5: `generar_data.py` — clientes nuevos por primera aparición

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (`new_customers_by_month`)
- Test: `reporte-sanjose/tests/test_clientes_nuevos.py`

- [ ] **Step 1: Escribir el test que falla** — `reporte-sanjose/tests/test_clientes_nuevos.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def test_new_customers(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 4, 1), "S1", "[GET01] x", "Cliente A", "SP", "LatinFood Florida", 1, 45.6, 45.6),
        canon_row(D(2026, 5, 1), "S2", "[GET01] x", "Cliente A", "SP", "LatinFood Florida", 1, 45.6, 45.6),
        canon_row(D(2026, 5, 2), "S3", "[GET01] x", "Cliente B", "Ana", "LatinFood Florida", 1, 45.6, 45.6),
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    nc = g.new_customers_by_month(rows)
    # Cliente A aparece por primera vez en abril (primer mes → NO cuenta como nuevo)
    assert (2026, 4) not in nc or nc[(2026, 4)] == []
    # En mayo, Cliente B es nuevo; Cliente A no
    nombres_may = {c["name"] for c in nc.get((2026, 5), [])}
    assert nombres_may == {"Cliente B"}
    assert nc[(2026, 5)][0]["salesperson"] == "Ana"
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest reporte-sanjose/tests/test_clientes_nuevos.py -v`
Expected: FAIL con `AttributeError: ... 'new_customers_by_month'`.

- [ ] **Step 3: Agregar `new_customers_by_month`** a `reporte-sanjose/generar_data.py`:
```python
def new_customers_by_month(rows):
    """Clientes que aparecen por primera vez en cada mes (excluye el primer mes cargado,
    que no tiene historia previa). Devuelve {(y,m): [{name, state, salesperson}]}."""
    months = sorted(set(month_key(r['date']) for r in rows))
    if not months:
        return {}
    first_month = months[0]
    seen = set()
    # normaliza el nombre para comparar
    def norm(name):
        return re.sub(r"\s+", " ", str(name)).strip().lower()
    # primera aparición cronológica de cada cliente
    first_seen = {}
    for r in sorted(rows, key=lambda r: r['date']):
        key = norm(r['customer'])
        if key and key not in first_seen:
            first_seen[key] = r
    out = {}
    for key, r in first_seen.items():
        ym = month_key(r['date'])
        if ym == first_month:
            continue
        out.setdefault(ym, []).append({
            "name": r['customer'], "state": r['state'], "salesperson": r['salesperson'],
        })
    return out
```

- [ ] **Step 4: Correr y verificar que pasa**

Run: `python -m pytest reporte-sanjose/tests/test_clientes_nuevos.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_clientes_nuevos.py
git commit -m "feat: clientes nuevos por primera aparición (calculado del histórico)"
```

---

## Task 6: `generar_data.py` — enriquecer builders y `main` (revNeto, descuentos, clientes nuevos)

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (`build_period_data`, `build_monthly_state_data`, `main` + template JS)
- Test: `reporte-sanjose/tests/test_enriquecido.py`

- [ ] **Step 1: Escribir el test que falla** — `reporte-sanjose/tests/test_enriquecido.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g

HEADER_LEN = 11
def _pad(row): return list(row) + [None] * (HEADER_LEN - len(row))


def _fuentes(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    rows = [
        canon_row(D(2026, 6, 22), "S1", "[GET01] Jack Mackerel in Brine",
                  "Key Food", "Alexandra J", "LatinFood Florida", 20, 45.6, 912),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 5, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 60, 10, "In Brine == > 3 cases", 88.68]),
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    return fuentes


def test_period_incluye_neto_y_descuentos(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "SOURCES_DIR", _fuentes(tmp_path))
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    inc = g.load_incentivos()
    nc = g.new_customers_by_month(rows)
    period = g.build_period_data(rows, months, op, pa, ga, inc, nc, month_set=set(months))

    assert period["summary"]["descuentos"] == 88.68
    assert period["summary"]["revNeto"] == round(912 - 88.68, 2)
    assert period["states"]["Florida"]["descuentos"] == 88.68
    assert period["states"]["Florida"]["revNeto"] == round(912 - 88.68, 2)
    prods = {d["product"]: d for d in period["descuentosPorProducto"]}
    assert prods["In Brine"]["amount"] == 88.68 and prods["In Brine"]["state"] == "Florida"
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m pytest reporte-sanjose/tests/test_enriquecido.py -v`
Expected: FAIL (build_period_data no acepta `incentivos`/`new_customers` ni emite `revNeto`).

- [ ] **Step 3: Enriquecer `build_period_data`.** Cambiar su firma y agregar el cálculo de descuentos/neto/clientes-nuevos.

Cambiar la firma:
```python
def build_period_data(rows, months, order_prices, product_avg, global_avg, month_set=None):
```
por:
```python
def build_period_data(rows, months, order_prices, product_avg, global_avg,
                      incentivos=None, new_customers=None, month_set=None):
```
Al inicio del cuerpo (después de calcular `month_set`), agregar el agregado de descuentos por estado/producto sobre los meses del periodo:
```python
    incentivos = incentivos or {}
    new_customers = new_customers or {}

    # Descuentos del periodo, por estado y por producto (de los bloques de incentivos)
    desc_by_state = defaultdict(float)
    desc_by_product = defaultdict(lambda: {"cases": 0, "amount": 0.0, "state": ""})
    for ym in month_set:
        for state, inc in incentivos.get(ym, {}).items():
            desc_by_state[state] += inc["totalDescuentos"]
            for d in inc["descuentos"]:
                k = (d["product"], state)
                desc_by_product[k]["cases"] += d["cases"]
                desc_by_product[k]["amount"] = round(desc_by_product[k]["amount"] + d["amount"], 2)
                desc_by_product[k]["state"] = state
    total_descuentos = round(sum(desc_by_state.values()), 2)

    # Clientes nuevos del periodo
    nuevos = [c for ym in month_set for c in new_customers.get(ym, [])]
```
En el bloque `summary`, agregar tras `avg_price`:
```python
        'descuentos': total_descuentos,
        'revNeto': round(total_rev - total_descuentos, 2),
        'clientesNuevos': len(nuevos),
```
En el armado de cada `states[state_name]`, agregar `descuentos`/`revNeto`. Reemplazar el bloque:
```python
    states = {}
    for state_name, d in state_data.items():
        states[state_name] = {
            'rev': round(d['rev'], 2),
            'cajas': d['cajas'],
            'pc': d['pc'],
            'pv': round(d['pv'], 2),
            'ord': len(d['orders'])
        }
    if 'Florida' not in states:
        states['Florida'] = {'rev': 0.0, 'cajas': 0, 'pc': 0, 'pv': 0.0, 'ord': 0}
    if 'Nueva York' not in states:
        states['Nueva York'] = {'rev': 0.0, 'cajas': 0, 'pc': 0, 'pv': 0.0, 'ord': 0}
```
por:
```python
    states = {}
    for state_name, d in state_data.items():
        dd = round(desc_by_state.get(state_name, 0.0), 2)
        states[state_name] = {
            'rev': round(d['rev'], 2),
            'cajas': d['cajas'],
            'pc': d['pc'],
            'pv': round(d['pv'], 2),
            'ord': len(d['orders']),
            'descuentos': dd,
            'revNeto': round(d['rev'] - dd, 2),
        }
    for missing in ('Florida', 'Nueva York'):
        if missing not in states:
            dd = round(desc_by_state.get(missing, 0.0), 2)
            states[missing] = {'rev': 0.0, 'cajas': 0, 'pc': 0, 'pv': 0.0, 'ord': 0,
                               'descuentos': dd, 'revNeto': round(-dd, 2)}
```
En el `return`, agregar dos claves:
```python
        'descuentosPorProducto': [
            {"product": p, "state": st, "cases": v["cases"], "amount": v["amount"]}
            for (p, st), v in desc_by_product.items()
        ],
        'clientesNuevosLista': nuevos,
```

- [ ] **Step 4: Enriquecer `build_monthly_state_data`** para agregar series `descuentos` y `revNeto` por mes. Cambiar su firma a aceptar `incentivos`:
```python
def build_monthly_state_data(rows, months, order_prices, product_avg, global_avg, incentivos=None):
```
Y en el armado de `result[state]`, tras las series existentes, agregar por mes las series de descuentos/neto. Reemplazar:
```python
    result = {}
    for state in ('fl', 'ny'):
        result[state] = {
            'rev': [], 'cajas': [], 'pv': [], 'pc': [], 'ord': []
        }
        for ym in months:
            d = state_month[state][ym]
            result[state]['rev'].append(round(d['rev'], 2))
            result[state]['cajas'].append(d['cajas'])
            result[state]['pv'].append(round(d['pv'], 2))
            result[state]['pc'].append(d['pc'])
            result[state]['ord'].append(len(d['orders']))
    return result
```
por:
```python
    incentivos = incentivos or {}
    state_name = {'fl': 'Florida', 'ny': 'Nueva York'}
    result = {}
    for state in ('fl', 'ny'):
        result[state] = {'rev': [], 'cajas': [], 'pv': [], 'pc': [], 'ord': [],
                         'descuentos': [], 'revNeto': []}
        for ym in months:
            d = state_month[state][ym]
            dd = round(incentivos.get(ym, {}).get(state_name[state], {}).get('totalDescuentos', 0.0), 2)
            result[state]['rev'].append(round(d['rev'], 2))
            result[state]['cajas'].append(d['cajas'])
            result[state]['pv'].append(round(d['pv'], 2))
            result[state]['pc'].append(d['pc'])
            result[state]['ord'].append(len(d['orders']))
            result[state]['descuentos'].append(dd)
            result[state]['revNeto'].append(round(d['rev'] - dd, 2))
    return result
```

- [ ] **Step 5: Actualizar `main()`** para cargar incentivos + clientes nuevos y pasarlos a los builders. Reemplazar:
```python
    order_prices, product_avg, global_avg = get_order_unit_prices(rows)

    # Build D (time series)
    D = build_monthly_state_data(rows, all_months, order_prices, product_avg, global_avg)

    # Build GENERAL_DATA
    general_data = {}

    # ALL
    general_data['ALL'] = build_period_data(rows, all_months, order_prices, product_avg, global_avg,
                                            month_set=set(all_months))

    # Per month
    for ym in all_months:
        label = month_label_short(ym)
        general_data[label] = build_period_data(rows, all_months, order_prices, product_avg, global_avg,
                                                month_set={ym})
```
por:
```python
    order_prices, product_avg, global_avg = get_order_unit_prices(rows)
    incentivos = load_incentivos()
    new_customers = new_customers_by_month(rows)

    # Build D (time series)
    D = build_monthly_state_data(rows, all_months, order_prices, product_avg, global_avg, incentivos)

    # Build GENERAL_DATA
    general_data = {}

    # ALL
    general_data['ALL'] = build_period_data(rows, all_months, order_prices, product_avg, global_avg,
                                            incentivos, new_customers, month_set=set(all_months))

    # Per month
    for ym in all_months:
        label = month_label_short(ym)
        general_data[label] = build_period_data(rows, all_months, order_prices, product_avg, global_avg,
                                                incentivos, new_customers, month_set={ym})
```

- [ ] **Step 6: Correr y verificar que pasa**

Run: `python -m pytest reporte-sanjose/tests/test_enriquecido.py -v`
Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_enriquecido.py
git commit -m "feat: data.js enriquecido — revNeto, descuentos por producto/estado, clientes nuevos"
```

---

## Task 7: Re-basear la regresión + verificación integral

**Files:**
- Modify: `reporte-sanjose/tests/test_regresion.py` (nueva base de números corregidos)
- Delete/replace: `reporte-sanjose/tests/reference_pre_sp2.json` → `reference_sp3.json`

- [ ] **Step 1: Correr la migración y el generador reales para producir el data.js corregido**

Run: `python reporte-sanjose/migrar_historico.py && python split_excel.py`
Expected: San José regenerado con la lógica nueva. Anotar del resumen que no hay errores.

- [ ] **Step 2: Capturar la nueva referencia (números corregidos)**

Run:
```bash
python -c "import re,json; t=open('reporte-sanjose/data.js',encoding='utf-8').read(); g=lambda n,pat: json.loads(re.search(r'\b'+n+r' = ('+pat+r');',t,re.S).group(1)); ref={'MONTHS':g('MONTHS',r'\[.*?\]'),'D':g('D',r'\{.*?\}'),'GENERAL_DATA':g('GENERAL_DATA',r'\{.*\}')}; open('reporte-sanjose/tests/reference_sp3.json','w',encoding='utf-8').write(json.dumps(ref,ensure_ascii=False,indent=2)); print('meses:', ref['MONTHS'])"
```
Expected: 9 meses (Oct 25 … Jun 26).

- [ ] **Step 3: Reescribir `reporte-sanjose/tests/test_regresion.py`** para guardar contra la nueva base y validar las correcciones. Reemplazar todo el archivo por:
```python
"""
Regresión de SP3a. La base ahora son los números CORREGIDOS (reference_sp3.json).
- Guarda contra cambios accidentales futuros (todos los meses).
- Verifica que junio trae descuentos → revNeto, y que hay clientes nuevos.
"""
import json
import re
from pathlib import Path

import pytest

import generar_data as g

SANJOSE = Path(__file__).resolve().parent.parent
SRC_HISTORICO = SANJOSE / "Historico Ventas Latin Food - San Jose.xlsx"
FUENTES = SANJOSE / "fuentes"
REF = json.load(open(Path(__file__).parent / "reference_sp3.json", encoding="utf-8"))


def _parse_data_js(text):
    def grab(name, pat):
        return json.loads(re.search(r"\b" + name + r" = (" + pat + r");", text, re.S).group(1))
    return {"MONTHS": grab("MONTHS", r"\[.*?\]"),
            "D": grab("D", r"\{.*?\}"),
            "GENERAL_DATA": grab("GENERAL_DATA", r"\{.*\}")}


@pytest.mark.skipif(not (FUENTES / "San-Jose-historico-hasta-2026-05.xlsx").exists(),
                    reason="faltan fuentes reales (correr migrar_historico.py + split_excel.py)")
def test_regresion_vs_base_sp3(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "OUTPUT_FILE", tmp_path / "data.js")
    g.main()
    out = _parse_data_js((tmp_path / "data.js").read_text(encoding="utf-8"))
    assert out["MONTHS"] == REF["MONTHS"]
    assert out["GENERAL_DATA"] == REF["GENERAL_DATA"]
    assert out["D"] == REF["D"]


@pytest.mark.skipif(not (FUENTES / "San-Jose-2026-06.xlsx").exists(),
                    reason="falta el fuentes de junio")
def test_junio_tiene_descuentos_y_neto(tmp_path, monkeypatch):
    monkeypatch.setattr(g, "OUTPUT_FILE", tmp_path / "data.js")
    g.main()
    out = _parse_data_js((tmp_path / "data.js").read_text(encoding="utf-8"))
    jun = out["GENERAL_DATA"]["Jun 26"]["summary"]
    assert jun["descuentos"] > 0
    assert jun["revNeto"] == round(jun["rev"] - jun["descuentos"], 2)
```

- [ ] **Step 4: Borrar el fixture viejo**

Run: `git rm reporte-sanjose/tests/reference_pre_sp2.json`
Y en `reporte-sanjose/.gitignore`, cambiar la excepción `!tests/reference_pre_sp2.json` por `!tests/reference_sp3.json`.

- [ ] **Step 5: Correr toda la suite**

Run: `python -m pytest -q`
Expected: todo verde.

- [ ] **Step 6: Commit**

```bash
git add reporte-sanjose/tests/test_regresion.py reporte-sanjose/tests/reference_sp3.json reporte-sanjose/.gitignore
git commit -m "test: re-basear regresión de San José a los números corregidos (SP3a)"
```

---

## Task 8: Aceptación — verificar el efecto del arreglo y publicar el data.js

**Files:**
- Verify + commit: `reporte-sanjose/data.js`

- [ ] **Step 1: Regenerar con el pipeline real**

Run: `python split_excel.py`
Expected: San José regenerado ✓, sin SKUs no mapeados.

- [ ] **Step 2: Verificar el efecto del arreglo de la regla** (revenue bruto baja algo y cajas promo suben, por las cajas gratis antes mal contadas)

Run:
```bash
python -c "import re,json; t=open('reporte-sanjose/data.js',encoding='utf-8').read(); d=json.loads(re.search(r'\bGENERAL_DATA = (\{.*\});',t,re.S).group(1)); a=d['ALL']['summary']; print('rev',a['rev'],'revNeto',a['revNeto'],'descuentos',a['descuentos'],'cajas',a['cajas'],'promo',a['pc'],'clientesNuevos',a['clientesNuevos'])"
```
Expected: `descuentos` > 0 (incluye los $628.68 de junio), `revNeto` = rev − descuentos, y valores coherentes.

- [ ] **Step 3: Verificar clientes nuevos y descuentos por producto de junio**

Run:
```bash
python -c "import re,json; t=open('reporte-sanjose/data.js',encoding='utf-8').read(); d=json.loads(re.search(r'\bGENERAL_DATA = (\{.*\});',t,re.S).group(1))['Jun 26']; print('descuentosPorProducto:', d['descuentosPorProducto']); print('clientesNuevos:', d['summary']['clientesNuevos'])"
```
Expected: In Brine y Tomate con sus montos; un conteo de clientes nuevos calculado (no el "15" manual).

- [ ] **Step 4: Verificar git y suite**

Run: `python -m pytest -q` (verde) y `git status --porcelain | grep -i '.xlsx' || echo "OK sin xlsx"`.

- [ ] **Step 5: Commit del data.js**

```bash
git add reporte-sanjose/data.js
git commit -m "data: San José v2 — montos corregidos, revenue neto, descuentos y clientes nuevos"
```

---

## Verificación final

- [ ] `python -m pytest -q` → todo verde (orientación, valuación estable, incentivos, clientes nuevos, enriquecido, regresión).
- [ ] `data.js` de San José: 9 meses, `revNeto`/`descuentos`/`clientesNuevos` presentes; `pv` estable (agregar un mes no mueve meses previos).
- [ ] El `index.html` actual sigue cargando (muestra números corregidos; los campos nuevos quedan disponibles para SP3b).
- [ ] Ningún `.xlsx` trackeado.

## Fuera de alcance (SP3b y más)

- **SP3b:** rediseño del dashboard para mostrar revenue neto, descuentos por producto/estado, clientes nuevos y cross-checks (con mockups en el navegador).
- Dashboards KOM/ROB; **SP4** (repo GetUp + subdominio privado).
