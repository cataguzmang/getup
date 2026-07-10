# Dashboards Kombuchacha y Robinson Crusoe — Plan de Implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dar de alta los dashboards de Kombuchacha (`KOM`) y Robinson Crusoe (`ROB`) con un core de datos compartido y un front clonado de GOA que degrada con gracia cuando hay pocos datos.

**Architecture:** Un módulo compartido `report_core.py` (generalización del generador de GOA, parametrizado por marca/prefijo-SKU/hoja) hace todos los cálculos. Cada marca tiene un `build_data.py` delgado que lo invoca y un `dashboard.html` clonado de GOA con reglas de mostrar/ocultar. Se marca KOM/ROB como `sp1_ready=True` para que `split_excel.py` los regenere solos cada mes. GOA y San José no se tocan.

**Tech Stack:** Python 3.14 + openpyxl (generador), HTML + Chart.js 4.4 (dashboard, un solo archivo), pytest.

**Regla del proyecto (obligatoria):** todo cálculo derivado vive en `report_core.py` en funciones explícitas; el `dashboard.html` solo formatea/renderiza y decide mostrar/ocultar bloques a partir de datos ya calculados.

---

## File Structure

- **Crear** `report_core.py` (raíz) — lógica de cálculo compartida y parametrizable.
- **Crear** `tests/test_report_core.py` (raíz) — tests del core con fixtures sintéticos.
- **Crear** `kombuchacha/build_data.py` — wrapper KOM.
- **Crear** `robinson-crusoe/build_data.py` — wrapper ROB.
- **Crear** `kombuchacha/dashboard.html` — clon de GOA con degradación (referencia).
- **Crear** `robinson-crusoe/dashboard.html` — copia de KOM re-marcada.
- **Crear** `kombuchacha/data.js`, `robinson-crusoe/data.js` — generados (se commitean).
- **Modificar** `split_excel.py:42-43` — KOM/ROB pasan a `sp1_ready=True`.
- **Modificar** `tests/test_main.py:61` — actualizar aserción (KOM ya no está "pendiente").
- **Modificar** `.claude/launch.json` — servidores de preview para KOM y ROB.

---

## Task 1: `report_core.py` — core de datos compartido

**Files:**
- Create: `report_core.py`
- Test: `tests/test_report_core.py`

El core es la generalización de `goa-inventory/parse_excel.py`: mismo objeto `reportData`, pero (a) parametrizado por marca/prefijo/hoja y (b) **preservando cantidades fraccionarias** (KOM vende medias cajas; `int()` las volvería 0).

- [ ] **Step 1: Escribir los tests que fallan**

Create `tests/test_report_core.py`:

```python
import datetime
import openpyxl

import report_core as rc

HEADER = ["Order Date", "Order", "Product Variant", "Customer", "Salesperson",
          "Company", "Qty Delivered", "Qty Invoiced", "Qty Ordered", "Unit Price", "Total"]
dt = datetime.datetime


def _xlsx(path, rows, sheet="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def _line(date, order, variant, customer, sp, qty=1, price=21.85, total=21.85,
          company="LatinFood US Corp."):
    return [date, order, variant, customer, sp, company, qty, qty, qty, price, total]


def test_sku_pattern_por_prefijo():
    pk = rc.sku_pattern("KOM")
    assert rc.split_sku("[KOM01] Kombuchacha Zero Blueberry", pk) == \
        ("KOM01", "Kombuchacha Zero Blueberry")
    # un prefijo distinto no matchea con el patrón de KOM
    assert rc.split_sku("[ROB05] Mussel en Brine", pk)[0] is None


def test_parse_transactions_preserva_cantidad_fraccionaria():
    pk = rc.sku_pattern("KOM")
    rows = [HEADER, _line(dt(2026, 6, 18), "S1", "[KOM01] Zero Blueberry",
                          "Tienda", "Mireya", qty=0.5, price=50.64, total=25.32)]
    lines = rc.parse_transactions(rows, pk)
    assert lines[0]["qtyOrdered"] == 0.5     # NO truncado a 0
    assert lines[0]["total"] == 25.32


def test_read_sheet_rows_respeta_sheet_candidates(tmp_path):
    p = _xlsx(tmp_path / "s.xlsx",
              [HEADER, _line(dt(2026, 6, 1), "S1", "[KOM01] X", "T", "A", qty=0.5)],
              sheet="KOM")
    rows = rc.read_sheet_rows(p, ("KOM", "Sheet1"))
    assert rows[0][0] == "Order Date" and rows[1][1] == "S1"


def test_build_report_kom_shape(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "KOM-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 18), "S44804", "[KOM01] Kombuchacha Zero Blueberry",
              "Convenience store flora llc", "Mireya Fernandez", qty=0.5, price=50.64, total=25.32),
        _line(dt(2026, 6, 18), "S44804", "[KOM02] Kombuchacha Zero Raspberry",
              "Convenience store flora llc", "Mireya Fernandez", qty=0.5, price=50.64, total=25.32),
    ], sheet="KOM")
    rep = rc.build_report(tmp_path, "Kombuchacha", "LatinFood US Corp", "KOM", ("KOM", "Sheet1"))
    assert len(rep["products"]) == 2
    assert rep["totals"]["units"] == 1.0
    assert rep["totals"]["revenue"] == 50.64
    assert rep["totals"]["customers"] == 1
    assert len(rep["salespeople"]) == 1
    assert rep["meta"]["brand"] == "Kombuchacha"
    assert len(rep["months"]) == 1


def test_build_report_rob_shape(tmp_path):
    fuentes = tmp_path / "fuentes"
    fuentes.mkdir()
    _xlsx(fuentes / "ROB-2026-06.xlsx", [HEADER,
        _line(dt(2026, 6, 1), "S43486", "[ROB05] Mussel en Brine /Mejillones",
              "Gala foods", "Luis Soler", qty=1, price=60, total=60),
    ], sheet="ROB")
    rep = rc.build_report(tmp_path, "Robinson Crusoe", "LatinFood US Corp", "ROB", ("ROB", "Sheet1"))
    assert len(rep["products"]) == 1
    assert rep["totals"]["revenue"] == 60.0
    assert rep["totals"]["units"] == 1
    assert rep["meta"]["distributor"] == "LatinFood US Corp"
```

- [ ] **Step 2: Correr los tests para verificar que fallan**

Run: `python -m pytest tests/test_report_core.py -v`
Expected: FAIL con `ModuleNotFoundError: No module named 'report_core'`.

- [ ] **Step 3: Escribir `report_core.py`**

Create `report_core.py`:

```python
"""
report_core.py — Núcleo compartido para generar el `data.js` de una marca a partir
de su carpeta `fuentes/` (formato canónico de 11 columnas del distribuidor).

Es la generalización del generador de GOA (`goa-inventory/parse_excel.py`):
mismo objeto `reportData`, pero parametrizado por marca / prefijo de SKU / hoja, y
**preservando cantidades fraccionarias** (algunas marcas venden medias cajas: 0.5).

Todo cálculo derivado vive aquí en funciones explícitas; el dashboard solo formatea.

Uso desde el wrapper de una marca:

    import report_core
    report_core.build_and_write(
        here=Path(__file__).parent,
        brand="Kombuchacha", distributor="LatinFood US Corp",
        sku_prefix="KOM", sheet_candidates=("KOM", "Sheet1"))
"""

import re
import json
import datetime
from pathlib import Path

import openpyxl

# Índices de columnas (0-based) en el bloque transaccional canónico
COL_DATE, COL_ORDER, COL_VARIANT, COL_CUSTOMER, COL_SALESPERSON = 0, 1, 2, 3, 4
COL_COMPANY, COL_QTY_DELIVERED, COL_QTY_INVOICED, COL_QTY_ORDERED = 5, 6, 7, 8
COL_UNIT_PRICE, COL_TOTAL = 9, 10

MESES_ES = [
    "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de limpieza / parametrización
# ─────────────────────────────────────────────────────────────────────────────
def num(value):
    """Convierte a número (float) de forma segura. None/texto/bool/NaN -> 0."""
    if value is None or isinstance(value, bool):
        return 0
    if isinstance(value, (int, float)):
        return 0 if value != value else value  # descarta NaN
    return 0


def q(value):
    """Cantidad: número redondeado a 2 decimales (preserva 0.5, evita ruido float)."""
    return round(num(value), 2)


def clean_str(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def sku_pattern(prefix):
    """Patrón para '[PREFIJO<n>] Nombre' -> ('PREFIJO<n>', 'Nombre')."""
    return re.compile(rf"^\s*\[({re.escape(prefix)}\d+)\]\s*(.*)$")


def split_sku(variant, pattern):
    m = pattern.match(str(variant))
    if not m:
        return None, clean_str(variant)
    return m.group(1), clean_str(m.group(2))


def month_key(date_iso):
    return date_iso[:7]


def month_label(key):
    y, m = key.split("-")
    return f"{MESES_ES[int(m)].capitalize()} {y}"


def list_source_files(folder):
    folder = Path(folder)
    if not folder.exists():
        return []
    return sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))


def read_sheet_rows(path, sheet_candidates):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = None
    for name in sheet_candidates:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        ws = wb.active
    return list(ws.iter_rows(values_only=True))


def is_transaction_row(row, pattern):
    return (
        isinstance(row[COL_DATE], datetime.datetime)
        and isinstance(row[COL_VARIANT], str)
        and pattern.match(row[COL_VARIANT])
    )


def parse_transactions(rows, pattern):
    """Líneas de pedido limpias (dicts). Cantidades preservadas como float."""
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


# ─────────────────────────────────────────────────────────────────────────────
# Incentivos (bloque manual del final; tolerante a fallos) — igual que GOA
# ─────────────────────────────────────────────────────────────────────────────
def empty_incentives():
    return {"freeUnitsBySalesperson": [], "costItems": [], "totalDescuentos": 0.0}


def merge_incentives(incs):
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
    counts = {}
    for ln in lines:
        counts[ln["month"]] = counts.get(ln["month"], 0) + 1
    return max(counts, key=counts.get)


def _last_number(row):
    for cell in reversed(list(row)):
        if isinstance(cell, (int, float)) and not isinstance(cell, bool):
            if cell == cell:  # no NaN
                return float(cell)
    return None


def _looks_like_cost_label(text):
    t = text.upper()
    return any(k in t for k in ("CAJA", "MUESTRA", "NEW CUSTOMER", "DESCUENTO", "INCENTIVO"))


def _find_cost_label(row):
    for cell in row:
        s = clean_str(cell)
        if s and _looks_like_cost_label(s):
            return s
    return None


def parse_incentives(rows):
    free_by_sp = []
    cost_items = []
    total_descuentos = 0.0

    start = None
    for i, row in enumerate(rows):
        if clean_str(row[COL_ORDER]).upper() == "SOLD":
            start = i
            break
    if start is None:
        return empty_incentives()

    header = rows[start]
    free_col = COL_VARIANT  # por defecto C
    for c, cell in enumerate(header):
        if clean_str(cell).upper() == "FREE":
            free_col = c
            break

    for row in rows[start + 1:]:
        label = clean_str(row[COL_DATE]) if len(row) else ""
        last_num = _last_number(row)

        free_val = num(row[free_col]) if len(row) > free_col else 0
        is_sp_row = bool(label) and not _looks_like_cost_label(label)
        if is_sp_row and free_val > 0:
            free_by_sp.append({"name": label, "free": int(free_val)})

        cost_label = _find_cost_label(row)
        if cost_label and last_num:
            cost_items.append({"label": cost_label, "amount": round(last_num, 2)})

        if last_num:
            total_descuentos = round(last_num, 2)

    return {
        "freeUnitsBySalesperson": free_by_sp,
        "costItems": cost_items,
        "totalDescuentos": total_descuentos,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Lectura de fuentes
# ─────────────────────────────────────────────────────────────────────────────
def load_all_sources(folder, pattern, sheet_candidates):
    """Lee todas las fuentes, dedup por orden entre archivos, e incentivos por mes.

    Devuelve (lines, incentives_by_month, source_names, dup_orders).
    """
    seen_orders = set()
    lines = []
    incentives_by_month = {}
    dup_orders = []
    sources = []

    for path in list_source_files(folder):
        try:
            rows = read_sheet_rows(path, sheet_candidates)
        except Exception as e:
            print(f"  ⚠ No se pudo leer {path.name}: {e} (se omite)")
            continue
        sources.append(path.name)
        file_lines = parse_transactions(rows, pattern)
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


# ─────────────────────────────────────────────────────────────────────────────
# Agregaciones (unidades como float, redondeadas a 2 decimales)
# ─────────────────────────────────────────────────────────────────────────────
def aggregate_products(lines):
    prods = {}
    for ln in lines:
        p = prods.setdefault(ln["sku"], {
            "sku": ln["sku"], "name": ln["product"],
            "units": 0.0, "revenue": 0.0, "_orders": set(), "_sp": {},
        })
        p["units"] += ln["qtyOrdered"]
        p["revenue"] += ln["total"]
        p["_orders"].add(ln["order"])
        sp = p["_sp"].setdefault(ln["salesperson"], {"units": 0.0, "revenue": 0.0})
        sp["units"] += ln["qtyOrdered"]
        sp["revenue"] += ln["total"]

    out = []
    for p in prods.values():
        salespeople = [
            {"name": n, "units": round(v["units"], 2), "revenue": round(v["revenue"], 2)}
            for n, v in p["_sp"].items()
        ]
        salespeople.sort(key=lambda s: (-s["units"], -s["revenue"]))
        out.append({
            "sku": p["sku"], "name": p["name"],
            "units": round(p["units"], 2), "revenue": round(p["revenue"], 2),
            "orders": len(p["_orders"]),
            "salespeople": salespeople,
        })
    out.sort(key=lambda p: (-p["units"], -p["revenue"]))
    return out


def aggregate_salespeople(lines):
    sps = {}
    for ln in lines:
        s = sps.setdefault(ln["salesperson"], {
            "name": ln["salesperson"], "units": 0.0, "revenue": 0.0,
            "_orders": set(), "_customers": set(),
        })
        s["units"] += ln["qtyOrdered"]
        s["revenue"] += ln["total"]
        s["_orders"].add(ln["order"])
        s["_customers"].add(ln["customer"])

    out = [{
        "name": s["name"], "units": round(s["units"], 2), "revenue": round(s["revenue"], 2),
        "orders": len(s["_orders"]), "customers": len(s["_customers"]),
    } for s in sps.values()]
    out.sort(key=lambda s: (-s["revenue"], -s["units"]))
    return out


def aggregate_customers(lines):
    cs = {}
    for ln in lines:
        c = cs.setdefault(ln["customer"], {
            "name": ln["customer"], "units": 0.0, "revenue": 0.0, "_orders": set(),
        })
        c["units"] += ln["qtyOrdered"]
        c["revenue"] += ln["total"]
        c["_orders"].add(ln["order"])

    out = [{
        "name": c["name"], "units": round(c["units"], 2), "revenue": round(c["revenue"], 2),
        "orders": len(c["_orders"]),
    } for c in cs.values()]
    out.sort(key=lambda c: (-c["revenue"], -c["units"]))
    return out


def aggregate_daily(lines):
    days = {}
    for ln in lines:
        d = days.setdefault(ln["date"], {
            "date": ln["date"], "units": 0.0, "revenue": 0.0, "_orders": set(),
        })
        d["units"] += ln["qtyOrdered"]
        d["revenue"] += ln["total"]
        d["_orders"].add(ln["order"])

    out = [{
        "date": d["date"], "units": round(d["units"], 2),
        "revenue": round(d["revenue"], 2), "orders": len(d["_orders"]),
    } for d in days.values()]
    out.sort(key=lambda d: d["date"])
    return out


def build_view(lines, incentives):
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
        "units": round(sum(p["units"] for p in products), 2),
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


def build_months(lines, incentives_by_month):
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


def merge_views(views):
    prods = {}
    for v in views:
        for p in v["products"]:
            t = prods.setdefault(p["sku"], {
                "sku": p["sku"], "name": p["name"],
                "units": 0.0, "revenue": 0.0, "orders": 0, "_sp": {}})
            t["units"] += p["units"]; t["revenue"] += p["revenue"]; t["orders"] += p["orders"]
            for sp in p["salespeople"]:
                s = t["_sp"].setdefault(sp["name"], {"units": 0.0, "revenue": 0.0})
                s["units"] += sp["units"]; s["revenue"] += sp["revenue"]
    products = []
    for t in prods.values():
        sps = [{"name": n, "units": round(s["units"], 2), "revenue": round(s["revenue"], 2)}
               for n, s in t["_sp"].items()]
        sps.sort(key=lambda s: (-s["units"], -s["revenue"]))
        products.append({"sku": t["sku"], "name": t["name"], "units": round(t["units"], 2),
                         "revenue": round(t["revenue"], 2), "orders": t["orders"],
                         "salespeople": sps})
    products.sort(key=lambda p: (-p["units"], -p["revenue"]))

    sps = {}
    for v in views:
        for s in v["salespeople"]:
            t = sps.setdefault(s["name"], {"name": s["name"], "units": 0.0, "revenue": 0.0,
                                           "orders": 0, "customers": 0, "freeUnits": 0})
            t["units"] += s["units"]; t["revenue"] += s["revenue"]; t["orders"] += s["orders"]
            t["customers"] += s["customers"]; t["freeUnits"] += s.get("freeUnits", 0)
    salespeople = [{**t, "units": round(t["units"], 2), "revenue": round(t["revenue"], 2)}
                   for t in sps.values()]
    salespeople.sort(key=lambda s: (-s["revenue"], -s["units"]))

    cs = {}
    for v in views:
        for c in v["customers"]:
            t = cs.setdefault(c["name"], {"name": c["name"], "units": 0.0, "revenue": 0.0, "orders": 0})
            t["units"] += c["units"]; t["revenue"] += c["revenue"]; t["orders"] += c["orders"]
    customers = [{**t, "units": round(t["units"], 2), "revenue": round(t["revenue"], 2)}
                 for t in cs.values()]
    customers.sort(key=lambda c: (-c["revenue"], -c["units"]))

    daily = sorted((d for v in views for d in v["daily"]), key=lambda d: d["date"])
    incentives = merge_incentives([v["incentives"] for v in views])

    totals = {
        "units": round(sum(p["units"] for p in products), 2),
        "revenue": round(sum(p["revenue"] for p in products), 2),
        "orders": sum(v["totals"]["orders"] for v in views),
        "customers": len(cs),
        "salespeople": len(sps),
        "freeUnits": sum(f["free"] for f in incentives["freeUnitsBySalesperson"]),
    }
    return {"totals": totals, "products": products, "salespeople": salespeople,
            "customers": customers, "daily": daily, "incentives": incentives}


def load_previous_report(path):
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


# ─────────────────────────────────────────────────────────────────────────────
# Orquestación
# ─────────────────────────────────────────────────────────────────────────────
def build_report(here, brand, distributor, sku_prefix, sheet_candidates):
    here = Path(here)
    sources_dir = here / "fuentes"
    output_file = here / "data.js"
    pattern = sku_pattern(sku_prefix)

    lines, inc_by_month, sources, dup_orders = load_all_sources(
        sources_dir, pattern, sheet_candidates)
    if not lines:
        raise RuntimeError(f"No se detectaron líneas en {sources_dir}. ¿Pusiste los Excel?")

    months = build_months(lines, inc_by_month)
    prev = load_previous_report(output_file)
    months, carried = carry_forward(months, prev)

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
            "brand": brand,
            "distributor": distributor,
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


def build_and_write(here, brand, distributor, sku_prefix, sheet_candidates):
    report = build_report(here, brand, distributor, sku_prefix, sheet_candidates)
    output_file = Path(here) / "data.js"
    js = "const reportData = " + json.dumps(report, indent=2, ensure_ascii=False) + ";\n"
    output_file.write_text(js, encoding="utf-8")

    m, t = report["meta"], report["totals"]
    print("✓ data.js generado")
    print(f"  Marca:      {m['brand']}")
    print(f"  Periodo:    {m['periodLabel']} ({m['periodStart']} → {m['periodEnd']})")
    print(f"  Fuentes:    {', '.join(m['sources']) or '(ninguna)'}")
    print(f"  Meses:      {', '.join(mo['key'] for mo in report['months'])}")
    print(f"  Líneas:     {m['lineItems']}")
    print(f"  Ingresos:   ${t['revenue']:,.2f}   Unidades: {t['units']}   Órdenes: {t['orders']}")
    if m["duplicateOrders"]:
        print(f"  ⚠ Órdenes duplicadas omitidas: {', '.join(m['duplicateOrders'])}")
    if m["carriedForward"]:
        print(f"  ⚠ Meses conservados: {', '.join(m['carriedForward'])}")
    return report
```

- [ ] **Step 4: Correr los tests para verificar que pasan**

Run: `python -m pytest tests/test_report_core.py -v`
Expected: PASS (6 tests).

- [ ] **Step 5: Commit**

```bash
git add report_core.py tests/test_report_core.py
git commit -m "feat: report_core.py — core de datos compartido (cantidades fraccionarias)"
```

---

## Task 2: Wrappers de marca + wiring en split_excel + data.js

**Files:**
- Create: `kombuchacha/build_data.py`
- Create: `robinson-crusoe/build_data.py`
- Modify: `split_excel.py:42-43`
- Modify: `tests/test_main.py:61`

- [ ] **Step 1: Crear el wrapper de Kombuchacha**

Create `kombuchacha/build_data.py`:

```python
"""Genera kombuchacha/data.js desde kombuchacha/fuentes/ usando el core compartido."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # raíz del repo
import report_core

if __name__ == "__main__":
    report_core.build_and_write(
        here=Path(__file__).parent,
        brand="Kombuchacha",
        distributor="LatinFood US Corp",
        sku_prefix="KOM",
        sheet_candidates=("KOM", "Sheet1"),
    )
```

- [ ] **Step 2: Crear el wrapper de Robinson Crusoe**

Create `robinson-crusoe/build_data.py`:

```python
"""Genera robinson-crusoe/data.js desde robinson-crusoe/fuentes/ usando el core compartido."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # raíz del repo
import report_core

if __name__ == "__main__":
    report_core.build_and_write(
        here=Path(__file__).parent,
        brand="Robinson Crusoe",
        distributor="LatinFood US Corp",
        sku_prefix="ROB",
        sheet_candidates=("ROB", "Sheet1"),
    )
```

- [ ] **Step 3: Marcar KOM/ROB como listos en `split_excel.py`**

En `split_excel.py`, reemplazar las líneas 42-43:

```python
    "KOM": Brand("Kombuchacha", "KOM", "kombuchacha", "build_data.py", "KOM", False),
    "ROB": Brand("Robinson Crusoe", "ROB", "robinson-crusoe", "build_data.py", "ROB", False),
```

por:

```python
    "KOM": Brand("Kombuchacha", "KOM", "kombuchacha", "build_data.py", "KOM", True),
    "ROB": Brand("Robinson Crusoe", "ROB", "robinson-crusoe", "build_data.py", "ROB", True),
```

- [ ] **Step 4: Actualizar la aserción rota en `tests/test_main.py`**

KOM ya está listo (`sp1_ready=True`), así que ningún estado imprime "pendiente". En
`tests/test_main.py`, reemplazar la línea 61:

```python
    assert "pendiente" in out  # KOM sigue sin estar listo (sp1_ready=False)
```

por:

```python
    # KOM ahora está listo (sp1_ready=True); con --no-build no se regenera su data.js
    assert (tmp_path / "kombuchacha" / "fuentes" / "KOM-2026-06.xlsx").exists()
    assert "build omitido" in out
```

- [ ] **Step 5: Correr la suite de tests de la raíz (verificar que nada se rompió)**

Run: `python -m pytest tests/ -v`
Expected: PASS (incluyendo `test_report_core.py` y el `test_main.py` actualizado).

- [ ] **Step 6: Generar los `data.js` reales desde las fuentes existentes**

Run (desde la raíz):
```bash
PYTHONUTF8=1 python kombuchacha/build_data.py
PYTHONUTF8=1 python robinson-crusoe/build_data.py
```
Expected: cada uno imprime "✓ data.js generado". KOM: Ingresos $50.64, Unidades 1, Órdenes 1. ROB: Ingresos $60.00, Unidades 1, Órdenes 1.

- [ ] **Step 7: Verificar la forma de los data.js generados**

Run:
```bash
PYTHONUTF8=1 python -c "import re,json; d=json.loads(re.search(r'const reportData = (.*);', open('kombuchacha/data.js',encoding='utf-8').read(), re.S).group(1)); print('KOM', d['meta']['brand'], d['totals'], len(d['products']), 'SKUs', len(d['months']), 'mes')"
PYTHONUTF8=1 python -c "import re,json; d=json.loads(re.search(r'const reportData = (.*);', open('robinson-crusoe/data.js',encoding='utf-8').read(), re.S).group(1)); print('ROB', d['meta']['brand'], d['totals'], len(d['products']), 'SKUs')"
```
Expected: KOM 2 SKUs, units 1.0, revenue 50.64; ROB 1 SKU, revenue 60.0.

- [ ] **Step 8: Commit**

```bash
git add kombuchacha/build_data.py robinson-crusoe/build_data.py split_excel.py tests/test_main.py kombuchacha/data.js robinson-crusoe/data.js
git commit -m "feat: wrappers KOM/ROB + sp1_ready + data.js inicial"
```

---

## Task 3: `kombuchacha/dashboard.html` (clon de GOA con degradación)

**Files:**
- Create: `kombuchacha/dashboard.html` (copiado de `goa-inventory/dashboard.html` y editado)

- [ ] **Step 1: Copiar el dashboard de GOA como base**

Run:
```bash
cp goa-inventory/dashboard.html kombuchacha/dashboard.html
```

- [ ] **Step 2: Cambiar el `<title>`**

Edit `kombuchacha/dashboard.html` — reemplazar:
```html
<title>Garden of the Andes – Reporte de Ventas GOA</title>
```
por:
```html
<title>Kombuchacha – Reporte de Ventas</title>
```

- [ ] **Step 3: Agregar CSS de la grilla de gráficos y la nota de datos escasos**

Edit `kombuchacha/dashboard.html` — reemplazar:
```html
  @media (max-width: 480px) {
    .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    .kpi-card { padding: 14px 14px; }
    .kpi-value { font-size: 1.5rem; }
    .header-title h1 { font-size: 1.15rem; }
  }
</style>
```
por:
```html
  @media (max-width: 480px) {
    .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    .kpi-card { padding: 14px 14px; }
    .kpi-value { font-size: 1.5rem; }
    .header-title h1 { font-size: 1.15rem; }
  }

  /* Grilla de gráficos: auto-fit para que los gráficos ocultos no dejen huecos */
  .charts-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
    gap: 16px; margin-bottom: 16px;
  }
  @media (max-width: 400px) { .charts-grid { grid-template-columns: 1fr; } }

  /* Nota discreta cuando hay pocos datos (marca nueva) */
  .sparse-note {
    background: var(--gold-light); color: #7a5a10;
    border: 1px solid var(--gold); border-radius: 10px;
    padding: 10px 16px; font-size: .82rem; margin-bottom: 20px;
  }
</style>
```

- [ ] **Step 4: Re-marcar el header (icono, nombre)**

Edit `kombuchacha/dashboard.html` — reemplazar:
```html
  <div class="header-brand">
    <div class="leaf-icon"><span>🌿</span></div>
    <div class="header-title">
      <h1>Garden of the Andes</h1>
      <p>Latin Foods</p>
    </div>
  </div>
```
por:
```html
  <div class="header-brand">
    <div class="leaf-icon"><span>🫧</span></div>
    <div class="header-title">
      <h1>Kombuchacha</h1>
      <p>Latin Foods</p>
    </div>
  </div>
```

- [ ] **Step 5: Insertar la nota de datos escasos antes del selector de mes**

Edit `kombuchacha/dashboard.html` — reemplazar:
```html
<main>
  <div class="month-selector" id="month-selector"></div>
```
por:
```html
<main>
  <div class="sparse-note" id="sparse-note">Datos iniciales — el reporte se enriquece a medida que se acumulan meses.</div>
  <div class="month-selector" id="month-selector"></div>
```

- [ ] **Step 6: Reemplazar las dos filas de gráficos por una sola grilla con ids**

Edit `kombuchacha/dashboard.html` — reemplazar:
```html
  <div class="charts-row">
    <div class="chart-card" style="margin:0">
      <div class="section-title">Ventas por SKU — Ingresos</div>
      <div class="chart-wrap rank"><canvas id="revenueChart"></canvas></div>
    </div>
    <div class="chart-card" style="margin:0">
      <div class="section-title">Ranking de vendedores por ingresos</div>
      <div class="chart-wrap rank"><canvas id="spChart"></canvas></div>
    </div>
  </div>

  <div class="charts-row">
    <div class="chart-card" style="margin:0">
      <div class="section-title">Ventas por SKU — Unidades (cajas)</div>
      <div class="chart-wrap rank"><canvas id="unitsChart"></canvas></div>
    </div>
    <div class="chart-card" style="margin:0">
      <div class="section-title">Tendencia de ventas en el periodo</div>
      <div class="chart-wrap rank"><canvas id="trendChart"></canvas></div>
    </div>
  </div>
```
por:
```html
  <div class="charts-grid" id="charts">
    <div class="chart-card" id="card-revenue" style="margin:0">
      <div class="section-title">Ventas por SKU — Ingresos</div>
      <div class="chart-wrap rank"><canvas id="revenueChart"></canvas></div>
    </div>
    <div class="chart-card" id="card-units" style="margin:0">
      <div class="section-title">Ventas por SKU — Unidades (cajas)</div>
      <div class="chart-wrap rank"><canvas id="unitsChart"></canvas></div>
    </div>
    <div class="chart-card" id="card-sp" style="margin:0">
      <div class="section-title">Ranking de vendedores por ingresos</div>
      <div class="chart-wrap rank"><canvas id="spChart"></canvas></div>
    </div>
    <div class="chart-card" id="card-trend" style="margin:0">
      <div class="section-title">Tendencia de ventas en el periodo</div>
      <div class="chart-wrap rank"><canvas id="trendChart"></canvas></div>
    </div>
  </div>
```

- [ ] **Step 7: Pluralizar el sub de "Unidades vendidas" (1 SKU vs N SKUs)**

Edit `kombuchacha/dashboard.html` — reemplazar:
```javascript
    { label: 'Unidades vendidas', value: qty(t.units), sub: `${qty(view.products.length)} SKUs` },
```
por:
```javascript
    { label: 'Unidades vendidas', value: qty(t.units), sub: `${qty(view.products.length)} SKU${view.products.length === 1 ? '' : 's'}` },
```

- [ ] **Step 8: Adaptar `shortName` al nombre de producto de KOM**

Edit `kombuchacha/dashboard.html` — reemplazar:
```javascript
const shortName = name => name.replace(' Herbal Tea', '').replace('Pure ', '');
```
por:
```javascript
const shortName = name => name.replace('Kombuchacha ', '');
```

- [ ] **Step 9: Añadir la lógica de degradación a los gráficos en `render()`**

Edit `kombuchacha/dashboard.html` — reemplazar el bloque desde `// Destruir gráficos previos` hasta el cierre del gráfico de tendencia:
```javascript
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
```
por:
```javascript
  // Destruir gráficos previos
  Object.values(charts).forEach(c => c.destroy());
  charts = {};

  // ── Degradación: un gráfico solo se dibuja si compara ≥2 elementos ──
  const show = {
    revenue: view.products.length >= 2,
    units:   view.products.length >= 2,
    sp:      view.salespeople.length >= 2,
    trend:   view.daily.length >= 2,
  };
  document.getElementById('card-revenue').style.display = show.revenue ? '' : 'none';
  document.getElementById('card-units').style.display   = show.units   ? '' : 'none';
  document.getElementById('card-sp').style.display      = show.sp      ? '' : 'none';
  document.getElementById('card-trend').style.display   = show.trend   ? '' : 'none';
  document.getElementById('charts').style.display =
    (show.revenue || show.units || show.sp || show.trend) ? '' : 'none';

  // Ventas por SKU — Ingresos
  if (show.revenue) horizontalBarChart('revenueChart',
    view.products.map(p => ({ label: shortName(p.name), value: p.revenue })),
    money0, ctx => ` ${money(ctx.parsed.x)}`);

  // Ventas por SKU — Unidades
  if (show.units) horizontalBarChart('unitsChart',
    view.products.map(p => ({ label: shortName(p.name), value: p.units })),
    qty, ctx => ` ${qty(ctx.parsed.x)} unidades`);

  // Ranking de vendedores por ingresos
  if (show.sp) horizontalBarChart('spChart',
    view.salespeople.map(s => ({ label: firstName(s.name), value: s.revenue, meta: s })),
    money0,
    ctx => ` ${money(ctx.parsed.x)} (${qty(Math.round(ctx.parsed.x / (t.revenue || 1) * 100))}%)`,
    s => `${qty(s.units)} u. · ${qty(s.orders)} órdenes · ${qty(s.customers)} clientes`);

  // Tendencia
  if (show.trend) charts.trendChart = new Chart(document.getElementById('trendChart'), {
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
```

- [ ] **Step 10: Dar id al bloque de incentivos y ocultarlo cuando está vacío**

Edit `kombuchacha/dashboard.html` — (a) reemplazar:
```html
  <!-- Incentivos -->
  <div class="two-col">
```
por:
```html
  <!-- Incentivos -->
  <div class="two-col" id="incentives-block">
```

(b) reemplazar:
```javascript
  document.getElementById('cost-total').textContent = money(view.incentives.totalDescuentos);
}
```
por:
```javascript
  document.getElementById('cost-total').textContent = money(view.incentives.totalDescuentos);

  // Degradación: ocultar el bloque de incentivos si no hay nada que mostrar
  const hasInc = freeRows.length || view.incentives.costItems.length || view.incentives.totalDescuentos;
  document.getElementById('incentives-block').style.display = hasInc ? '' : 'none';
}
```

- [ ] **Step 11: Ocultar selector de mes y togglear la nota en la inicialización**

Edit `kombuchacha/dashboard.html` — reemplazar:
```javascript
// Render inicial: vista general
render(VIEWS[0].view, VIEWS[0].header);
```
por:
```javascript
// Degradación global (independiente de la vista):
//  nota de datos iniciales y selector de mes solo cuando corresponde
document.getElementById('sparse-note').style.display = d.months.length < 2 ? '' : 'none';
if (d.months.length < 2) document.getElementById('month-selector').style.display = 'none';

// Render inicial: vista general
render(VIEWS[0].view, VIEWS[0].header);
```

- [ ] **Step 12: Verificar el dashboard en el navegador (preview)**

Añadir a `.claude/launch.json` (si aún no está) una entrada temporal o servir manualmente:
```bash
python -m http.server 8138 --directory kombuchacha
```
Abrir `http://localhost:8138/dashboard.html` y verificar:
- KPIs visibles (Ingresos $51, Unidades 1, Clientes 1, Vendedor top Mireya Fernandez).
- Se dibuja "Ventas por SKU — Ingresos" y "— Unidades" (2 barras cada uno).
- NO se dibujan "Ranking de vendedores" ni "Tendencia" (1 vendedor, 1 día).
- Tablas de producto (2 filas + total), vendedores (1 fila) y clientes (1 fila) presentes.
- Bloque de incentivos oculto; nota "Datos iniciales…" visible; sin selector de mes.
- Sin errores en la consola del navegador.

Detener el server tras verificar.

- [ ] **Step 13: Commit**

```bash
git add kombuchacha/dashboard.html
git commit -m "feat: dashboard Kombuchacha (clon de GOA con degradación)"
```

---

## Task 4: `robinson-crusoe/dashboard.html` (copia re-marcada)

**Files:**
- Create: `robinson-crusoe/dashboard.html` (copia de `kombuchacha/dashboard.html` + 4 cambios de marca)

- [ ] **Step 1: Copiar el dashboard de Kombuchacha**

Run:
```bash
cp kombuchacha/dashboard.html robinson-crusoe/dashboard.html
```

- [ ] **Step 2: Cambiar el `<title>`**

Edit `robinson-crusoe/dashboard.html` — reemplazar:
```html
<title>Kombuchacha – Reporte de Ventas</title>
```
por:
```html
<title>Robinson Crusoe – Reporte de Ventas</title>
```

- [ ] **Step 3: Re-marcar el header (icono, nombre)**

Edit `robinson-crusoe/dashboard.html` — reemplazar:
```html
    <div class="leaf-icon"><span>🫧</span></div>
    <div class="header-title">
      <h1>Kombuchacha</h1>
      <p>Latin Foods</p>
    </div>
```
por:
```html
    <div class="leaf-icon"><span>🦪</span></div>
    <div class="header-title">
      <h1>Robinson Crusoe</h1>
      <p>Latin Foods</p>
    </div>
```

- [ ] **Step 4: Neutralizar `shortName` (los nombres de ROB no llevan prefijo de marca)**

Edit `robinson-crusoe/dashboard.html` — reemplazar:
```javascript
const shortName = name => name.replace('Kombuchacha ', '');
```
por:
```javascript
const shortName = name => name;
```

- [ ] **Step 5: Verificar el dashboard en el navegador (preview)**

Run:
```bash
python -m http.server 8139 --directory robinson-crusoe
```
Abrir `http://localhost:8139/dashboard.html` y verificar:
- KPIs visibles (Ingresos $60, Unidades 1, Clientes 1, Vendedor top Luis Soler).
- NINGÚN gráfico dibujado (1 SKU, 1 vendedor, 1 día) → grilla de gráficos oculta.
- Tablas de producto (1 fila + total), vendedores (1 fila) y clientes (1 fila) presentes.
- Bloque de incentivos oculto; nota "Datos iniciales…" visible; sin selector de mes.
- Header dice "Robinson Crusoe"; sin errores en consola.

Detener el server tras verificar.

- [ ] **Step 6: Commit**

```bash
git add robinson-crusoe/dashboard.html
git commit -m "feat: dashboard Robinson Crusoe (copia re-marcada de KOM)"
```

---

## Task 5: Preview configs + verificación final

**Files:**
- Modify: `.claude/launch.json`

- [ ] **Step 1: Añadir servidores de preview para KOM y ROB**

Edit `.claude/launch.json` — reemplazar el array `configurations` para incluir las dos entradas nuevas (puertos 8138 y 8139):

```json
{
  "version": "0.0.1",
  "configurations": [
    {
      "name": "sanjose-dashboard",
      "runtimeExecutable": "python",
      "runtimeArgs": ["-m", "http.server", "8137", "--directory", "reporte-sanjose"],
      "port": 8137
    },
    {
      "name": "kombuchacha-dashboard",
      "runtimeExecutable": "python",
      "runtimeArgs": ["-m", "http.server", "8138", "--directory", "kombuchacha"],
      "port": 8138
    },
    {
      "name": "robinson-dashboard",
      "runtimeExecutable": "python",
      "runtimeArgs": ["-m", "http.server", "8139", "--directory", "robinson-crusoe"],
      "port": 8139
    }
  ]
}
```

- [ ] **Step 2: Correr la suite completa de tests**

Run: `python -m pytest -q`
Expected: todo PASS (raíz, goa-inventory, reporte-sanjose). Ninguna regresión.

- [ ] **Step 3: Verificación de integración del dispatcher (opcional pero recomendado)**

Confirmar que `split_excel.py` reconoce KOM/ROB como listos revisando la config:
```bash
PYTHONUTF8=1 python -c "import split_excel as s; print({k:(b.name,b.sp1_ready) for k,b in s.BRANDS.items()})"
```
Expected: KOM y ROB con `sp1_ready=True`.

- [ ] **Step 4: Commit**

```bash
git add .claude/launch.json
git commit -m "chore: launch configs de preview para KOM y ROB"
```

---

## Cierre

Tras completar todas las tareas:
- Anunciar: "I'm using the finishing-a-development-branch skill to complete this work."
- **REQUIRED SUB-SKILL:** Use superpowers:finishing-a-development-branch para verificar tests, presentar opciones y ejecutar la elección (merge a `main` → despliega en GitHub Pages).

Nota: los `.xlsx` de `fuentes/` están gitignoreados (datos comerciales); solo se commitea el `data.js` generado, igual que en GOA. Recordar la advertencia de privacidad: los dashboards quedan públicos en GitHub Pages hasta que se haga SP4 (subdominio privado).
```
