"""
parse_altagama.py — Extractor del reporte de sell-through de Alta Gama (GOA)

Alta Gama entrega un reporte MUCHO más limitado que Latin Foods: una tabla con
una fila por producto y una columna por periodo, con las UNIDADES VENDIDAS de
cada periodo. No trae precios, clientes, órdenes, inventario ni fechas de
transacción, así que el análisis se limita al sell-through EN UNIDADES.

Estructura del Excel (hoja Sheet1; la posición de la tabla cambia entre envíos):
  - Una fila de encabezado con un periodo por columna. Los meses completos
    vienen como fecha (Excel interpreta "Jan 26" → 2026-01-26: el 26 es el AÑO)
    o como texto ("Jul 26"), y el periodo parcial como texto libre
    ("Jul 1 - 25, 26"). Puede haber una columna `TOTAL` por fila (se usa para
    validar) y filas sin cifras como `Qty` o `Products` (se ignoran).
  - Una fila por producto, con el formato `"<SKU> (<Nombre>)"`. Solo se
    importan los SKU `GA …`: desde jul–oct 2026 Alta Gama mezcla en el mismo
    archivo Kombuchacha y El Charrua.
  - Fila de total (`TOTAL` o `Total Inventory`, que pese al nombre suma
    ventas): NO se importa como producto; solo valida el archivo.

Varias fuentes: el script combina todos los Excel de la carpeta con el
histórico `altagama-historico.json` (ene–jul 2026, reconstruido desde lo ya
publicado). Cada mes lo aporta el archivo con el corte más reciente que lo
traiga; ver `merge_sources`.

Este script escribe `data-altagama.js` (objeto `altaGamaData`). Es
independiente de `parse_excel.py` / `data.js` (Latin Foods): ni lee ni escribe
nada de esa fuente.

Uso:
    python parse_altagama.py
"""

import re
import json
import calendar
import datetime
from pathlib import Path

import openpyxl

# ─────────────────────────────────────────────────────────────────────────────
# Configuración
# ─────────────────────────────────────────────────────────────────────────────
HERE = Path(__file__).parent
SOURCE_DIR = HERE / "Inventarios-GOA-altagama"
SOURCE_GLOB = "*.xlsx"           # Alta Gama cambia el nombre en cada envío
HISTORY_FILE = HERE / "altagama-historico.json"
OUTPUT_FILE = HERE / "data-altagama.js"

BRAND = "Garden of the Andes"
DISTRIBUTOR = "Alta Gama"
METRIC = "Unidades vendidas (sell-through en unidades)"

SKU_PREFIX = "GA "             # el resto de las filas son de otras marcas
TOTAL_LABEL = "TOTAL"          # fila/columna de control, no es un producto
TOTAL_LABELS = {"TOTAL", "TOTAL INVENTORY"}   # "Total Inventory" suma ventas, no stock
SKIP_LABELS = {"PRODUCTS"}     # rótulo de sección, sin cifras

MESES_ES = [
    "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]
MESES_ES_ABBR = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                 "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

MONTHS_EN = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# "Jul 1 - 25, 26" / "Jul 1-25, 2026"  → periodo parcial
PARTIAL_RE = re.compile(
    r"^\s*([A-Za-z]{3,})\.?\s*(\d{1,2})\s*[-–—]\s*(\d{1,2})\s*,?\s*(\d{2,4})\s*$")
# "Jan 26" / "Jan-26" / "Jan 2026"     → mes completo escrito como texto
FULL_TEXT_RE = re.compile(r"^\s*([A-Za-z]{3,})\.?\s*[-\s]\s*(\d{2,4})\s*$")
# "GA Green Tea 1.4oz (Garden of the Andes - Green Tea 1.4oz)"
PRODUCT_RE = re.compile(r"^\s*(.+?)\s*\(\s*(.+?)\s*\)\s*$")


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de limpieza
# ─────────────────────────────────────────────────────────────────────────────
def clean_str(value):
    """Normaliza un string: trim y colapsa espacios internos. None -> ''."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def units(value):
    """Celda de unidades → (int, None), o (None, motivo) si NO está disponible.

    Semántica confirmada con el negocio (2026-08-18): Alta Gama escribe `0`
    cuando **no hubo movimiento** en el periodo, así que un 0 es un dato real y
    se conserva como 0. Una celda vacía sí es dato ausente: se devuelve `None`
    y nunca se sustituye por cero."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None, "vacía"
    if isinstance(value, bool):
        return None, "booleana"
    if isinstance(value, (int, float)):
        if value != value:                      # NaN
            return None, "NaN"
        return int(round(value)), None
    text = clean_str(value).replace(",", "")
    try:
        return int(round(float(text))), None
    except ValueError:
        return None, "no numérica (%r)" % clean_str(value)


def four_digit_year(raw):
    """26 -> 2026 · 2026 -> 2026."""
    year = int(raw)
    return 2000 + year if year < 100 else year


def full_month_period(year, month):
    last = calendar.monthrange(year, month)[1]
    return {
        "key": "%d-%02d" % (year, month),
        "label": "%s %d" % (MESES_ES[month].capitalize(), year),
        "short": MESES_ES_ABBR[month],
        "year": year,
        "month": month,
        "start": "%d-%02d-01" % (year, month),
        "end": "%d-%02d-%02d" % (year, month, last),
        "partial": False,
        "coverageDays": last,
        "monthDays": last,
    }


def partial_period(year, month, day_from, day_to):
    """Periodo que no cubre el mes entero.

    Decisión comercial (2026-08-18): el dashboard lo presenta como un mes más y
    avisa una sola vez al principio; el corte real vive aquí (`partial`,
    `coverageDays`) para el reporte interno y el correo al distribuidor."""
    last = calendar.monthrange(year, month)[1]
    return {
        "key": "%d-%02d" % (year, month),
        "label": "%s %d" % (MESES_ES[month].capitalize(), year),
        "short": MESES_ES_ABBR[month],
        "year": year,
        "month": month,
        "start": "%d-%02d-%02d" % (year, month, day_from),
        "end": "%d-%02d-%02d" % (year, month, day_to),
        "partial": True,
        "coverageDays": day_to - day_from + 1,
        "monthDays": last,
    }


def parse_period_header(value):
    """Encabezado de columna → periodo. Lanza ValueError si no se reconoce.

    Excel guarda "Jan 26" como fecha 2026-01-26 (el 26 es el AÑO, no el día):
    de ahí se toman año y mes, y el periodo se expande al mes completo.
    """
    if isinstance(value, (datetime.datetime, datetime.date)):
        return full_month_period(value.year, value.month)

    text = clean_str(value)
    if not text:
        raise ValueError("encabezado de periodo vacío")

    m = PARTIAL_RE.match(text)
    if m:
        name, day_from, day_to, year = m.groups()
        month = MONTHS_EN.get(name[:3].lower())
        if month is None:
            raise ValueError("mes no reconocido en %r" % text)
        return partial_period(four_digit_year(year), month, int(day_from), int(day_to))

    m = FULL_TEXT_RE.match(text)
    if m:
        name, year = m.groups()
        month = MONTHS_EN.get(name[:3].lower())
        if month is None:
            raise ValueError("mes no reconocido en %r" % text)
        return full_month_period(four_digit_year(year), month)

    raise ValueError("encabezado de periodo no reconocido: %r" % text)


def split_product(value):
    """'GA Green Tea 1.4oz (Garden of the Andes - Green Tea 1.4oz)'
       -> ('GA Green Tea 1.4oz', 'Garden of the Andes - Green Tea 1.4oz').

    Si el texto no trae el paréntesis con el nombre largo, el SKU queda en None
    y el texto completo se usa como nombre: no se inventa un SKU."""
    text = clean_str(value)
    m = PRODUCT_RE.match(text)
    if not m:
        return None, text
    return clean_str(m.group(1)), clean_str(m.group(2))


def short_name(name, sku):
    """Etiqueta corta para gráficos: 'Garden of the Andes - Green Tea 1.4oz'
       -> 'Green Tea 1.4oz'."""
    base = name.split(" - ", 1)[1] if " - " in name else name
    base = clean_str(base)
    return base or clean_str(sku) or name


def pct(part, whole, digits=1):
    return round(part / whole * 100, digits) if whole else None


# ─────────────────────────────────────────────────────────────────────────────
# Lectura
# ─────────────────────────────────────────────────────────────────────────────
def find_sources(folder=SOURCE_DIR, pattern=SOURCE_GLOB):
    """Todos los Excel de la carpeta, por nombre; ignora temporales (~$)."""
    folder = Path(folder)
    if not folder.exists():
        return []
    return sorted(p for p in folder.glob(pattern) if not p.name.startswith("~$"))


def read_sheet_rows(path):
    """Filas (values_only) de la hoja de datos."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
    return list(ws.iter_rows(values_only=True))


def locate_header(rows):
    """Primera fila con al menos un encabezado de periodo → (índice, [(col, periodo)]).

    Alta Gama mueve la tabla de un envío a otro (la de 2026 YTD empezaba en A1;
    la de jul–oct empieza en la columna E con una fila `Qty` debajo), así que no
    se asume ninguna posición fija."""
    for idx, row in enumerate(rows):
        found = []
        for col, cell in enumerate(row or ()):
            if cell is None or clean_str(cell) == "":
                continue
            try:
                found.append((col, parse_period_header(cell)))
            except ValueError:
                pass
        if found:
            return idx, found
    raise RuntimeError("No se encontró ninguna fila con encabezados de periodo.")


def row_label(row, first_period_col):
    """Etiqueta de la fila: la última celda con texto a la izquierda de los periodos."""
    for cell in reversed(row[:first_period_col]):
        text = clean_str(cell)
        if text:
            return text
    return ""


def parse_sheet(rows):
    """Filas crudas → fuente con periodos, productos GOA, filas ajenas y controles.

    La fila de total se devuelve aparte (`totalRow`): nunca entra como producto.
    Las filas que no son de Garden of the Andes (Alta Gama mezcla Kombuchacha y
    El Charrua en el mismo archivo) van a `excluded` y no se importan.
    """
    if not rows:
        raise RuntimeError("La hoja está vacía.")

    header_idx, period_cols = locate_header(rows)
    header = rows[header_idx]
    periods = [p for _, p in period_cols]
    columns = [c for c, _ in period_cols]
    first_col = columns[0]
    total_col = next((c for c, cell in enumerate(header)
                      if c > first_col and clean_str(cell).upper() == TOTAL_LABEL), None)

    keys = [p["key"] for p in periods]
    if len(set(keys)) != len(keys):
        raise RuntimeError("Periodos duplicados en el encabezado: %s" % keys)

    products, excluded, total_row, issues = [], [], None, []
    all_rows = {k: 0 for k in keys}        # todas las filas, para cuadrar con el total
    for row_no, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        row = list(row or ())
        raw = row_label(row, first_col)
        if not raw or raw.upper() in SKIP_LABELS:
            continue

        values, reasons = {}, []
        for key, col in zip(keys, columns):
            value, reason = units(row[col] if col < len(row) else None)
            values[key] = value
            if reason:
                reasons.append("fila %d (%s) · %s: celda %s" % (row_no, raw, key, reason))

        if raw.upper() in TOTAL_LABELS:
            total_row = values
            continue

        for key in keys:
            all_rows[key] += values[key] or 0

        sku, name = split_product(raw)
        if not (sku and sku.startswith(SKU_PREFIX)):
            excluded.append({"label": raw, "units": sum_units(values.values())})
            continue

        issues.extend(reasons)
        if total_col is not None:
            expected, _ = units(row[total_col] if total_col < len(row) else None)
            if expected is not None and expected != sum_units(values.values()):
                issues.append("fila %d (%s): la columna TOTAL dice %d y los periodos suman %d"
                              % (row_no, raw, expected, sum_units(values.values())))

        products.append({
            "sku": sku,
            "name": name,
            "shortName": short_name(name, sku),
            "byPeriod": values,
        })

    if not products:
        raise RuntimeError("No se encontró ninguna fila de producto de %s." % BRAND)

    skus = [p["sku"] for p in products]
    dup = sorted({s for s in skus if skus.count(s) > 1})
    if dup:
        raise RuntimeError("SKU duplicados en el Excel: %s" % dup)

    names = [p["name"].lower() for p in products]
    dup_names = sorted({n for n in names if names.count(n) > 1})
    if dup_names:
        raise RuntimeError("Productos duplicados por nombre: %s" % dup_names)

    return {
        "periods": periods,
        "products": products,
        "excluded": excluded,
        "totalRow": total_row,
        "allRows": all_rows,
        "issues": issues,
    }


def load_excel(path):
    source = parse_sheet(read_sheet_rows(path))
    source["name"] = Path(path).name
    return source


def period_from_range(start, end):
    """'2026-07-01', '2026-07-25' → periodo completo o parcial."""
    y, m, d_from = (int(x) for x in start.split("-"))
    d_to = int(end.split("-")[2])
    if d_from == 1 and d_to == calendar.monthrange(y, m)[1]:
        return full_month_period(y, m)
    return partial_period(y, m, d_from, d_to)


def load_history(path=HISTORY_FILE):
    """Histórico ya publicado (JSON del repo) → fuente con la misma forma que un Excel.

    Existe porque el Excel 2026 YTD original se perdió: ene–jul se reconstruyeron
    desde el `data-altagama.js` publicado. Siempre es la fuente de menor prioridad."""
    path = Path(path)
    if not path.exists():
        return None
    raw = json.loads(path.read_text(encoding="utf-8"))
    periods = [period_from_range(p["start"], p["end"]) for p in raw["periods"]]
    keys = [p["key"] for p in periods]
    products = [{"sku": p["sku"], "name": p["name"],
                 "shortName": short_name(p["name"], p["sku"]),
                 "byPeriod": {k: p["byPeriod"].get(k) for k in keys}}
                for p in raw["products"]]
    return {
        "name": raw["source"] + " (histórico)",
        "periods": periods,
        "products": products,
        "excluded": [],
        "totalRow": raw.get("totalRow"),
        "allRows": {k: sum_units(p["byPeriod"][k] for p in products) for k in keys},
        "issues": [],
    }


# ─────────────────────────────────────────────────────────────────────────────
# Agregación
# ─────────────────────────────────────────────────────────────────────────────
def sum_units(values):
    """Suma ignorando los None (dato no disponible ≠ cero)."""
    return sum(v for v in values if v is not None)


def cutoff(source):
    return max(p["end"] for p in source["periods"])


def merge_sources(sources):
    """Combina varias fuentes por periodo: cada mes lo aporta la fuente con el
    corte más reciente que lo traiga (el histórico siempre pierde).

    Así un mes parcial (jul 1–25) queda reemplazado por el mes completo cuando
    llega, y un reporte de Alta Gama que solo trae los últimos meses no borra
    los anteriores. Cada reemplazo queda registrado en `overrides`; si un
    producto BAJA al reemplazar un parcial, o un mes completo cambia de cifra,
    se anota como incidencia."""
    ordered = sorted(sources, key=lambda s: (not s["name"].endswith("(histórico)"), cutoff(s)))
    winner, overrides, issues = {}, [], []
    for src in ordered:
        issues.extend("%s · %s" % (src["name"], i) for i in src["issues"])
        for period in src["periods"]:
            key = period["key"]
            prev = winner.get(key)
            winner[key] = (src, period)
            if prev is None:
                continue
            old_src, old_period = prev
            old_vals = {p["sku"]: p["byPeriod"][key] for p in old_src["products"]}
            new_vals = {p["sku"]: p["byPeriod"][key] for p in src["products"]}
            overrides.append({
                "key": key, "from": old_src["name"], "to": src["name"],
                "before": sum_units(old_vals.values()),
                "after": sum_units(new_vals.values()),
                "partialBefore": old_period["partial"], "partialAfter": period["partial"],
            })
            for sku, old in old_vals.items():
                new = new_vals.get(sku)
                if old is None or new is None or old == new:
                    continue
                if old_period["partial"] and old_period["end"] < period["end"]:
                    if new < old:
                        issues.append("%s · %s: %s bajó de %d a %d al completar el mes"
                                      % (src["name"], key, sku, old, new))
                else:
                    issues.append("%s · %s: %s reexpresado de %d a %d"
                                  % (src["name"], key, sku, old, new))

    keys = sorted(winner)
    periods = [dict(winner[k][1]) for k in keys]

    # Productos: unión de los SKU de las fuentes ganadoras. El nombre sale de la
    # fuente más reciente que lo menciona.
    catalog = {}
    for k in keys:
        for p in winner[k][0]["products"]:
            catalog[p["sku"]] = {"sku": p["sku"], "name": p["name"], "shortName": p["shortName"]}
    products = []
    for sku, base in catalog.items():
        by_period = {}
        for k in keys:
            src = winner[k][0]
            row = next((p for p in src["products"] if p["sku"] == sku), None)
            by_period[k] = None if row is None else row["byPeriod"][k]
            if row is None:
                issues.append("%s · %s: %s no aparece en el archivo (queda n/d)"
                              % (src["name"], k, sku))
        products.append(dict(base, byPeriod=by_period))

    used = []
    for src in ordered:
        mine = [k for k in keys if winner[k][0] is src]
        if mine:
            used.append({"file": src["name"], "periods": mine})

    latest = ordered[-1]
    return {
        "periods": periods,
        "products": products,
        "winner": {k: winner[k][0] for k in keys},
        "sources": used,
        "latest": latest["name"],
        "excluded": [dict(e, file=latest["name"]) for e in latest["excluded"]],
        "overrides": overrides,
        "issues": issues,
    }


def build_report(sources):
    merged = merge_sources(sources)
    periods, products, issues = merged["periods"], merged["products"], merged["issues"]

    keys = [p["key"] for p in periods]
    full_keys = [p["key"] for p in periods if not p["partial"]]
    partial_keys = [p["key"] for p in periods if p["partial"]]

    # Totales por periodo, recalculados SIEMPRE desde las filas de producto
    for period in periods:
        period["units"] = sum_units(p["byPeriod"][period["key"]] for p in products)

    # Validación contra la fila de total del archivo que aporta cada periodo. Se
    # compara con TODAS sus filas (también las ajenas a GOA, que el total suma);
    # si el archivo no trae fila de total, el periodo queda sin control (ok=None).
    checks, all_match = [], True
    for period in periods:
        src = merged["winner"][period["key"]]
        expected = (src["totalRow"] or {}).get(period["key"])
        computed = src["allRows"][period["key"]]
        ok = None if expected is None else expected == computed
        all_match = all_match and ok is not False
        checks.append({"key": period["key"], "label": period["label"], "source": src["name"],
                       "excelTotal": expected, "computed": computed, "ok": ok,
                       "mixedBrands": bool(src["excluded"])})

    grand_total = sum(p["units"] for p in periods)
    total_full = sum(p["units"] for p in periods if not p["partial"])

    # Variación mes a mes, entre periodos consecutivos. El periodo parcial entra
    # como un mes más: los días que le faltan no mueven la aguja lo suficiente
    # como para justificar dejarlo fuera de la lectura comercial.
    prev = None
    for period in periods:
        period["delta"] = None if prev is None else period["units"] - prev
        period["deltaPct"] = (None if prev in (None, 0)
                              else pct(period["units"] - prev, prev))
        prev = period["units"]

    # Productos
    for prod in products:
        prod["units"] = sum_units(prod["byPeriod"].values())
        prod["unitsFullMonths"] = sum_units(prod["byPeriod"][k] for k in full_keys)
        prod["share"] = pct(prod["units"], grand_total)
        prod["avgFullMonth"] = (round(prod["unitsFullMonths"] / len(full_keys), 1)
                                if full_keys else None)
        prod["missingPeriods"] = [k for k in keys if prod["byPeriod"][k] is None]
        best_key = max((k for k in keys if prod["byPeriod"][k] is not None),
                       key=lambda k: prod["byPeriod"][k], default=None)
        best_period = next((p for p in periods if p["key"] == best_key), None)
        prod["best"] = None if best_period is None else {
            "key": best_key, "label": best_period["label"],
            "units": prod["byPeriod"][best_key]}

    products.sort(key=lambda p: (-p["units"], p["shortName"]))

    top_product = products[0] if products else None
    best_month = max(periods, key=lambda p: p["units"], default=None)
    last_period = periods[-1] if periods else None
    partial = next((p for p in periods if p["partial"]), None)

    period_start = min(p["start"] for p in periods)
    period_end = max(p["end"] for p in periods)
    period_label = ("%s – %s" % (periods[0]["label"], periods[-1]["label"])
                    if len(periods) > 1 else periods[0]["label"])

    if partial:
        # `note` es la frase larga para el correo/reporte interno; `short` es la
        # única mención que aparece en el dashboard comercial.
        partial_note = {
            "key": partial["key"],
            "short": ("%s incluye ventas hasta el %d de %s. Los demás meses están completos."
                      % (partial["label"], int(partial["end"][-2:]),
                         MESES_ES[partial["month"]])),
            "note": ("Los datos de %s corresponden al periodo comprendido entre "
                     "el %d y el %d de %s de %d."
                     % (partial["label"].split(" ")[0].lower(),
                        int(partial["start"][-2:]), int(partial["end"][-2:]),
                        MESES_ES[partial["month"]], partial["year"])),
        }
    else:
        partial_note = None
    cutoff_note = partial_note["note"] if partial_note else \
        "Todos los periodos del reporte son meses completos."

    return {
        "meta": {
            "brand": BRAND,
            "distributor": DISTRIBUTOR,
            "metric": METRIC,
            "source": merged["latest"],
            "sources": merged["sources"],
            "periodStart": period_start,
            "periodEnd": period_end,
            "periodLabel": period_label,
            "cutoffDate": period_end,
            "partialPeriods": partial_keys,
            "partialNote": partial_note,
            "notes": [
                "Alta Gama proporciona información de sell-through en unidades. "
                "El reporte no incluye precios ni valores monetarios.",
                cutoff_note,
            ],
            "validation": {
                "totalRowPresent": all(c["excelTotal"] is not None for c in checks),
                "allMatch": all_match,
                "checks": checks,
                "overrides": merged["overrides"],
                "issues": issues,
            },
        },
        "totals": {
            "units": grand_total,
            "unitsFullMonths": total_full,
            "products": len(products),
            "periods": len(periods),
            "fullMonths": len(full_keys),
            "partialPeriods": len(partial_keys),
            # Promedio sobre todos los periodos (así se lee el dashboard) y
            # sobre los meses completos (para el análisis interno).
            "avgPeriod": round(grand_total / len(periods), 1) if periods else None,
            "avgFullMonth": round(total_full / len(full_keys), 1) if full_keys else None,
        },
        "highlights": {
            "topProduct": None if not top_product else {
                "sku": top_product["sku"], "shortName": top_product["shortName"],
                "units": top_product["units"], "share": top_product["share"]},
            "bestMonth": None if not best_month else {
                "key": best_month["key"], "label": best_month["label"],
                "units": best_month["units"]},
            "lastMonth": None if not last_period else {
                "key": last_period["key"], "label": last_period["label"],
                "units": last_period["units"], "delta": last_period["delta"],
                "deltaPct": last_period["deltaPct"]},
            "partial": None if not partial else {
                "key": partial["key"], "label": partial["label"],
                "short": partial["short"], "units": partial["units"],
                "start": partial["start"], "end": partial["end"],
                "coverageDays": partial["coverageDays"],
                "monthDays": partial["monthDays"]},
        },
        "periods": periods,
        "products": products,
    }


def main():
    paths = find_sources()
    history = load_history()
    if not paths and history is None:
        raise SystemExit("✗ No se encontró ningún '%s' en %s ni el histórico %s."
                         % (SOURCE_GLOB, SOURCE_DIR, HISTORY_FILE.name))

    sources = ([history] if history else []) + [load_excel(p) for p in paths]
    report = build_report(sources)
    excluded = merge_sources(sources)["excluded"]

    m, t, v = report["meta"], report["totals"], report["meta"]["validation"]
    print("✓ data-altagama.js generado")
    print("  Distribuidor: %s" % m["distributor"])
    print("  Fuentes:")
    for s in m["sources"]:
        print("    · %s → %s" % (s["file"], ", ".join(s["periods"])))
    print("  Periodo:      %s (%s → %s)" % (m["periodLabel"], m["periodStart"], m["periodEnd"]))
    print("  Productos:    %d   Periodos: %d (%d completos + %d parcial)"
          % (t["products"], t["periods"], t["fullMonths"], t["partialPeriods"]))
    print("  Unidades:     %d total · %d en meses completos"
          % (t["units"], t["unitsFullMonths"]))
    for c in v["checks"]:
        mark = {True: "✓", False: "✗", None: "·"}[c["ok"]]
        print("    %s %s: GOA %s · archivo %s (fila total %s)"
              % (mark, c["key"], next(p["units"] for p in report["periods"] if p["key"] == c["key"]),
                 c["computed"], "sin fila" if c["excelTotal"] is None else c["excelTotal"]))
    if not v["allMatch"]:
        print("  ⚠ Algún total calculado NO cuadra con la fila de total del Excel.")
    for o in v["overrides"]:
        print("  ↻ %s: %s (%d u.) reemplazado por %s (%d u.)"
              % (o["key"], o["from"], o["before"], o["to"], o["after"]))
    if excluded:
        # Solo consola: data-altagama.js es público y lo ve GOA; las ventas de
        # otras marcas no se publican ahí.
        print("  Filas ajenas a %s (no se importan ni se publican):" % BRAND)
        for e in excluded:
            print("    · %s — %d u." % (e["label"], e["units"]))
    if v["issues"]:
        print("  ⚠ Incidencias:")
        for i in v["issues"]:
            print("    · %s" % i)

    # Si el archivo mezcla marcas, sus totales incluyen ventas ajenas a GOA:
    # se quedan en la consola y en el JSON solo va si cuadró.
    for check in v["checks"]:
        if check.pop("mixedBrands"):
            check["excelTotal"] = check["computed"] = None
    js = "const altaGamaData = " + json.dumps(report, indent=2, ensure_ascii=False) + ";\n"
    OUTPUT_FILE.write_text(js, encoding="utf-8")


if __name__ == "__main__":
    main()
