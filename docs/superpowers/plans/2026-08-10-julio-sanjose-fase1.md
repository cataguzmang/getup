# Fase 1 — Julio 2026: plomería compartida + San José — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Absorber los cambios de layout del Excel de julio 2026 (C1–C6 del spec `2026-08-10-julio-sanjose-design.md`) sin alterar en un solo byte los nueve meses históricos de San José.

**Architecture:** La normalización del estado vive en `split_excel.py` (D1): `fuentes/` garantiza valores canónicos de `Company`. Una hoja irresoluble aborta la corrida antes de escribir (D2). Marca presente sin ventas → archivo solo-encabezado en el mes dominante (D3). El cross-check de `SOLD` acepta entregadas o pagadas (D4); el campo `base` se emite **solo** cuando calza contra pagadas, para que el data.js histórico quede idéntico (la regresión es la prueba que manda).

**Tech Stack:** Python 3 + openpyxl + pytest. Dashboard: HTML/JS estático (puerto 8137).

**Restricciones:** `split_excel.py --only GET`. No tocar `report_core.py`, `goa-inventory/parse_excel.py` ni dashboards GOA/KOM/ROB. Único cambio en `reporte-sanjose/index.html`: mostrar la base del cross-check. Sin merge a main, sin push.

---

### Task 0: Spec en la rama

**Files:**
- Create: `docs/superpowers/specs/2026-08-10-julio-sanjose-design.md` (ya extraído de e8ca411)

- [x] **Step 1:** `git show e8ca411:docs/... > docs/...` (hecho)
- [ ] **Step 2: Commit**

```bash
git add docs/superpowers/specs/2026-08-10-julio-sanjose-design.md docs/superpowers/plans/2026-08-10-julio-sanjose-fase1.md
git commit -m "docs: spec + plan fase 1 — plomeria compartida + San Jose julio 2026"
```

### Task 1: C1 — Normalizar `Company` al escribir el canónico

**Files:**
- Modify: `split_excel.py` (constante `CANONICAL_COMPANY`, `canonical_row`, `collect`)
- Test: `tests/test_estado_canonico.py` (nuevo)

- [ ] **Step 1: Tests que fallan** — `tests/test_estado_canonico.py`:

```python
"""C1/C2 del spec 2026-08-10: Company canónico en fuentes/ y aborto por estado irresoluble."""
import openpyxl

import split_excel as sx
from helpers import HEADER, tx, make_workbook, D


def _company_col(path):
    """Valores de la columna Company del archivo canónico escrito."""
    ws = openpyxl.load_workbook(path).active
    rows = list(ws.iter_rows(values_only=True))
    return [r[sx.COL_COMPANY] for r in rows[1:]]


def test_company_ausente_hoja_ny_escribe_canonico(tmp_path, monkeypatch):
    """Julio: las hojas NY vienen sin columna Company → se escribe LatinFood US Corp."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"; entrada.mkdir()
    hdr = [h for h in HEADER if h != "Company"]          # 10 columnas, sin Company
    row = [D(2026, 7, 6), "S1", "[GET01] Jack Mackerel in Brine",
           "Key Food", "Katherine", 8, 8, 8, 34.2, 273.6]
    make_workbook(entrada / "07. GET Jul 2026.xlsx",
                  {"NY Jul - Jack Mackerel": [hdr, row]})
    rc = sx.main(["--no-build"])
    assert rc == 0
    out = tmp_path / "reporte-sanjose" / "fuentes" / "San-Jose-2026-07.xlsx"
    assert _company_col(out) == ["LatinFood US Corp."]


def test_company_basura_hoja_mia_escribe_canonico(tmp_path, monkeypatch):
    """Julio: MIA trae 'Florida - LF' / 'Florida- LatinFood FL Inc.' → LatinFood Florida."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"; entrada.mkdir()
    make_workbook(entrada / "07. GET Jul 2026.xlsx", {"MIA Jul - Jack Mackerel": [
        HEADER,
        tx(D(2026, 7, 2), "S1", "[GET01] x", "c", "sp", company="Florida - LF"),
        tx(D(2026, 7, 3), "S2", "[GET01] x", "c", "sp",
           company="Florida- LatinFood FL Inc."),
    ]})
    rc = sx.main(["--no-build"])
    assert rc == 0
    out = tmp_path / "reporte-sanjose" / "fuentes" / "San-Jose-2026-07.xlsx"
    assert _company_col(out) == ["LatinFood Florida", "LatinFood Florida"]
```

- [ ] **Step 2:** `python -m pytest tests/test_estado_canonico.py -v` → FAIL (Company escrito crudo: `None` / `Florida - LF`).
- [ ] **Step 3: Implementación** en `split_excel.py`:

```python
# Debajo de STATE_BY_SHEET_PREFIX:
CANONICAL_COMPANY = {"Florida": "LatinFood Florida", "Nueva York": "LatinFood US Corp."}

# canonical_row gana el estado resuelto:
def canonical_row(row, cols, state=None):
    """Rearma la fila en el orden canónico de HEADER según el mapa de columnas.
    Con `state` resuelto, escribe en Company el valor canónico del estado (la
    info venía en el nombre de la hoja; el Company crudo del distribuidor cambia)."""
    vals = [get(row, cols, f) for f in CANONICAL_FIELDS]
    if state in CANONICAL_COMPANY:
        vals[COL_COMPANY] = CANONICAL_COMPANY[state]
    return vals
```

En `collect()`, dentro del loop de filas (reemplaza el cuerpo desde `crow = ...`):

```python
crow = canonical_row(row, cols)
if not is_transaction_row(crow):
    continue
prefix = sku_prefix(crow[COL_VARIANT])
mk = month_key(crow[COL_DATE])
if prefix not in BRANDS:
    unmatched.append(crow)
    continue
state = state_of(crow, ws.title)
crow = canonical_row(row, cols, state)
rows[prefix][mk].append(crow)
stats[prefix][mk][state] += 1
tx_codes.append(prefix)
tx_months.append(mk)
tx_states.append(state)
```

- [ ] **Step 4:** `python -m pytest tests/test_estado_canonico.py tests/ -v` → todo PASS.
- [ ] **Step 5: Commit** — `feat: split_excel escribe Company canónico según el estado resuelto (C1)`

### Task 2: C2 — Abortar si una hoja queda sin estado

**Files:**
- Modify: `split_excel.py` (`Collected.unresolved`, `collect`, `main`)
- Test: `tests/test_estado_canonico.py`

- [ ] **Step 1: Test que falla** (agregar a `tests/test_estado_canonico.py`):

```python
def test_estado_irresoluble_aborta_sin_escribir(tmp_path, monkeypatch, capsys):
    """Hoja con transacciones y estado '?' → rc != 0, fuentes/ sin tocar, error con
    archivo + hoja + Company."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"; entrada.mkdir()
    make_workbook(entrada / "07. GET Jul 2026.xlsx", {"Julio - Jack Mackerel": [
        HEADER,
        tx(D(2026, 7, 2), "S1", "[GET01] x", "c", "sp", company="Otra Cosa"),
    ]})
    rc = sx.main(["--no-build"])
    assert rc != 0
    assert not (tmp_path / "reporte-sanjose" / "fuentes").exists()
    assert not (tmp_path / "unmatched").exists()
    out = capsys.readouterr().out
    assert "07. GET Jul 2026.xlsx" in out
    assert "Julio - Jack Mackerel" in out
    assert "Otra Cosa" in out
```

- [ ] **Step 2:** correr → FAIL (hoy rc==0 y escribe con estado '?').
- [ ] **Step 3: Implementación**: `Collected` gana `unresolved: list` (tuplas `(archivo, hoja, {companies})`). En `collect()`, tras el loop de filas de cada hoja, registrar la hoja si alguna fila de marca conocida quedó en `?`:

```python
# en el loop de filas, acumular:
if state == "?":
    bad_companies.add(crow[COL_COMPANY] if crow[COL_COMPANY] is not None else "(sin Company)")
# tras el loop de la hoja:
if bad_companies:
    unresolved.append((Path(path).name, ws.title, sorted(map(str, bad_companies))))
```

En `main()`, justo después de `data = collect(files)` y antes de escribir nada:

```python
if data.unresolved:
    print(f"✗ split_excel.py — {', '.join(f.name for f in files)}\n")
    for fname, sheet, companies in data.unresolved:
        print(f"  ✗ {fname} / '{sheet}': estado irresoluble "
              f"(Company: {', '.join(companies)})")
    print("\n✗ No se escribió nada en fuentes/ ni se regeneró ningún data.js.")
    return 1
```

- [ ] **Step 4:** `python -m pytest tests/ -v` → todo PASS.
- [ ] **Step 5: Commit** — `feat: split_excel aborta si una hoja queda sin estado resoluble (C2)`

### Task 3: C3 — Mes en cero para marca presente sin ventas

**Files:**
- Modify: `split_excel.py` (`BRAND_BY_SHEET_HINT`, `brand_from_sheet_name`, `Collected.empty_months`, `collect`, `main`)
- Test: `tests/test_mes_cero.py` (nuevo)

- [ ] **Step 1: Tests que fallan** — `tests/test_mes_cero.py`:

```python
"""C3 del spec 2026-08-10: marca presente por nombre de hoja pero sin ventas →
archivo canónico solo con encabezado, en el mes dominante del archivo."""
import openpyxl

import split_excel as sx
from helpers import HEADER, tx, make_workbook, D


def test_marca_presente_sin_ventas_escribe_solo_encabezado(tmp_path, monkeypatch):
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"; entrada.mkdir()
    make_workbook(entrada / "07. GET Jul 2026.xlsx", {
        "MIA Jul - Jack Mackerel": [
            HEADER,
            tx(D(2026, 7, 2), "S1", "[GET01] x", "c", "sp", company="LatinFood Florida"),
        ],
        "NY Jul - Kombuchacha": [HEADER],   # presente, sin ventas ("No sales reported")
    })
    rc = sx.main(["--no-build"])
    assert rc == 0
    kom = tmp_path / "kombuchacha" / "fuentes" / "KOM-2026-07.xlsx"
    assert kom.exists()
    ws = openpyxl.load_workbook(kom)["KOM"]
    assert list(ws.iter_rows(values_only=True)) == [tuple(sx.HEADER)]
    # ROB no aparece en ninguna hoja → no se inventa archivo
    assert not (tmp_path / "robinson-crusoe" / "fuentes").exists()


def test_archivo_sin_transacciones_no_escribe_nada(tmp_path, monkeypatch):
    """Sin transacciones no hay mes que atribuir: no se escribe nada."""
    monkeypatch.setattr(sx, "ROOT", tmp_path)
    entrada = tmp_path / "entrada"; entrada.mkdir()
    make_workbook(entrada / "vacio.xlsx", {"NY Jul - Kombuchacha": [HEADER]})
    rc = sx.main(["--no-build"])
    assert rc == 0
    assert not (tmp_path / "kombuchacha" / "fuentes").exists()


def test_brand_from_sheet_name():
    f = sx.brand_from_sheet_name
    assert f("NY Jul - Kombuchacha") == "KOM"
    assert f("MIA Jul - Salmón y  Mejillones") == "ROB"   # acentos y espacios dobles
    assert f("MIA Jul - Jack Mackerel") == "GET"
    assert f("MIA Jun - Jack Mckarel") == "GET"           # el typo de junio
    assert f("MIA Jul - GOA") == "GOA"
    assert f("Hoja Cualquiera") is None
```

- [ ] **Step 2:** correr → FAIL (`brand_from_sheet_name` no existe).
- [ ] **Step 3: Implementación** en `split_excel.py`:

```python
import unicodedata   # arriba, con los demás imports

# Marca por nombre de hoja (subcadena, sin mayúsculas ni acentos) — para registrar
# marcas presentes aunque no aporten filas (mes en cero, D3).
BRAND_BY_SHEET_HINT = {
    "jack mackerel": "GET", "jack mckarel": "GET",
    "kombuchacha": "KOM",
    "salmon y mejillones": "ROB",
    "goa": "GOA",
}

def brand_from_sheet_name(name):
    """Código de marca por subcadena del nombre de hoja; None si no reconoce."""
    n = unicodedata.normalize("NFD", str(name))
    n = "".join(c for c in n if not unicodedata.combining(c))
    n = re.sub(r"\s+", " ", n).strip().lower()
    for hint, code in BRAND_BY_SHEET_HINT.items():
        if hint in n:
            return code
    return None
```

`Collected` gana `empty_months: dict` (code → set de meses). En `collect()`, por archivo: acumular `file_months` (Counter de meses de todas las transacciones), `present` (marcas por nombre de hoja) y `with_rows` (marcas que aportaron filas en este archivo). Al cerrar el archivo:

```python
if file_months:
    dom_month = file_months.most_common(1)[0][0]
    for code in present - with_rows:
        empty_months[code].add(dom_month)
```

En `main()`, en el loop de marcas (reemplaza `months = ...` / `if not months: continue`):

```python
months = data.rows.get(code, {})
empty = sorted(set(data.empty_months.get(code, ())) - set(months))
if not months and not empty:
    continue
```

y tras el loop `for mk in sorted(months):`, escribir los meses en cero:

```python
for mk in empty:
    out = ROOT / brand.folder / "fuentes" / f"{brand.code}-{mk}.xlsx"
    write_canonical(out, [], [], brand.sheet)
    rel = out.relative_to(ROOT).as_posix()
    print(f"    → {rel}   (mes en cero: presente sin ventas)")
```

- [ ] **Step 4:** `python -m pytest tests/ -v` → todo PASS.
- [ ] **Step 5: Commit** — `feat: split_excel registra mes en cero para marca presente sin ventas (C3)`

### Task 4: C4 — El monto del descuento se lee de su columna

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (`_parse_incentive_segment`)
- Test: `reporte-sanjose/tests/test_incentivos.py`

- [ ] **Step 1: Test que falla** (agregar a `test_incentivos.py`):

```python
def test_descuento_sin_monto_no_agarra_el_free(tmp_path, monkeypatch):
    """Julio: 'In Brine == > 0 cases' con la celda de monto vacía → $0, no el 5
    de la columna FREE (bug C4: $305 en vez de $300)."""
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    rows = [
        canon_row(D(2026, 7, 2), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", 10, 38.4, 384),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 6, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 20, 5, "In Brine == > 0 cases"]),   # monto vacío
        _pad(["Angela Quimbay", 31, 5, "Tomate == > 12 cases", 300]),
    ]
    make_canonical(fuentes / "San-Jose-2026-07.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    fl = g.load_incentivos()[(2026, 7)]["Florida"]
    productos = {d["product"]: d for d in fl["descuentos"]}
    assert productos["In Brine"]["amount"] == 0.0
    assert productos["Tomate"]["amount"] == 300
    assert fl["totalDescuentos"] == 300.0
```

- [ ] **Step 2:** correr → FAIL (`In Brine` amount == 5.0, total 305).
- [ ] **Step 3: Implementación** en `_parse_incentive_segment` (línea del amount):

```python
# El monto vive de la columna de montos en adelante (después de la etiqueta
# en col D): _last_number sobre la fila completa agarraría el FREE (col C)
# cuando la celda de monto viene vacía (julio: $305 en vez de $300).
amount = _last_number(row[4:]) or 0.0
```

- [ ] **Step 4:** `python -m pytest reporte-sanjose/tests/ -v` → todo PASS (regresión incluida: para junio el monto vive en col E+, mismo resultado).
- [ ] **Step 5: Commit** — `fix: el monto del descuento se lee de su columna, no del último número (C4)`

### Task 5: C5 — Cross-check de `SOLD` con dos bases

**Files:**
- Modify: `reporte-sanjose/generar_data.py` (`build_period_data`, bloque `cross_check`)
- Test: `reporte-sanjose/tests/test_crosscheck.py`

- [ ] **Step 1: Tests que fallan** (agregar a `test_crosscheck.py`):

```python
def _period_con_sold(tmp_path, monkeypatch, pagadas, promo, sold_reportado, mes=7):
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    rows = [
        canon_row(D(2026, mes, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", pagadas, 38.4, pagadas * 38.4),
        canon_row(D(2026, mes, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", promo, 38.4, 0),
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, mes - 1, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", sold_reportado, promo]),
    ]
    make_canonical(fuentes / f"San-Jose-2026-{mes:02d}.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    r = g.load_rows()
    months = sorted(set(g.month_key(x["date"]) for x in r))
    op, pa, ga = g.get_order_unit_prices(r)
    inc = g.load_incentivos()
    return g.build_period_data(r, months, op, pa, ga, inc, {}, month_set=set(months))


def test_crosscheck_sold_base_entregadas_sin_campo(tmp_path, monkeypatch):
    """Junio: SOLD = entregadas. Calza como siempre y NO gana el campo base
    (el data.js histórico debe quedar idéntico)."""
    cc = _period_con_sold(tmp_path, monkeypatch, pagadas=20, promo=21,
                          sold_reportado=41, mes=6)["crossCheck"]["Florida"]
    assert cc["sold"] == {"computado": 41, "reportado": 41}


def test_crosscheck_sold_base_pagadas(tmp_path, monkeypatch):
    """Julio: SOLD = pagadas (entregadas − free). Calza contra pagadas y lo registra."""
    cc = _period_con_sold(tmp_path, monkeypatch, pagadas=51, promo=10,
                          sold_reportado=51)["crossCheck"]["Florida"]
    assert cc["sold"] == {"computado": 51, "reportado": 51, "base": "pagadas"}


def test_crosscheck_sold_sin_base_no_calza(tmp_path, monkeypatch):
    """SOLD que no calza ni con entregadas ni con pagadas → mismatch real."""
    cc = _period_con_sold(tmp_path, monkeypatch, pagadas=20, promo=21,
                          sold_reportado=40)["crossCheck"]["Florida"]
    assert cc["sold"] == {"computado": 41, "reportado": 40}   # sin base, computado ≠ reportado
```

- [ ] **Step 2:** correr → FAIL (base_pagadas: computado 61 ≠ 51, sin campo base).
- [ ] **Step 3: Implementación** — reemplazar el bloque `cross_check` de `build_period_data`:

```python
# Cross-check: cajas calculadas (de transacciones) vs. reportadas (del bloque).
# El SOLD del vendedor cambió de base entre meses (junio = entregadas,
# julio = pagadas): calza contra cualquiera de las dos y `base` registra cuál.
# Para la base histórica (entregadas) el campo se omite: el data.js de los
# meses ya publicados debe quedar idéntico (regresión).
cross_check = {}
for st, sd in states.items():
    entregadas = sd['cajas']
    pagadas = max(entregadas - sd['pc'], 0)
    rep = sold_rep_by_state.get(st, 0)
    sold = {"computado": entregadas, "reportado": rep}
    if rep > 0 and rep != entregadas and rep == pagadas:
        sold = {"computado": pagadas, "reportado": rep, "base": "pagadas"}
    cross_check[st] = {
        "free": {"computado": sd['pc'], "reportado": free_rep_by_state.get(st, 0)},
        "sold": sold,
    }
```

- [ ] **Step 4:** `python -m pytest reporte-sanjose/tests/ -v` → todo PASS.
- [ ] **Step 5: Commit** — `feat: cross-check de SOLD acepta entregadas o pagadas y registra la base (C5)`

### Task 6: C6 — `index.html` muestra la base validada

**Files:**
- Modify: `reporte-sanjose/index.html` (`renderCrossCheck`, rama sin issues)

- [ ] **Step 1: Implementación** — en la rama `issues.length===0` de `renderCrossCheck`:

```js
if(issues.length===0){const bases=Object.keys(cc).map(function(st){const s=cc[st].sold;return s&&s.base?st+': SOLD '+s.reportado+' = cajas '+s.base:'';}).filter(Boolean);el.innerHTML='<span class="cc-ok"><svg viewBox="0 0 24 24" fill="none" stroke-width="2"><path d="M20 6 9 17l-5-5"/></svg>Datos validados'+(bases.length?' · '+bases.join(' · '):'')+'</span>';return;}
```

(la etiqueta viene resuelta de `generar_data.py` — el JS solo la muestra; sin `base` el texto queda como hoy)

- [ ] **Step 2:** verificación visual en Task 9 (preview). Commit — `feat: dashboard indica contra qué base validó el cross-check de SOLD (C6)`

### Task 7: Regresión histórica (la prueba que manda)

- [ ] **Step 1:** con `fuentes/` sin julio (estado actual), correr `python generar_data.py` en `reporte-sanjose/`.
- [ ] **Step 2:** comparar contra main ignorando solo la línea `// Generado:`:

```bash
git show main:reporte-sanjose/data.js | grep -v '^// Generado:' > /tmp/main-data.js
grep -v '^// Generado:' reporte-sanjose/data.js > /tmp/new-data.js
diff /tmp/main-data.js /tmp/new-data.js && echo IDENTICO
```

Expected: `IDENTICO`. Si difiere → arreglar la plomería antes de seguir.

### Task 8: Corrida real de julio + números

- [ ] **Step 1:** `python split_excel.py --only GET` (entrada/ tiene junio y julio). Expected: escribe `San-Jose-2026-06.xlsx` y `San-Jose-2026-07.xlsx`, orientación `normal`, regenera `data.js`, rc 0.
- [ ] **Step 2:** verificar números de julio contra el spec (script scratch que parsea `GENERAL_DATA["Jul 26"]`): NY $820.80/24/6/3/3 · FL $2,371.20/61/10/3/3 · Total $3,192.00/85/16/6/6 · descuentos FL $300.00 · neto $2,892.00 · crossCheck FL sold `{computado:51, reportado:51, base:"pagadas"}`.
- [ ] **Step 3:** re-verificar la regresión con el June regenerado por el split nuevo: mover `San-Jose-2026-07.xlsx` fuera de `fuentes/`, regenerar, diff vs main (== IDENTICO), restaurar julio, regenerar.
- [ ] **Step 4:** actualizar `reporte-sanjose/tests/reference_sp3.json` con el data.js nuevo (incluye julio) usando `_parse_data_js` — es la práctica mensual: la referencia guarda los números verificados.
- [ ] **Step 5:** `python -m pytest tests/ reporte-sanjose/tests/ -v` → todo verde.
- [ ] **Step 6: Commit** — `data: San José julio 2026 (fuentes regenerados, data.js, referencia de regresión)`

### Task 9: Dashboard en el navegador

- [ ] **Step 1:** preview `sanjose-dashboard` (puerto 8137), recargar.
- [ ] **Step 2:** consola sin errores; `Jul 2026` en el selector; al elegirlo, cross-check `Datos validados · Florida: SOLD 51 = cajas pagadas`.
- [ ] **Step 3:** screenshot como prueba.
