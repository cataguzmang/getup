# SP3b — Rediseño del dashboard de San José — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans for the DATA tasks (1–4). The HTML redesign tasks (5–7) are built directly with browser preview against the approved mockups — not TDD. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Rediseñar el dashboard de San José para mostrar revenue neto, descuentos por producto/estado, clientes nuevos y un cross-check discreto, moviendo además todo cálculo derivado al data layer (Python).

**Architecture:** Pieza 1 (Python, TDD): `generar_data.py` gana helpers de razón (`_promo_pct`, `_roi`, `_rev_por_caja`), expone esas razones + `crossCheck` en `summary`/`states`, y agrega `D.total` + series `promoPct`/`roi` por mes. Pieza 2 (HTML, preview): `index.html` se rediseña consumiendo solo valores precalculados; el JS queda como vista (formateo + render + el helper `pct()`).

**Tech Stack:** Python 3.14, openpyxl, pytest; HTML + Chart.js 4.4.1 (CDN), sin build.

**Spec:** [docs/superpowers/specs/2026-07-09-sp3b-sanjose-dashboard-design.md](../specs/2026-07-09-sp3b-sanjose-dashboard-design.md)

**Regla rectora:** todo cálculo derivado se hace en funciones explícitas de `generar_data.py`. El dashboard solo formatea y renderiza.

---

## Task 1: Razones derivadas al data layer (`promoPct`, `roi`, `revPorCaja`)

**Files:** Modify `reporte-sanjose/generar_data.py`; Test `reporte-sanjose/tests/test_razones.py`.

- [ ] **Step 1: Test que falla** — `reporte-sanjose/tests/test_razones.py`:
```python
import generar_data as g


def test_helpers_de_razon():
    assert g._promo_pct(100, 25) == 25.0
    assert g._promo_pct(0, 0) == 0.0
    assert g._roi(900, 300) == 3.0
    assert g._roi(900, 0) == 0.0
    assert g._rev_por_caja(912, 20, 0) == 45.6
    assert g._rev_por_caja(0, 0, 0) == 0.0


def test_summary_y_states_traen_razones(tmp_path, monkeypatch):
    from sjhelpers import make_canonical, canon_row, D
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 6, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 20, 45.6, 912),
        canon_row(D(2026, 6, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "SP", "LatinFood Florida", 5, 45.6, 0),   # promo
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    period = g.build_period_data(rows, months, op, pa, ga, month_set=set(months))
    s = period["summary"]
    assert s["promoPct"] == g._promo_pct(s["cajas"], s["pc"])
    assert s["roi"] == g._roi(s["rev"], s["pv"])
    assert s["revPorCaja"] == g._rev_por_caja(s["rev"], s["cajas"], s["pc"])
    fl = period["states"]["Florida"]
    assert fl["promoPct"] == g._promo_pct(fl["cajas"], fl["pc"])
    assert "roi" in fl and "revPorCaja" in fl
```

- [ ] **Step 2: Correr → falla** (`AttributeError: _promo_pct`).

- [ ] **Step 3: Agregar helpers** en `generar_data.py` (antes de `build_period_data`):
```python
def _promo_pct(cajas, pc):
    return round(pc / cajas * 100, 1) if cajas > 0 else 0.0


def _roi(rev, pv):
    return round(rev / pv, 2) if pv > 0 else 0.0


def _rev_por_caja(rev, cajas, pc):
    paid = cajas - pc
    return round(rev / paid, 2) if paid > 0 else 0.0
```

- [ ] **Step 4: Sumar razones al `summary`.** En `build_period_data`, en el dict `summary`, tras `'clientesNuevos': len(nuevos),` agregar:
```python
        'promoPct': _promo_pct(total_cajas, total_pc),
        'roi': _roi(total_rev, total_pv),
        'revPorCaja': _rev_por_caja(total_rev, total_cajas, total_pc),
```

- [ ] **Step 5: Sumar razones a cada `states[...]`.** En el `for state_name, d in state_data.items()`, dentro del dict `states[state_name]`, tras `'revNeto': round(d['rev'] - dd, 2),` agregar:
```python
            'promoPct': _promo_pct(d['cajas'], d['pc']),
            'roi': _roi(d['rev'], d['pv']),
            'revPorCaja': _rev_por_caja(d['rev'], d['cajas'], d['pc']),
```
Y en el bloque de estados faltantes (`for missing in ('Florida', 'Nueva York')`), agregar al dict por defecto:
```python
                               'promoPct': 0.0, 'roi': 0.0, 'revPorCaja': 0.0,
```

- [ ] **Step 6: Correr → pasa.** `python -m pytest reporte-sanjose/tests/test_razones.py -v`

- [ ] **Step 7: Commit**
```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_razones.py
git commit -m "feat: razones (promoPct, roi, revPorCaja) precalculadas en el data layer"
```

---

## Task 2: `crossCheck` en `build_period_data`

**Files:** Modify `reporte-sanjose/generar_data.py`; Test `reporte-sanjose/tests/test_crosscheck.py`.

- [ ] **Step 1: Test que falla** — `reporte-sanjose/tests/test_crosscheck.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g

HEADER_LEN = 11
def _pad(row): return list(row) + [None] * (HEADER_LEN - len(row))


def test_crosscheck_free_match(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    rows = [
        canon_row(D(2026, 6, 1), "S1", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", 20, 45.6, 912),
        canon_row(D(2026, 6, 2), "S2", "[GET01] Jack Mackerel in Brine",
                  "C", "Alexandra J", "LatinFood Florida", 21, 45.6, 0),  # 21 promo
    ]
    block = [
        _pad(["__INCENTIVOS__", "Florida"]),
        _pad([D(2026, 5, 1), "SOLD", "FREE", "Descuentos"]),
        _pad(["Alexandra Jimenez", 20, 21, None, None]),   # SOLD 20, FREE 21
    ]
    make_canonical(fuentes / "San-Jose-2026-06.xlsx", rows + block)
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    r = g.load_rows()
    months = sorted(set(g.month_key(x["date"]) for x in r))
    op, pa, ga = g.get_order_unit_prices(r)
    inc = g.load_incentivos()
    period = g.build_period_data(r, months, op, pa, ga, inc, {}, month_set=set(months))
    cc = period["crossCheck"]["Florida"]
    assert cc["free"]["computado"] == 21 and cc["free"]["reportado"] == 21
    assert cc["sold"]["reportado"] == 20   # sold reportado del bloque
```

- [ ] **Step 2: Correr → falla** (`KeyError: 'crossCheck'`).

- [ ] **Step 3: Agregar la agregación y el `crossCheck`.** En `build_period_data`, junto al cálculo de `desc_by_state` (en el mismo `for ym in month_set` o uno nuevo), agregar los reportados:
```python
    free_rep_by_state = defaultdict(int)
    sold_rep_by_state = defaultdict(int)
    for ym in month_set:
        for state, inc in incentivos.get(ym, {}).items():
            free_rep_by_state[state] += sum(inc["freeReportado"].values())
            sold_rep_by_state[state] += sum(inc["soldReportado"].values())
```
Después de construir el dict `states` (con Florida y Nueva York garantizados), construir el cross-check:
```python
    cross_check = {}
    for st, sd in states.items():
        cross_check[st] = {
            "free": {"computado": sd['pc'], "reportado": free_rep_by_state.get(st, 0)},
            "sold": {"computado": sd['cajas'] - sd['pc'], "reportado": sold_rep_by_state.get(st, 0)},
        }
```
En el `return`, agregar la clave:
```python
        'crossCheck': cross_check,
```

- [ ] **Step 4: Correr → pasa.** `python -m pytest reporte-sanjose/tests/test_crosscheck.py -v`

- [ ] **Step 5: Commit**
```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_crosscheck.py
git commit -m "feat: crossCheck (free/sold calculado vs reportado) por estado en data.js"
```

---

## Task 3: `D.total` + series `promoPct`/`roi` por mes

**Files:** Modify `reporte-sanjose/generar_data.py` (`build_monthly_state_data`); Test `reporte-sanjose/tests/test_series_mensuales.py`.

- [ ] **Step 1: Test que falla** — `reporte-sanjose/tests/test_series_mensuales.py`:
```python
from sjhelpers import make_canonical, canon_row, D
import generar_data as g


def test_d_total_y_series(tmp_path, monkeypatch):
    fuentes = tmp_path / "fuentes"; fuentes.mkdir()
    make_canonical(fuentes / "m.xlsx", [
        canon_row(D(2026, 6, 1), "S1", "[GET01] x", "C", "SP", "LatinFood Florida", 20, 45.6, 912),
        canon_row(D(2026, 6, 2), "S2", "[GET01] x", "C", "SP", "LatinFood US Corp.", 10, 45.6, 456),
        canon_row(D(2026, 6, 3), "S3", "[GET01] x", "C", "SP", "LatinFood Florida", 5, 45.6, 0),  # promo
    ])
    monkeypatch.setattr(g, "SOURCES_DIR", fuentes)
    rows = g.load_rows()
    months = sorted(set(g.month_key(r["date"]) for r in rows))
    op, pa, ga = g.get_order_unit_prices(rows)
    Dd = g.build_monthly_state_data(rows, months, op, pa, ga)
    i = 0
    # total = fl + ny
    assert Dd["total"]["rev"][i] == round(Dd["fl"]["rev"][i] + Dd["ny"]["rev"][i], 2)
    assert Dd["total"]["cajas"][i] == Dd["fl"]["cajas"][i] + Dd["ny"]["cajas"][i]
    # series de razón presentes y coherentes
    assert Dd["fl"]["promoPct"][i] == g._promo_pct(Dd["fl"]["cajas"][i], Dd["fl"]["pc"][i])
    assert Dd["total"]["roi"][i] == g._roi(Dd["total"]["rev"][i], Dd["total"]["pv"][i])
```

- [ ] **Step 2: Correr → falla** (`KeyError: 'total'`).

- [ ] **Step 3: Extender `build_monthly_state_data`.** En el `result[state] = {...}` inicial de cada estado, agregar `'promoPct': [], 'roi': []`. Dentro del `for ym in months`, tras los `append` existentes, agregar:
```python
            result[state]['promoPct'].append(_promo_pct(d['cajas'], d['pc']))
            result[state]['roi'].append(_roi(d['rev'], d['pv']))
```
Antes del `return result`, construir el combinado `total`:
```python
    n = len(months)
    tot = {k: [] for k in ('rev', 'cajas', 'pv', 'pc', 'ord', 'descuentos', 'revNeto', 'promoPct', 'roi')}
    for i in range(n):
        trev = round(result['fl']['rev'][i] + result['ny']['rev'][i], 2)
        tcaj = result['fl']['cajas'][i] + result['ny']['cajas'][i]
        tpv = round(result['fl']['pv'][i] + result['ny']['pv'][i], 2)
        tpc = result['fl']['pc'][i] + result['ny']['pc'][i]
        tdesc = round(result['fl']['descuentos'][i] + result['ny']['descuentos'][i], 2)
        tot['rev'].append(trev)
        tot['cajas'].append(tcaj)
        tot['pv'].append(tpv)
        tot['pc'].append(tpc)
        tot['ord'].append(result['fl']['ord'][i] + result['ny']['ord'][i])
        tot['descuentos'].append(tdesc)
        tot['revNeto'].append(round(trev - tdesc, 2))
        tot['promoPct'].append(_promo_pct(tcaj, tpc))
        tot['roi'].append(_roi(trev, tpv))
    result['total'] = tot
    return result
```

- [ ] **Step 4: Correr → pasa.** `python -m pytest reporte-sanjose/tests/test_series_mensuales.py -v`

- [ ] **Step 5: Commit**
```bash
git add reporte-sanjose/generar_data.py reporte-sanjose/tests/test_series_mensuales.py
git commit -m "feat: D.total (FL+NY) y series promoPct/roi por mes en el data layer"
```

---

## Task 4: Re-basear la referencia + regenerar

**Files:** Modify `reporte-sanjose/tests/reference_sp3.json` (regenerado); regenerar `reporte-sanjose/data.js`.

- [ ] **Step 1: Regenerar** (requiere las fuentes reales; si el archivo `San-Jose-2026-06.xlsx` está abierto en Excel, cerrarlo antes):
```bash
python reporte-sanjose/migrar_historico.py && python split_excel.py
```
Expected: San José regenerado ✓, sin SKUs no mapeados.

- [ ] **Step 2: Re-capturar la referencia** (ahora con `crossCheck`, `promoPct`, `roi`, `revPorCaja`, `D.total` y series nuevas):
```bash
python -c "import re,json; t=open('reporte-sanjose/data.js',encoding='utf-8').read(); g=lambda n,pat: json.loads(re.search(r'\b'+n+r' = ('+pat+r');',t,re.S).group(1)); ref={'MONTHS':g('MONTHS',r'\[.*?\]'),'D':g('D',r'\{.*?\}'),'GENERAL_DATA':g('GENERAL_DATA',r'\{.*\}')}; open('reporte-sanjose/tests/reference_sp3.json','w',encoding='utf-8').write(json.dumps(ref,ensure_ascii=False,indent=2)); print('claves D.fl:', list(ref['D']['fl'].keys())); print('total en D:', 'total' in ref['D'])"
```
Expected: imprime las series de `D.fl` incluyendo `promoPct`/`roi` y `total in D: True`.

- [ ] **Step 3: Verificar paridad** (los valores que antes calculaba el JS ahora vienen iguales):
```bash
python -c "import re,json; t=open('reporte-sanjose/data.js',encoding='utf-8').read(); d=json.loads(re.search(r'\bGENERAL_DATA = (\{.*\});',t,re.S).group(1)); a=d['ALL']['summary']; print('promoPct', a['promoPct'], '≈', round(a['pc']/a['cajas']*100,1)); cc=d['Jun 26']['crossCheck']['Florida']; print('free jun', cc['free'])"
```
Expected: `promoPct` coincide con `pc/cajas*100`; `free jun` computado 21 y reportado 21.

- [ ] **Step 4: Suite completa** → verde.
```bash
python -m pytest -q
```

- [ ] **Step 5: Commit** (referencia + data.js del data layer)
```bash
git add reporte-sanjose/tests/reference_sp3.json reporte-sanjose/data.js
git commit -m "data: re-basear referencia con crossCheck/razones/D.total (data layer SP3b)"
```

---

## Task 5: Rediseño del `index.html` — pestaña General (preview-driven)

**Files:** Modify `reporte-sanjose/index.html`. **No es TDD** — se construye con preview en el navegador contra los mockups aprobados. Cada bloque **consume campos precalculados**; el JS solo formatea (`f$`, `fN`) y arma DOM/charts.

- [ ] **Step 1: Lanzar preview del dashboard actual** para iterar en vivo. Usar `preview_start` (crear `.claude/launch.json` con un `python -m http.server` en `reporte-sanjose/` si no existe) o abrir el archivo; confirmar que carga sin errores de consola antes de tocar nada.

- [ ] **Step 2: Barra superior + cross-check discreto.** Junto al `period-pill`, agregar el indicador:
  - Regla: recorrer `crossCheck` del periodo visible; contar descuadres donde `reportado > 0` y `computado != reportado`.
  - 0 descuadres → texto tenue "Datos validados" con `ti-circle-check` muted.
  - ≥1 → aviso ámbar "N por revisar antes de publicar" + "Ver detalle" (despliega estado · métrica · `computado` vs `reportado`).

- [ ] **Step 3: Titular financiero (Opción B).** Reemplazar la fila de KPIs de solo-bruto por el flujo en tres cajas, leyendo del `summary`: `rev` (Bruto) → `descuentos` (−) → `revNeto` (Neto, en verde). Debajo, la tira de KPIs: Cajas entregadas (`cajas`), Cajas promo (`pc` + `promoPct`%), Clientes nuevos (`clientesNuevos`), Pedidos (`ord`), Precio/caja (`avg_price`).

- [ ] **Step 4: Por Estado (FL/NY).** A las tarjetas actuales, que hoy calculaban `pp`/`roi`/`rpc` en JS, cambiarlas para **leer** `states[e].promoPct`, `states[e].roi`, `states[e].revPorCaja`, y **agregar** `revNeto` y `descuentos`. (Eliminar los cálculos inline de `smHTML`.)

- [ ] **Step 5: Descuentos por producto.** Nueva sección: tabla/cards desde `descuentosPorProducto` (`product`, `state`, `cases`, `amount`) + total (`summary.descuentos`).

- [ ] **Step 6: Clientes nuevos.** Nueva sección: conteo (`summary.clientesNuevos`) + lista compacta colapsable desde `clientesNuevosLista` (`name`, `salesperson`).

- [ ] **Step 7: Evolución mensual.** Conservar los gráficos actuales pero alimentarlos desde `D.total` (en vez del objeto `A` que armaba el JS). Agregar una serie/gráfico de **Revenue neto** por mes (`D.*.revNeto`). Reemplazar `pFL`/`roiFL` inline por `D.*.promoPct` / `D.*.roi`.

- [ ] **Step 8: Verificar contra los mockups** en el navegador (titular B, KPIs, por-estado con neto/descuentos, descuentos por producto, clientes nuevos, cross-check discreto). Consola sin errores. El selector de período actualiza todo. Forzar (temporalmente, en datos de prueba) un descuadre para ver el aviso ámbar.

- [ ] **Step 9: Commit**
```bash
git add reporte-sanjose/index.html
git commit -m "feat: rediseño pestaña General — titular neto, descuentos, clientes nuevos, cross-check"
```

---

## Task 6: Rediseño del `index.html` — pestaña Mes vs. Mes (preview-driven)

**Files:** Modify `reporte-sanjose/index.html`.

- [ ] **Step 1: Variación Global.** Agregar tarjetas de **Revenue neto** (`D.total.revNeto`), **Descuentos** (`D.total.descuentos`, inverso; "nuevo" si el mes previo era 0) y **Clientes nuevos** (por mes) con su delta vía el helper `pct()`. Mantener las existentes; alimentar todo desde `D.total`/`D.fl`/`D.ny` (no recomputar sumas en JS).

- [ ] **Step 2: Por Estado (comparación).** Agregar filas de `revNeto`/`descuentos` a la comparación por estado (desde `D.fl`/`D.ny`).

- [ ] **Step 3: Descuentos y clientes nuevos del mes.** Para el mes actual de la comparación, mostrar `descuentosPorProducto` y `clientesNuevosLista` (de `GENERAL_DATA[mesActual]`).

- [ ] **Step 4: Series/gráficos.** Reemplazar cualquier cálculo inline restante (`roiFL`, `%promo` por mes) por `D.*.roi` / `D.*.promoPct`.

- [ ] **Step 5: Verificar** en el navegador contra el mockup de Mes vs. Mes. Consola limpia.

- [ ] **Step 6: Commit**
```bash
git add reporte-sanjose/index.html
git commit -m "feat: rediseño Mes vs. Mes — neto, descuentos y clientes nuevos con deltas"
```

---

## Task 7: Aceptación final

- [ ] **Step 1: Regenerar y abrir el dashboard real** (`python split_excel.py` si hace falta; abrir `reporte-sanjose/index.html`).
- [ ] **Step 2: Checklist visual** contra los mockups aprobados (General + Mes vs. Mes), en el mes con incentivos (junio) y en "Todos". Sin cálculo de negocio en el JS (revisar que `index.html` no recomputa razones/sumas — grep de `pc/`, `/pv`, `+D.ny`).
- [ ] **Step 3:** `python -m pytest -q` → verde. `git status` sin `.xlsx` trackeado.
- [ ] **Step 4:** Mostrar el dashboard real a la usuaria para visto bueno antes de mergear.
- [ ] **Step 5: Commit** de cualquier ajuste final del `data.js`/`index.html`.

---

## Verificación final

- [ ] `python -m pytest -q` verde (razones, crossCheck, series, regresión re-baseada).
- [ ] `data.js` trae `crossCheck`, `promoPct`, `roi`, `revPorCaja`, `D.total`, series `promoPct`/`roi`.
- [ ] `index.html` no hace cálculo de negocio inline (solo formateo + `pct()` + render).
- [ ] Dashboard coincide con los mockups aprobados; cross-check discreto funciona.

## Fuera de alcance

- Dashboards KOM/ROB; **SP4** (repo GetUp + subdominio privado).
