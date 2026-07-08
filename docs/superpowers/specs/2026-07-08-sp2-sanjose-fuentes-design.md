# SP2 — San José al modelo `fuentes/` (plomería, sin cambiar el dashboard)

**Fecha:** 2026-07-08
**Estado:** propuesto (pendiente de revisión de la usuaria)
**Sub-proyecto:** 2 de una serie (ver "Contexto y alcance")

---

## Contexto y alcance

En SP1 construimos `split_excel.py`, que ya deposita `reporte-sanjose/fuentes/San-Jose-2026-06.xlsx`
(formato canónico) pero **no** regenera el `data.js` de San José, porque su generador
todavía lee el Excel histórico viejo con otro esquema.

**SP2 es solo plomería:** migrar el histórico al modelo `fuentes/` y refactorizar
`generar_data.py` para que lea el **formato canónico**, **conservando exactamente la
lógica y los números actuales**. Al terminar, San José queda en el pipeline nuevo y el
`python split_excel.py` mensual también lo regenera.

**Explícitamente FUERA de alcance (va a SP3, con su propio brainstorming):**
- Arreglar la regla de montos `max(Total, Unit Price)` (tiene sesgos conocidos; ver más abajo).
- Incorporar el bloque de incentivos (SOLD/FREE/Descuentos/NEW CUSTOMERS) al dashboard.
- Revisar el modelo de promos in-line.
- Rediseñar el dashboard.

En SP2 el `index.html` **no se toca** y el `data.js` conserva su forma exacta.

---

## Objetivo

Que `python reporte-sanjose/generar_data.py` (y por lo tanto `python split_excel.py`)
produzca un `data.js` que:
- **Reproduce idénticos** los números de los meses históricos (Oct 2025 – May 2026) que
  hay hoy, y
- **Agrega junio 2026** (que ya está en `fuentes/` desde SP1).

Si los meses históricos cambian aunque sea un número, la plomería tiene un bug.

---

## Formato canónico (recordatorio)

Es el mismo de SP1 (una hoja, 11 columnas, passthrough fiel):

```
Order Date | Order | Product Variant | Customer | Salesperson | Company |
Qty Delivered | Qty Invoiced | Qty Ordered | Unit Price | Total
```

El **estado** (Florida / Nueva York) no es columna: se deriva de `Company`. El **producto**
sale de quitar el prefijo `[SKU]` de `Product Variant`.

---

## Pieza 1 — Migración única del histórico

**Script nuevo:** `reporte-sanjose/migrar_historico.py` (se corre una vez; idempotente —
sobrescribe).

Lee `reporte-sanjose/Historico Ventas Latin Food - San Jose.xlsx` (hoja `Historico`,
13 columnas) y escribe `reporte-sanjose/fuentes/San-Jose-historico-hasta-2026-05.xlsx`
(**un solo archivo**, hoja `San Jose`, formato canónico de 11 columnas).

Reordenamiento de columnas (verificado contra el archivo real):

| Canónico | ← Histórico |
|---|---|
| Order Date | Order Date (col 4) |
| Order | Order (col 3) |
| Product Variant | `Cod Prod` + " " + `Product` → `"[GET01] Jack Mackerel in Brine"` (cols 5+6) |
| Customer | Customer (col 2) |
| Salesperson | Salesperson (col 10) |
| Company | Company (col 0) |
| Qty Delivered / Invoiced / Ordered | cols 7 / 8 / 9 |
| Unit Price | Unit Price (col 12) |
| Total | Total (col 11) |

Reglas:
- **Fiel:** los montos se copian **tal cual** (mayo 2026 viene con Total/Unit Price
  intercambiados; eso lo resuelve el parser con la regla actual, sin normalizar aquí).
- Se **descarta** la columna `State` (redundante con `Company`).
- El `Historico...xlsx` original **queda como respaldo** (gitignoreado por `*.xlsx`, no se borra).

Hechos verificados que sostienen esto (sobre el archivo real, 211 filas con dato):
- `Company` ↔ `State` es 1:1 sin excepciones (`LatinFood Florida`↔Florida,
  `LatinFood US Corp`↔Nueva York) → derivar estado desde `Company` reproduce el estado viejo.
- 0 filas sin `Cod Prod` → el `Product Variant` siempre se reconstruye.

---

## Pieza 2 — Refactor de `generar_data.py`

**Solo cambia `load_rows()`** (y el constante de fuente + los mensajes que la nombran).
**Todo lo demás se conserva sin tocar**: `get_order_unit_prices`, `promo_box_price`,
`build_monthly_state_data`, `build_period_data`, `find_product_first_months`, la
serialización de `main()` y **la regla de montos actual**.

`load_rows()` nuevo:
1. Lee **todos** los `reporte-sanjose/fuentes/*.xlsx` (orden de nombre, sin temporales `~$`).
   Un mismo archivo puede traer varios meses; se separan por fecha aguas abajo (igual que GOA).
2. De cada workbook toma su hoja (primera / activa) y recorre las filas.
3. Filtra a **filas de pedido reales**: `datetime` en col A **y** `Product Variant` con
   patrón `[SKU]` en col C. Esto **ignora** el bloque de incentivos de junio (filas sin
   fecha o sin SKU), los separadores en blanco y cualquier encabezado.
4. Por cada fila real, construye el mismo dict que hoy espera el resto del código:
   - `state`: derivado de `Company` (`LatinFood Florida`→`Florida`,
     `LatinFood US Corp[.]`→`Nueva York`). Si `Company` no mapea, se **omite** la fila
     (equivalente al viejo `if not state: continue`).
   - `product`: `Product Variant` sin el prefijo `[SKU] `, `.strip()`.
   - `qty = Qty Delivered`; se **omite** si `qty <= 0` (igual que hoy).
   - `a = Total`, `b = Unit Price`; `is_promo = (a == 0 and b == 0)`;
     `line_rev = max(a, b)`; `box_price = line_rev / qty` — **idéntico a hoy** (SP3 lo revisa).
   - `company`, `customer`, `order`, `date`, `salesperson`: passthrough.

El resto del pipeline queda igual → el `data.js` mantiene su forma exacta
(`MONTHS, MONTHS_BTN, PAIRS, PLBLS, PERIOD_TEXT, PROD_FIRST_MONTH, D, GENERAL_DATA`),
ahora con Oct 2025 – **Jun 2026**.

> **Nota de deuda conocida (SP3):** la regla `line_rev = max(a, b)` sobrestima ingresos y
> subcuenta promos en ~13 filas históricas (ventas con descuento donde `Total < Unit Price`
> a qty 1, y cajas gratis con `Total = 0` pero `Unit Price ≠ 0`). Se **conserva** en SP2
> para no cambiar números; se arregla en SP3.

---

## Pieza 3 — Enganche a `split_excel.py`

Cambiar `BRANDS["GET"].sp1_ready` de `False` a `True`. Con eso, `python split_excel.py`
regenera San José automáticamente tras depositar su `fuentes/`. (El campo `generator` ya
apunta a `generar_data.py`.)

Actualizar el mensaje del resumen de SP1 no hace falta: al ser `sp1_ready=True`, ya
imprime "↻ data.js regenerado ✓" para GET.

---

## Verificación de regresión (el corazón de SP2)

**Antes de tocar código**, se captura una referencia: extraer de la `data.js` actual (la
commiteada, generada por el código viejo) los `GENERAL_DATA` por mes histórico y los
arreglos `D`, y guardarlos como fixture `reporte-sanjose/tests/reference_pre_sp2.json`.

Tras el refactor, un test de integración corre el `generar_data.py` nuevo (leyendo el
`fuentes/` real: histórico migrado + junio) y afirma:
- Para cada mes histórico `m` en `["Oct 25", ..., "May 26"]`:
  `GENERAL_DATA[m]` **idéntico** al de la referencia (summary, states, customerChart,
  salesChart, customerTable, productsCards).
- `D["fl"][:8]` y `D["ny"][:8]` (los 8 meses históricos) **idénticos** a la referencia.
- Aparece `"Jun 26"` en `MONTHS` y una novena entrada en `D`.

Si algo histórico difiere, es bug de migración/refactor y hay que corregirlo (no la referencia).

Verificación manual adicional: `git diff` de `reporte-sanjose/data.js` debe mostrar
**solo** agregados de junio (nuevo mes, nuevas entradas), no cambios en los bloques históricos.

---

## Robustez y casos borde

- **Fila sin `Company` mapeable:** se omite (no debería pasar con datos GET; se cuenta y avisa).
- **Bloque de incentivos de junio:** ignorado por el filtro de fila real (SP3 lo incorporará).
- **`Cod Prod` faltante en el histórico:** no ocurre (verificado); si ocurriera, `Product
  Variant` quedaría `" Producto"` — el migrador avisa si encuentra alguno.
- **Producto con espacios finales** (`"...Tomato Sauce  "`): se normaliza con `.strip()`
  igual que hoy.
- **Re-correr la migración:** sobrescribe el mismo archivo (idempotente).
- **Re-correr el split:** GOA y San José se regeneran; San José lee todo `fuentes/`.

---

## Pruebas (pytest)

Nueva carpeta `reporte-sanjose/tests/` (con `conftest.py` que hace importable
`generar_data.py` y `migrar_historico.py`, y helpers de workbook **dentro del conftest**
para no colisionar con `tests/helpers.py` de la raíz).

1. **Migración:** dado un Historico sintético (esquema viejo), el canónico resultante tiene
   header canónico, `Product Variant` reconstruido (`"[GET01] ..."`), `Company` preservado,
   sin columna `State`, y una fila por fila de dato.
2. **`load_rows` sobre canónico:** filtra filas reales (ignora bloque de incentivos y
   separadores); deriva estado desde `Company`; saca el nombre de producto sin `[SKU]`;
   omite `qty_delivered <= 0`; marca promo cuando ambos montos son 0.
3. **Regresión (integración):** reproduce idénticos los meses históricos vs
   `reference_pre_sp2.json` y agrega `Jun 26` (descrito arriba).
4. **`split_excel` engancha San José:** con `GET` marcado `sp1_ready=True`, `main` sí
   corre su generador (test con `ROOT` monkeypatched y un `generar_data.py` dummy, análogo
   al patrón de SP1).

---

## Decisiones ya tomadas

- Histórico como **un solo archivo** `San-Jose-historico-hasta-2026-05.xlsx`.
- Se **conserva** la regla `max(Total, Unit Price)` en SP2 (su arreglo es SP3).
- El bloque de incentivos y el rediseño del dashboard son **SP3**.
- El `Historico...xlsx` viejo queda de **respaldo** (no se borra).
- Estado derivado de `Company` (verificado consistente con la columna `State` vieja).

---

## Fuera de alcance (recordatorio)

- **SP3:** arreglo de la regla de montos + bloque de incentivos + revisión de promos in-line
  + rediseño del dashboard de San José.
- **SP4:** migración del repo al GitHub de GetUp + subdominio privado.
