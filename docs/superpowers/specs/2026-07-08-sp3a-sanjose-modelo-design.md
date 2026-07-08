# SP3a — San José v2: modelo de datos (regla de montos + incentivos + clientes nuevos)

**Fecha:** 2026-07-08
**Estado:** propuesto (pendiente de revisión de la usuaria)
**Sub-proyecto:** 3a (el 3b, rediseño del dashboard, va aparte con su propio brainstorming)

---

## Contexto y alcance

SP2 dejó a San José leyendo `fuentes/` canónico, **conservando la lógica vieja** (incluida
la regla de montos `max(Total, Unit Price)` y el ignorar el bloque de incentivos). SP3a
corrige y enriquece el **modelo de datos**; **no toca la UI** (`index.html`).

**En alcance (SP3a):**
1. **Arreglar la regla de montos** — orientación por mes + detección correcta de cajas gratis.
2. **Parsear el bloque de incentivos** por estado — descuentos por producto, total; free/sold
   como verificación cruzada; créditos de clientes nuevos.
3. **Calcular clientes nuevos** cruzando `customer` contra el histórico (primera aparición).
4. **Emitir un `data.js` enriquecido** (aditivo): revenue bruto y neto, descuentos por
   producto/estado, cajas gratis (con cross-check), clientes nuevos.

**Fuera de alcance:**
- **SP3b:** rediseñar el dashboard para mostrar todo esto (con mockups en el navegador).
- Dashboards de KOM/ROB; migración del repo (SP4).

SP3a **cambia números a propósito** (la regla vieja estaba mal). El `index.html` actual
sigue funcionando y mostrará los números corregidos + ignorará los campos nuevos hasta SP3b.

---

## Hechos del negocio (confirmados con la usuaria)

- Las líneas transaccionales traen **precio bruto**. Los **descuentos se restan después** de
  la factura → **Revenue neto = Revenue bruto − Descuentos**.
- **Descuentos**: en dinero, por **producto** (ej. "In Brine → 3 cajas, $88.68"; "Tomate →
  18 cajas, $540"; total $628.68), y pueden ser **por estado**.
- **SOLD / FREE** del bloque = resúmenes manuales (cajas vendidas / gratis por vendedor) →
  **solo verificación cruzada**, no fuente de verdad. Las cajas gratis reales salen de las
  líneas in-line.
- **Clientes nuevos**: el número del bloque no es confiable → se **calculan** desde las
  transacciones (primera aparición del cliente vs. meses anteriores).
- El bloque de incentivos puede aparecer **por estado** (hoy solo Florida; NY podría a futuro).

---

## Pieza 1 — Regla de montos: orientación por mes

**Problema (medido sobre el histórico real):** las columnas `Total` / `Unit Price` vienen
en dos orientaciones — normal (`Total = qty × UnitPrice`) y, en mayo 2026, **intercambiada**
(`UnitPrice = qty × Total`). La regla actual `line_total = max(Total, Unit Price)` falla en
~13 filas: ventas con descuento in-line (qty 1, `Total < UnitPrice`) y **cajas gratis con
`Total = 0` pero `Unit Price ≠ 0`** (que hoy se cuentan como venta y subcuentan promos).

**Solución:** en `load_rows`, **agrupar las filas por mes** y **detectar la orientación de
cada mes** con las filas inequívocas (qty > 1 y ambos montos ≠ 0):
- cuenta cuántas cumplen `Total ≈ qty × UnitPrice` (normal) vs. `UnitPrice ≈ qty × Total`
  (swap); gana la mayoría; si no hay señal suficiente, asume **normal**.
- Con la orientación del mes: `line_total` = la columna de total-de-línea correcta
  (`Total` en normal, `Unit Price` en swap); `precio_caja` = la otra columna.
- **Caja gratis** = `line_total == 0` con entrega > 0 (ya no exige que ambas columnas sean 0).
- El precio para valuar una caja gratis = `precio_caja` de esa fila si existe (> 0); si es 0,
  cae al fallback por orden/producto (ver Pieza 2).

**Importante — detección POR MES, no por archivo:** el histórico migrado trae Oct–May en un
solo archivo; Oct–Abr son normales y May es swap. Por eso la orientación se decide por mes.

## Pieza 2 — Valuación de cajas gratis (`pv`), estable

Deuda de SP2: `product_avg`/`global_avg` se calculaban sobre TODO el pool, así que sumar un
mes corría los `pv` históricos. SP3a lo hace **estable**: el fallback de precio para valuar
una caja gratis usa, en orden, (a) el precio de caja de una línea pagada **del mismo orden**,
(b) el promedio de precio de caja **del mismo producto en el mismo mes**, (c) el promedio del
mes. Nunca el pool global → agregar un mes futuro no altera meses previos.

---

## Pieza 3 — Parseo del bloque de incentivos (por estado)

**Cambio en `split_excel.py` (pequeño):** al anexar un bloque de incentivos a un archivo
canónico, anteponer una fila-marcador `["__INCENTIVOS__", <estado>]`, donde `<estado>` es el
estado dominante de la hoja de origen (Florida/Nueva York). Es retro-compatible: no tiene
`datetime` ni SKU, así que los parsers que filtran filas de pedido (GOA, San José) la ignoran;
y no contiene `SOLD`, así que no interfiere con el ancla del parser de incentivos de GOA.

**Parser de incentivos en `generar_data.py`** (defensivo, como el de GOA — el bloque es manual
e irregular): por cada bloque marcado, extrae para su estado:
- **`descuentos`**: lista `[{product, cases, amount}]` a partir de las filas con etiqueta tipo
  `"In Brine == > 3 cases"` + monto (product ∈ {In Brine, Tomate, …} inferido de la etiqueta;
  `cases` del número en la etiqueta; `amount` del monto de la fila).
- **`totalDescuentos`**: suma de los montos de descuento del bloque (el valor que se resta de
  la factura). Es la cifra **autoritativa** para el revenue neto.
- **`freeReportado`** y **`soldReportado`** por vendedor: de las columnas FREE/SOLD → solo
  para cross-check contra lo calculado de las transacciones.
- **`vendedoresClientesNuevos`**: nombres de la lista "NEW CUSTOMERS".

Si el bloque cambia de forma o falta, degradar a listas/valores vacíos (no romper el pipeline).

---

## Pieza 4 — Clientes nuevos (calculados)

Un cliente es **nuevo en el mes M** si su `customer` (normalizado: trim + colapsar espacios +
case-insensitive) **no apareció en ningún mes anterior** dentro del histórico cargado.
Se calcula por mes y por estado, guardando el vendedor de su primera venta. Se expone el
**conteo** y la **lista**; y se compara contra `vendedoresClientesNuevos` del bloque como
verificación (sin sobrescribir el cálculo).

Nota: el "primer mes cargado" no puede evaluar novedad (no hay historia previa); esos clientes
no se cuentan como nuevos (mismo criterio que el `PROD_FIRST_MONTH` actual).

---

## `data.js` enriquecido (contrato para SP3b, aditivo)

Se **conservan** todas las claves actuales (`MONTHS, MONTHS_BTN, PAIRS, PLBLS, PERIOD_TEXT,
PROD_FIRST_MONTH, D, GENERAL_DATA`) para no romper el `index.html` de hoy; sus **valores
numéricos cambian** por el arreglo de la regla. Se **agregan**:

- En cada vista de `GENERAL_DATA[period]`:
  - `summary`: `revNeto` (= rev − totalDescuentos), `descuentos` (total), `clientesNuevos` (conteo).
  - `states[estado]`: `descuentos`, `revNeto`.
  - `descuentosPorProducto`: `[{product, cases, amount, state}]`.
  - `clientesNuevosLista`: `[{name, state, salesperson}]`.
  - `crossCheck`: `{ free: {computado, reportado}, sold: {computado, reportado} }` por estado.
- En `D` (serie temporal por estado): agregar arreglos `descuentos` y `revNeto` por mes.

*(La forma exacta de presentación la afina SP3b; SP3a garantiza que estos datos existan y sean
correctos.)*

---

## Verificación

SP3a **invalida a propósito** el test de regresión de SP2 (`test_regresion_plomeria_historico_solo`
comparaba contra los números viejos, ahora corregidos). Manejo:

1. **Re-baseline:** regenerar la referencia (`reference_pre_sp2.json` → renombrar a
   `reference_pre_sp3.json` con los números **corregidos**) y adaptar el test de regresión para
   guardar contra la **nueva** base (evita cambios accidentales futuros).
2. **Tests dirigidos del arreglo** (los que documentan el cambio deliberado):
   - Orientación por mes: un mes swap (May) se lee correcto; un mes normal, también.
   - Caja gratis `Total=0 / UnitPrice≠0` → detectada como promo (antes contaba como venta).
   - Venta con descuento in-line (qty1, `Total<UnitPrice`) → revenue = `Total` (no `max`).
   - `pv` estable: agregar un mes NO cambia el `pv` de meses previos (la deuda de SP2 queda saldada).
3. **Descuentos → neto:** `revNeto == revBruto − totalDescuentos` por estado y total.
4. **Clientes nuevos:** en un histórico sintético, un cliente que aparece por primera vez en M
   cuenta como nuevo en M y no antes; los del primer mes no cuentan.
5. **Cross-check:** `free.computado` (de líneas) vs `free.reportado` (del bloque) se exponen; el
   test valida que ambos se calculan (no que sean iguales — el reporte manual puede diferir).
6. **Aceptación:** correr `python split_excel.py`; San José regenera; revisar que el revenue
   bruto baja levemente y las cajas promo suben (por las cajas gratis antes mal contadas), que
   aparece el total de descuentos de junio ($628.68) y el revenue neto, y los clientes nuevos.

---

## Decisiones ya tomadas

- Regla de montos: **orientación por mes** (no `max` fila-a-fila, no fecha hardcodeada).
- Cajas gratis: `line_total == 0` con entrega > 0 (captura las de `Total=0 / UnitPrice≠0`).
- Valuación de promos **estable** (fallback orden → producto-mes → mes; nunca pool global).
- Descuentos = autoritativos del bloque; **Revenue neto = bruto − descuentos**, por estado/producto.
- SOLD/FREE del bloque = cross-check; clientes nuevos = **calculados** de las transacciones.
- Estado del bloque = fila-marcador `__INCENTIVOS__` que agrega `split_excel.py`.
- `data.js` aditivo (no rompe el `index.html` actual); rediseño = SP3b.

---

## Fuera de alcance (recordatorio)

- **SP3b:** rediseño del dashboard de San José (visual) para mostrar revenue neto, descuentos
  por producto/estado, clientes nuevos y cross-checks.
- Dashboards KOM/ROB; **SP4** (repo GetUp + subdominio privado).
