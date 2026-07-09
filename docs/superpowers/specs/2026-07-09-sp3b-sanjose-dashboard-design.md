# SP3b — Rediseño del dashboard de San José

**Fecha:** 2026-07-09
**Estado:** propuesto (pendiente de revisión de la usuaria)
**Sub-proyecto:** 3b (rediseño visual; consume el modelo enriquecido de SP3a)

---

## Contexto y alcance

SP3a dejó el `data.js` de San José enriquecido (revenue bruto/neto, descuentos por
producto/estado, clientes nuevos) pero el `index.html` **no muestra nada de eso todavía**.
SP3b rediseña el dashboard para incorporarlo, validado con mockups aprobados por la usuaria.

**En alcance:**
1. **Pieza de datos (chica):** exponer `crossCheck` en `data.js` (cajas free/sold
   calculadas vs. reportadas por estado). Lo demás ya existe desde SP3a.
2. **Rediseño de `index.html`** (pestañas General y Mes vs. Mes) según los mockups.

**Fuera de alcance:** dashboards de KOM/ROB; migración del repo (SP4). Se **conserva** el
stack actual (un solo HTML + Chart.js por CDN, tema claro azul/cian, formato de números
existente) — es un rediseño de contenido/layout, no un cambio de tecnología.

---

## Principio: todo cálculo derivado se computa en el data layer

**Regla (nueva, aplica de aquí en adelante):** cualquier valor **derivado** (sumas, totales,
razones, porcentajes, ROI, revenue neto, cross-checks, deltas) se calcula en **funciones
explícitas de `generar_data.py`** y se expone **precalculado** en `data.js`. El `index.html`
**solo formatea y renderiza**: mapea valores ya calculados a DOM/charts, usando helpers de
**presentación** explícitos (formato de moneda/número, ancho de barras) — nunca cálculo de
negocio inline y disperso.

Estado hoy: **parcialmente**. Python ya hace las agregaciones y varios derivados; pero el
dashboard aún calcula inline en JS: `% promo`, `ROI`, `revenue/caja`, el total combinado
FL+NY de las series, y `%promo`/`ROI` por mes. SP3b los **mueve a Python** (Pieza 1) y deja
el JS como vista.

---

## Pieza 1 — Cálculos derivados al data layer (Python) + `crossCheck`

Todo lo derivado se calcula con **funciones explícitas** en `generar_data.py` y se expone en
`data.js`. Concretamente SP3b agrega:

**a) `crossCheck` (dato nuevo).** En `build_period_data`, por estado, sobre los meses del periodo:
```
crossCheck[estado] = {
  "free": {"computado": <pc del estado>,        "reportado": <suma freeReportado>},
  "sold": {"computado": <cajas - pc del estado>, "reportado": <suma soldReportado>},
}
```
`freeReportado`/`soldReportado` vienen de `load_incentivos()` (ya se parsean). Los meses
históricos (sin bloque) tienen `reportado = 0` → "sin reporte que cruzar" (no marca descuadre).

**b) Razones que hoy calcula el JS → ahora en Python** (funciones/helpers explícitos, expuestas
en `summary` y `states[estado]`):
- `promoPct` = cajas promo ÷ cajas × 100.
- `roi` = revenue ÷ inversión promos (pv).
- `revPorCaja` = revenue ÷ cajas pagadas.
(Se agregan como campos precalculados; el JS solo los muestra.)

**c) Totales combinados FL+NY de la serie temporal → en Python.** Se agrega `D.total` con
`rev, cajas, pv, pc, ord, revNeto, descuentos` (hoy el JS arma el objeto `A` sumando FL+NY).

**d) Series por mes derivadas → en Python.** En `build_monthly_state_data`, agregar por estado
las series `promoPct` y `roi` por mes (hoy el JS arma `pFL`/`roiFL` inline).

El dashboard queda sin cálculo de negocio: solo formateo (`f$`, `fN`, `fmtNum`), deltas vía el
helper explícito `pct()` (ya existe), y armado de DOM/charts.

---

## Pieza 2 — Rediseño de `index.html`

### Barra superior + cross-check discreto (ambas pestañas)
- Marca + período (como hoy).
- **Indicador de validación discreto** junto al período:
  - Todo cuadra o sin reporte que cruzar → texto tenue "Datos validados" con check muted.
  - Hay descuadre en el periodo visible → aviso ámbar sobrio "N dato(s) por revisar antes
    de publicar" + "Ver detalle" que despliega qué no coincide (estado · métrica ·
    calculado vs. reportado).
  - **Regla del aviso:** solo se evalúa cuando hay `reportado > 0` (o sea, cuando existe un
    bloque de incentivos ese mes). Si `reportado == 0` no se marca nada.

### Pestaña General (orden aprobado)
1. **Titular financiero (Opción B):** flujo `Bruto → −Descuentos → Neto` en tres cajas
   (neto en verde/success, descuentos en danger). Reemplaza el titular de solo-bruto.
2. **Tira de KPIs:** Cajas entregadas · Cajas promo (con %) · Clientes nuevos · Pedidos ·
   Precio/caja. *(Inversión promos y ROI se mueven a las tarjetas por estado.)*
3. **Por Estado (FL / NY):** a lo que ya muestran se agrega **Revenue neto** y **Descuentos**.
4. **Descuentos por producto:** tabla/cards por producto (In Brine / Tomate) y estado, con
   cajas y monto; total del periodo. (De `descuentosPorProducto`.)
5. **Clientes nuevos:** conteo destacado + **lista compacta** (cliente + vendedor que lo
   trajo), colapsable. (De `summary.clientesNuevos` y `clientesNuevosLista`.)
6. **Evolución mensual:** se conservan los gráficos actuales; se **agrega** una serie de
   **Revenue neto** (y opcionalmente descuentos) por mes, aprovechando `D.*.revNeto` /
   `D.*.descuentos`.
7. **Clientes & Equipo, Tabla de Clientes, Productos:** se conservan como están.

### Pestaña Mes vs. Mes (lo nuevo va acá también — es la vista de revisión)
- **Variación Global:** se agregan tarjetas de **Revenue neto**, **Descuentos** y
  **Clientes nuevos** con su delta vs. el mes anterior (descuentos es "inverso": subir = a
  vigilar; y muestra "nuevo" cuando el mes previo no tenía bloque). Se mantienen las
  existentes.
- **Por Estado:** se agregan filas de neto/descuentos a la comparación.
- **Descuentos del mes por producto** y **Clientes nuevos del mes** (lista) para el mes
  actual de la comparación.
- Los gráficos mensuales existentes se conservan; se puede sumar neto/descuentos.

### Selector de período
- El selector actual ("Todos" + cada mes) sigue manejando toda la vista General. Todos los
  bloques nuevos responden al período elegido (los datos ya vienen por periodo en
  `GENERAL_DATA`).

---

## Datos que consume el dashboard (contrato SP3a + Pieza 1)

Por periodo en `GENERAL_DATA[period]`:
- `summary`: `rev`, `revNeto`, `descuentos`, `clientesNuevos`, `promoPct`, `roi`,
  `revPorCaja`, `avg_price`, `cajas`, `pc`, `pv`, `ord`.
- `states[estado]`: `rev`, `revNeto`, `descuentos`, `promoPct`, `roi`, `revPorCaja`, `pc`,
  `pv`, `cajas`, `ord`.
- `crossCheck[estado]`: `{free:{computado,reportado}, sold:{computado,reportado}}`.
- `descuentosPorProducto`: `[{product, state, cases, amount}]`.
- `clientesNuevosLista`: `[{name, state, salesperson}]`.

En `D`: por estado `fl`/`ny` y también `total` (FL+NY), cada uno con las series
`rev, revNeto, descuentos, cajas, pc, pv, ord, promoPct, roi` (una entrada por mes).

---

## Verificación

- **Datos:** tests de que `build_period_data` emite `crossCheck` (junio: free computado 21 ==
  reportado 21) y los derivados `promoPct`/`roi`/`revPorCaja` con los mismos valores que hoy
  calcula el JS (para no cambiar lo que se ve); y de que `D.total` == suma FL+NY y las series
  `D.*.promoPct`/`roi` son correctas.
- **Paridad JS→Python:** confirmar que los campos precalculados igualan lo que el JS mostraba
  (ej. `promoPct` = `pc/cajas*100`), así el rediseño no altera números por mover el cálculo.
- **Regresión:** re-basear `reference_sp3.json` (agrega las claves nuevas); los valores
  existentes no cambian (Pieza 1 es puramente aditiva).
- **Visual/manual:** abrir el `index.html` regenerado en el navegador y confirmar contra los
  mockups aprobados: titular B, KPIs, por-estado con neto/descuentos, descuentos por
  producto, clientes nuevos con lista, cross-check discreto (validado en meses sin descuadre;
  aviso cuando se fuerce uno), y la pestaña Mes vs. Mes con neto/descuentos/clientes nuevos.
  Verificar que carga sin errores de consola y que el selector de período actualiza todo.
- **Sin romper lo existente:** todas las secciones actuales siguen funcionando con los
  números corregidos de SP3a.

---

## Decisiones ya tomadas (de los mockups)

- Titular financiero = **Opción B** (flujo bruto → −descuentos → neto).
- Cross-check **discreto**: solo visible como aviso cuando hay descuadre en el periodo
  (con `reportado > 0`); si todo cuadra, un check tenue "Datos validados".
- Clientes nuevos = conteo + **lista** (cliente + vendedor), colapsable.
- Lo nuevo (neto, descuentos, clientes nuevos) aparece **en ambas pestañas**, con énfasis en
  Mes vs. Mes (vista de revisión).
- Se conserva el stack (HTML único + Chart.js), el tema y el formato de números actuales.
- **Todo cálculo derivado vive en funciones explícitas de `generar_data.py`** (data layer);
  el dashboard solo formatea y renderiza. Los cálculos inline que hoy hace el JS se mueven.

---

## Implementación (nota de enfoque)

- La **Pieza 1** (crossCheck) es un cambio de datos acotado y testeable (TDD).
- El **rediseño del HTML** es un cambio holístico y visual; se implementa directo con
  **preview en el navegador** para iterar contra los mockups, no por TDD. Al final se muestra
  el **dashboard real** para el visto bueno de la usuaria antes de mergear.

## Fuera de alcance (recordatorio)

- Dashboards KOM/ROB; **SP4** (repo GetUp + subdominio privado).
