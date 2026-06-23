# Diseño — Histórico mensual GOA con desglose por mes

**Fecha:** 2026-06-23
**Proyecto:** Dashboard de ventas GOA (Garden of the Andes · LatinFood US Corp)

## Contexto

El distribuidor entrega un export transaccional en Excel (una fila por línea de
pedido). Hasta ahora el dashboard mostraba **un solo mes** (mayo 2026), generado
por `parse_excel.py` desde un archivo fijo (`Latin-GOA-MAY.xlsx`) hacia `data.js`,
que consume `dashboard.html` (HTML + Chart.js, sin build) y se publica en GitHub
Pages.

Llegó un archivo nuevo, **`GOA clientes recurrentes.xlsx`**, con un formato algo
distinto:

- La hoja se llama `Sheet1` (no `GOA`).
- Trae **varios meses juntos** (febrero, marzo y abril 2026), no uno solo.
- Viene limpio: sin filas de subtotal por vendedor y **sin bloque de incentivos**.

Las columnas son las mismas de siempre: `Order Date · Order · Product Variant ·
Customer · Salesperson · Company · Qty Delivered · Qty Invoiced · Qty Ordered ·
Unit Price · Total`.

## Objetivo

1. Ingerir el formato nuevo (multi-mes, hoja `Sheet1`, sin incentivos).
2. **Acumular en un histórico**: cada Excel que vaya llegando se suma al registro,
   sin perder los meses anteriores.
3. Mostrar en el dashboard un **total general** (todos los meses) y poder
   **desglosar por mes** mediante un selector arriba.

## Decisiones (acordadas en brainstorming)

| Tema | Decisión |
|---|---|
| Entregable | Solo el dashboard web (no se genera Excel aparte). |
| Navegación | Selector de mes arriba (`Todos · Febrero · …`) que filtra todo el tablero. |
| Mayo | Se incluye como un mes más, junto a feb-mar-abr. |
| Histórico | Acumulativo **a prueba de pérdidas**: si un mes ya publicado desaparece de las fuentes, se conserva del `data.js` anterior. |
| Privacidad | Los Excel son `*.xlsx` → ya ignorados por git. Solo se publica `data.js` (agregado). |

## Arquitectura

Se mantiene el pipeline actual: **Python agrega → `data.js` → el dashboard solo
pinta.** Los cambios son aditivos.

```
fuentes/*.xlsx  ──►  parse_excel.py  ──►  data.js  ──►  dashboard.html
(local, gitignored)   (combina, dedup,     (general +     (selector de mes
                       histórico, agrega)    months[])      + render(vista))
```

### 1. `parse_excel.py` — ingesta y acumulación

- **Carpeta de fuentes:** nueva carpeta `fuentes/`. El usuario deja ahí **todos**
  los Excel recibidos. El script lee todos los `*.xlsx` de la carpeta e ignora los
  archivos de bloqueo temporales de Excel (`~$...`).
- **Detección de hoja:** por cada archivo, usar `GOA` si existe, si no `Sheet1`,
  si no la primera hoja.
- **Parser reutilizado:** `is_transaction_row` / `parse_transactions` actuales ya
  descartan subtotales de vendedor y el bloque de incentivos, así que **sirven para
  ambos formatos sin cambios**. Cada línea se etiqueta con su mes `YYYY-MM` derivado
  de la fecha (no del nombre del archivo, porque un archivo puede traer varios meses).
- **Dedup por número de orden, entre archivos:** los archivos se procesan en orden
  de nombre; la **primera aparición** de un número de orden gana. Si un mismo número
  de orden aparece en un archivo posterior, se saltan sus líneas y se imprime un
  aviso con los IDs omitidos. (Dentro de un mismo archivo, las líneas repetidas de
  una orden son legítimas y se conservan.)
- **Incentivos por archivo:** el bloque de incentivos se parsea por archivo (con la
  lógica tolerante actual) y se atribuye al/los mes(es) de ese archivo. El archivo
  de mayo aporta sus incentivos a mayo; el archivo nuevo no tiene.
- **Histórico a prueba de pérdidas (carry-forward):**
  1. Cargar el `data.js` anterior si existe (parsear el objeto `reportData`).
  2. Construir los meses desde las fuentes actuales.
  3. Para cada mes que estaba en el `data.js` anterior **y** ya no aparece en la
     reconstrucción, **conservar ese mes tal cual** del `data.js` previo e imprimir
     un aviso claro (`⚠ Mes 2026-03 conservado del histórico anterior: no se encontró
     en fuentes/`).
  4. El total general se recalcula sobre la **unión** (meses reconstruidos +
     conservados).
  - Si el `data.js` anterior es del formato viejo (sin `months`), no hay nada que
    conservar: la reconstrucción lo reemplaza.

### 2. `data.js` — estructura de datos

`reportData` mantiene en el nivel superior la **vista general** (compatibilidad con
el dashboard actual) y agrega un arreglo `months`:

```js
const reportData = {
  meta: {
    brand, distributor,
    periodStart, periodEnd,        // rango global (todos los meses)
    periodLabel,                   // ej. "Febrero – Mayo 2026"
    lineItems,                     // total de líneas tras dedup
    sources: ["...","..."],        // archivos ingeridos
    carriedForward: ["2026-03"],   // meses conservados del histórico (si hubo)
  },
  // ─ Vista general (todos los meses combinados) ─
  totals, products, salespeople, customers, daily, incentives,
  // ─ Desglose por mes (cronológico) ─
  months: [
    {
      key: "2026-02",
      label: "Febrero 2026",
      periodStart, periodEnd,
      totals, products, salespeople, customers, daily, incentives,
    },
    // ...
  ],
};
```

La forma de `totals / products / salespeople / customers / daily / incentives` es
**idéntica** a la actual, tanto en el nivel general como dentro de cada mes. Esto
permite que el dashboard use la misma función de render para cualquier vista.

### 3. `dashboard.html` — selector y render

- **Selector de mes:** una barra de botones arriba (antes de los KPIs), generada
  desde `d.months`: `Todos · Febrero · Marzo · Abril · Mayo …`. Reusa el estilo de
  `.chart-toolbar button` ya existente. `Todos` es el estado inicial (vista general).
- **`render(vista)`:** se refactoriza el bloque imperativo actual en una función que
  recibe una vista (el objeto general `d`, o un elemento de `d.months`) y dibuja
  KPIs, gráficos y tablas a partir de ella. El encabezado (periodo) y los subtítulos
  de los KPIs reflejan la vista seleccionada.
- **Ciclo de vida de Chart.js:** se guardan las instancias de los gráficos y se
  llama `.destroy()` antes de recrearlas en cada cambio de vista. Los contenedores
  de altura fija ya evitan el bug de redimensionado infinito.
- El resto del marcado (tablas de producto/vendedor/cliente, incentivos) no cambia
  de estructura; solo se alimenta desde `vista` en lugar de `d` directamente.

### 4. Documentación

Actualizar la sección "Actualizar datos" del `README.md` al nuevo flujo:

1. Dejar el/los Excel recibidos en la carpeta `fuentes/`.
2. Ejecutar `python parse_excel.py`.
3. `git add data.js && git commit && git push`.
4. El dashboard se actualiza solo en GitHub Pages, con el mes nuevo en el selector.

## Manejo de errores / casos borde

- **Carpeta vacía / sin transacciones:** error claro (`No se detectaron líneas en
  fuentes/. ¿Pusiste los Excel?`), sin escribir `data.js`.
- **Archivo corrupto o sin hoja válida:** se avisa y se salta ese archivo; se sigue
  con los demás.
- **Orden duplicada entre archivos:** se conserva la primera, se avisa de las
  omitidas.
- **Mes desaparecido de fuentes:** se conserva del histórico anterior (carry-forward)
  con aviso; nunca se pierde en silencio.
- **Incentivos ausentes (formato nuevo):** el dashboard ya maneja secciones vacías
  ("Sin datos" / "Sin ítems de costo").

## Validación

- Ejecutar `parse_excel.py` y verificar: total de líneas tras dedup, meses
  detectados, y que **la suma de los meses cuadre con el total general** (unidades,
  ingresos, órdenes).
- Abrir `dashboard.html`, cambiar entre `Todos` y cada mes, confirmar que KPIs,
  gráficos y tablas se redibujan correctamente y que los gráficos no se rompen al
  alternar.
- Verificar el carry-forward: con un mes ya publicado, quitar su Excel de `fuentes/`,
  recalcular y confirmar que el mes se conserva con su aviso.

## Fuera de alcance (YAGNI)

- Generar un Excel/PDF de reporte aparte.
- Comparativas mes contra mes (variación %, tendencias inter-mensuales) más allá del
  desglose simple.
- Cualquier análisis específico de "clientes recurrentes" más allá de las tablas
  actuales (la data nueva es transaccional estándar; el nombre del archivo solo
  refleja su origen).
- Persistir el detalle transaccional en el repo público (seguiría siendo privado en
  los Excel locales; solo se publica el agregado `data.js`).
