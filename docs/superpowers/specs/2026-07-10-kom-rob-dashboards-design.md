# Dashboards Kombuchacha y Robinson Crusoe — Diseño

**Fecha:** 2026-07-10
**Estado:** aprobado (arquitectura + reglas de degradación validadas con la usuaria)

## Objetivo

Dar de alta los dashboards de **Kombuchacha** (`KOM`) y **Robinson Crusoe** (`ROB`)
como marcas nuevas del pipeline, replicando el modelo de GOA, de modo que:

1. Su `data.js` se **regenere solo** cada mes cuando se corra `split_excel.py`
   (marcándolas `sp1_ready=True`).
2. El dashboard se vea **digno hoy** aunque haya muy pocos datos, y se **enriquezca
   solo** a medida que se acumulen meses (degradación con gracia — Opción B).

## Contexto / realidad de los datos

Son marcas nuevas: su único dato es junio 2026, y es mínimo.

| Marca | Órdenes | SKUs | Clientes | Vendedores | Meses | Estado | Total |
|-------|---------|------|----------|------------|-------|--------|-------|
| Kombuchacha | 1 (S44804) | 2 (Zero Blueberry, Zero Raspberry) | 1 | 1 (Mireya Fernandez) | 1 | NY | $50.64 |
| Robinson Crusoe | 1 (S43486) | 1 (Mussel en Brine/Mejillones) | 1 | 1 (Luis Soler) | 1 | NY | $60.00 |

Ambas son 100% Nueva York (`Company = LatinFood US Corp.`). El Excel de junio de
estas marcas **no** trae bloque de incentivos ni cajas gratis.

## Decisiones de alcance (confirmadas)

- **Sin histórico que migrar.** Se arranca desde junio 2026. No hay `migrar_historico`.
- **Sin bloque de incentivos por ahora.** El core lo soporta (heredado de GOA) pero
  con datos vacíos simplemente no se muestra.
- **Sin desglose por estado.** Igual que GOA (marca de un solo distribuidor). Ambas
  son NY; no se agrega dimensión de estado.
- **No se toca GOA.** `goa-inventory/parse_excel.py` funciona y está testeado; se deja
  intacto. El core compartido lo usan solo KOM y ROB. (Migrar GOA al core queda como
  posible trabajo futuro — YAGNI ahora.)

## Arquitectura

Dos capas, sin duplicar el generador:

```
report_core.py                 (raíz)  ← lógica de cálculo compartida, parametrizable
kombuchacha/build_data.py              ← wrapper delgado: config KOM + escribe data.js
kombuchacha/dashboard.html             ← clon de GOA con degradación
kombuchacha/data.js                    ← generado
robinson-crusoe/build_data.py          ← wrapper delgado: config ROB + escribe data.js
robinson-crusoe/dashboard.html         ← clon de GOA con degradación
robinson-crusoe/data.js                ← generado
split_excel.py                         ← KOM/ROB pasan a sp1_ready=True
tests/test_report_core.py     (raíz)  ← tests del core
kombuchacha/tests/ , robinson-crusoe/tests/  ← test de salida por marca
```

### `report_core.py` (módulo compartido)

Es la generalización de `goa-inventory/parse_excel.py`. Todo su cuerpo ya es genérico
salvo tres constantes que pasan a ser **parámetros de configuración**:

- `brand` (ej. "Kombuchacha")
- `distributor` (ej. "LatinFood US Corp")
- `sku_prefix` (ej. "KOM") → arma el patrón `^\s*\[(KOM\d+)\]\s*(.*)$`
- `sheet_candidates` (ej. `("KOM", "Sheet1")`)
- `here` (carpeta de la marca: define `fuentes/` y `data.js`)

Responsabilidades (todas en funciones explícitas — regla de la usuaria: **todo cálculo
vive en el script de datos, el dashboard solo formatea/renderiza**):

- Leer todos los `.xlsx` de `fuentes/` (ignorando temporales `~$`), dedup de órdenes
  entre archivos, carry-forward de meses ausentes desde el `data.js` previo.
- Agregar por producto (con desglose por vendedor), vendedor, cliente y día.
- Vistas por mes + vista general + metadatos de periodo.
- Emitir el mismo objeto `reportData` que consume el dashboard de GOA
  (`meta`, `totals`, `products`, `salespeople`, `customers`, `daily`, `incentives`,
  `months`).

**Corrección de correctness respecto al clon de GOA — cantidades fraccionarias:**
`parse_excel.py` castea las cantidades con `int(...)` porque GOA siempre vende cajas
enteras. **KOM vende medias cajas (0.5).** Si el core truncara con `int()`, KOM
mostraría 0 unidades. Por lo tanto el core **preserva cantidades como float**
(redondeadas a 2 decimales), no como entero. Esto aplica a `qtyOrdered/Delivered/Invoiced`
y a todos los agregados de unidades/cajas. El dashboard formatea sin decimales cuando
el valor es entero, y con decimales cuando no (ej. "0.5 cajas").

### `build_data.py` por marca (wrapper delgado)

~10-15 líneas. Importa el core (agregando la raíz al `sys.path`), define la config de
la marca, llama a la función de orquestación del core que construye y escribe `data.js`.
Es el nombre que ya espera `split_excel.py` (`generator="build_data.py"`).

### `dashboard.html` por marca (clon de GOA con degradación)

Clon de `goa-inventory/dashboard.html`, con el nombre/título de la marca, apuntando a
su `data.js`, y con la lógica de degradación de la sección siguiente.

### `split_excel.py`

Cambiar `KOM` y `ROB` a `sp1_ready=True` para que el dispatcher regenere su `data.js`
automáticamente en cada corrida mensual.

## Reglas de degradación con gracia (Opción B)

Principio: **un gráfico solo se dibuja si compara algo; una tabla siempre se lee aunque
tenga una fila.** Toda la decisión de mostrar/ocultar vive en el `dashboard.html`
(presentación), a partir de los datos ya calculados por el core.

- **KPIs: siempre.** Unidades, ingresos, órdenes, clientes, SKUs, cajas gratis.
- **Gráficos: solo si hay ≥2 elementos comparables.**
  - Ventas por SKU (ingresos y unidades) → solo si `products.length ≥ 2`.
  - Ranking de vendedores → solo si `salespeople.length ≥ 2`.
  - Tendencia en el periodo → solo si `months.length ≥ 2`.
  - Si no aplica, el bloque del gráfico se **oculta** (no se dibuja una barra/punto solo).
- **Tablas: siempre** (detalle por producto, vendedores, top clientes), aunque tengan
  una fila.
- **Secciones vacías se ocultan por completo:** cajas gratis (si nadie tiene),
  incentivos/muestras (si no hay bloque).
- **Nota de contexto discreta** cuando los datos son escasos (ej. un solo mes o pocas
  líneas): *"Datos iniciales — el reporte se enriquece a medida que se acumulan meses."*

Comportamiento esperado hoy:
- **Kombuchacha:** KPIs + gráfico "Ventas por SKU" (2 SKUs → se dibuja) + tablas de
  detalle/vendedor/cliente. Se ocultan tendencia y ranking de vendedores.
- **Robinson Crusoe:** KPIs + tablas (1 fila c/u). Todos los gráficos ocultos.
- El mes que viene los gráficos aparecen solos al superar los umbrales.

## Estrategia de pruebas

- **`tests/test_report_core.py` (raíz):** con fixtures canónicos en memoria/tmp,
  verifica agregaciones, dedup, carry-forward, y en particular **la preservación de
  cantidades fraccionarias** (una línea de 0.5 → unidades 0.5, no 0).
- **`kombuchacha/tests/` y `robinson-crusoe/tests/`:** corren su `build_data.py` sobre
  su `fuentes/` real y verifican los totales conocidos de junio (KOM: 2 SKUs, rev
  50.64, unidades 1.0; ROB: 1 SKU, rev 60.0, unidades 1) y la forma del `reportData`.
- Regeneración vía `split_excel.py --only KOM,ROB` como verificación de integración
  end-to-end (el dispatcher encuentra el generador y produce `data.js`).

## No-objetivos

- No migrar histórico (no existe).
- No parsear incentivos/descuentos ahora (no hay datos; el core igual lo soporta).
- No desglose por estado.
- No modificar `goa-inventory/` ni `reporte-sanjose/`.
- No la parte de privacidad/subdominio (eso es SP4, aparte).
