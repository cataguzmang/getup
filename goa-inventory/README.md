# Dashboard de Ventas — Garden of the Andes

Dashboard de ventas y distribución para Garden of the Andes.

## 🔗 [Ver dashboard](https://cataguzmang.github.io/getup/goa-inventory/dashboard.html)

## Dos distribuidores, dos secciones

El dashboard tiene una pestaña por distribuidor. **No se suman ni se cruzan**:
entregan información con distinto nivel de detalle y sus métricas no son
equivalentes.

| Pestaña | Distribuidor | Fuente | Generador | Datos | Métricas |
|---|---|---|---|---|---|
| Latin Foods | LatinFood US Corp | `fuentes/*.xlsx` | `parse_excel.py` → `data.js` | Export transaccional | Unidades, **ingresos**, órdenes, clientes, vendedores, incentivos |
| Alta Gama | Alta Gama | `Inventarios-GOA-altagama/GOA_CM Sell Through Data_*.xlsx` | `parse_altagama.py` → `data-altagama.js` | Unidades vendidas por producto y mes | **Solo unidades** |

Cada pestaña lee su propio archivo de datos y lleva su propio filtro de periodo;
un filtro nunca altera la otra sección. Si falta `data-altagama.js`, las
pestañas se ocultan y el dashboard funciona igual con Latin Foods.

## Formato del Excel (definitivo)

El distribuidor entrega un **export transaccional** (hoja `GOA`): una fila por
línea de pedido con columnas `Order Date · Order · Product Variant · Customer ·
Salesperson · Company · Qty Delivered/Invoiced/Ordered · Unit Price · Total`.
Al final hay un bloque manual de incentivos (cajas gratis, muestras).

El script `parse_excel.py` lee ese archivo, lo limpia y agrega, y genera
`data.js` con: productos, vendedores, clientes, serie temporal e incentivos —
incluyendo **ingresos ($)**, no solo unidades.

El distribuidor puede entregar un mes suelto (hoja `GOA`) o un export con varios
meses (hoja `Sheet1`); el script detecta la hoja y separa por mes según la fecha de
cada línea. El bloque de incentivos es opcional.

## Actualizar datos

1. Deja el/los Excel que recibas en la carpeta `fuentes/` (no se suben al repo;
   los `*.xlsx` están en `.gitignore` por privacidad). Un mismo archivo puede traer
   varios meses; se separan solos.
2. Ejecuta: `python parse_excel.py`
   - Lee **todos** los Excel de `fuentes/`, descarta órdenes duplicadas y reconstruye
     el histórico completo (todos los meses) en `data.js`.
   - Si un mes que ya estaba publicado ya no tiene Excel en la carpeta, se conserva
     del `data.js` anterior (sale un aviso).
3. Sube `data.js` al repo:
   ```
   git add data.js
   git commit -m "data: actualizar reporte GOA"
   git push
   ```
4. El dashboard se actualiza automáticamente en GitHub Pages, con el mes nuevo
   disponible en el selector de arriba.

## Actualizar los datos de Alta Gama

Alta Gama entrega un único Excel acumulado del año (`GOA_CM Sell Through
Data_<AÑO>_YTD.xlsx`): una fila por producto y una columna por periodo con las
**unidades vendidas**. No trae precios, clientes, órdenes ni inventario.

1. Deja el Excel nuevo en `Inventarios-GOA-altagama/` (carpeta privada, entera
   en `.gitignore`). Reemplaza al anterior: es acumulado, no incremental.
2. Ejecuta: `python parse_altagama.py`
   - Reconstruye `data-altagama.js` completo desde ese único archivo (por eso
     no puede haber duplicados).
   - Recalcula los totales desde las filas de producto y los compara con la
     fila `TOTAL` del Excel; imprime ✓/✗ por periodo. **Si sale ✗, no publiques:
     revisa el Excel primero.**
   - La última columna puede ser un periodo parcial (`Jul 1 - 25, 26`); se
     detecta solo y se marca como tal en todo el dashboard.
3. Sube `data-altagama.js` al repo (el Excel no se sube).

### Los ceros son reales

Alta Gama escribe `0` cuando **no hubo movimiento** en el periodo; es un dato
informado, no un hueco, y se cuenta como cero. Solo una celda **en blanco** se
trata como dato ausente (`n/d`) y queda fuera de las sumas — hoy no hay ninguna.
El dashboard atenúa los ceros para distinguirlos a simple vista de un `n/d`.

### Los CSV semanales están discontinuados

Alta Gama dejó de enviar los snapshots `inventory_*.csv` en **mayo de 2026**
(el último es `inventory_2026-05-11.csv`). Los 10 archivos se conservan en la
carpeta como histórico interno, pero **no alimentan el dashboard y no habrá
más**. Aunque llegaran, sus columnas de venta son ventanas móviles de
30/60/90/180 días, **no** meses calendario: no se suman ni se cruzan con las
unidades del Excel.

### Qué NO se muestra para Alta Gama, y por qué

Ingresos, precios, márgenes, valor de inventario, inventario disponible y
sell-through **porcentual**: la fuente no los trae y no hay denominador fiable
para deducirlos. Tampoco se cruza nada con Latin Foods — ni siquiera los SKU,
que no están mapeados entre distribuidores a propósito.

## Requisitos

```
pip install openpyxl
```

## Archivos

| Archivo | Rol |
|---|---|
| `fuentes/` | Carpeta local con todos los Excel de Latin Foods (no se sube al repo) |
| `parse_excel.py` | Extrae y agrega el Excel de Latin Foods → `data.js` |
| `data.js` | Datos de Latin Foods consumidos por el dashboard (`reportData`) |
| `Inventarios-GOA-altagama/` | Carpeta local con los archivos de Alta Gama (no se sube al repo) |
| `parse_altagama.py` | Extrae y valida el sell-through de Alta Gama → `data-altagama.js` |
| `data-altagama.js` | Datos de Alta Gama consumidos por el dashboard (`altaGamaData`) |
| `dashboard.html` | Dashboard con las dos pestañas (HTML + Chart.js, sin build) |
