# Dashboard de Ventas — Garden of the Andes

Dashboard de ventas y distribución para Garden of the Andes · LatinFood US Corp.

## 🔗 [Ver dashboard](https://cataguzmang.github.io/getup/goa-inventory/dashboard.html)

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

## Requisitos

```
pip install openpyxl
```

## Archivos

| Archivo | Rol |
|---|---|
| `fuentes/` | Carpeta local con todos los Excel recibidos (no se sube al repo) |
| `parse_excel.py` | Extrae y agrega el Excel → `data.js` |
| `data.js` | Datos consumidos por el dashboard (`reportData`) |
| `dashboard.html` | Dashboard (HTML + Chart.js, sin build) |
