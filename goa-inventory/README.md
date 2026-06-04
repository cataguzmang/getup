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

## Actualizar datos

1. Colocar el nuevo Excel del distribuidor como `Latin-GOA-MAY.xlsx`
   (o ajustar `EXCEL_FILE` en `parse_excel.py` si cambia el nombre).
2. Ejecutar: `python parse_excel.py`
3. Subir `data.js` al repo:
   ```
   git add data.js
   git commit -m "data: actualizar reporte GOA"
   git push
   ```
4. El dashboard se actualiza automáticamente en GitHub Pages.

## Requisitos

```
pip install openpyxl
```

## Archivos

| Archivo | Rol |
|---|---|
| `parse_excel.py` | Extrae y agrega el Excel → `data.js` |
| `data.js` | Datos consumidos por el dashboard (`reportData`) |
| `dashboard.html` | Dashboard (HTML + Chart.js, sin build) |
