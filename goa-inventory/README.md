# Dashboard Inventario — Garden of the Andes

Dashboard de reporte de inventario y distribución para Garden of the Andes · Latinfoods.

## 🔗 [Ver dashboard](https://cataguzmang.github.io/getup/goa-inventory/dashboard.html)

## Actualizar datos

1. Completar el Excel con los datos nuevos
2. Ejecutar: `python parse_excel.py`
3. Subir `data.js` al repo:
   ```
   git add data.js
   git commit -m "data: actualizar reporte GOA"
   git push
   ```
4. El dashboard se actualiza automáticamente en GitHub Pages
