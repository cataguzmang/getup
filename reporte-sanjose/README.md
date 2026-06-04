# Reporte Ventas - San José

Dashboard de ventas LatinFood · Florida & Nueva York.

## 🔗 Ver el reporte

**[https://cataguzmang.github.io/getup/reporte-sanjose/](https://cataguzmang.github.io/getup/reporte-sanjose/)**

El dashboard incluye KPIs, ingresos mensuales, desempeño por cliente y proveedor, productos por SKU, efectividad de campañas (ROI) y análisis por estado.

## Uso mensual

1. Agregar los datos al Excel (`Historico Ventas Latin Food - San Jose.xlsx`).
2. Ejecutar el generador de datos:
   ```bash
   python generar_data.py
   ```
3. Subir el `data.js` actualizado al repositorio.
4. El sitio se actualiza automáticamente en GitHub Pages.

## Archivos

| Archivo | Descripción |
| --- | --- |
| `index.html` | Página del dashboard. |
| `data.js` | Datos de ventas generados desde el Excel. |
| `generar_data.py` | Script que convierte el Excel en `data.js`. |
| `Historico Ventas Latin Food - San Jose.xlsx` | Fuente de datos. |
