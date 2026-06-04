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
| `data.js` | Datos de ventas generados desde el Excel (autogenerado, no editar a mano). |
| `generar_data.py` | Script que convierte el Excel en `data.js`. |
| `Historico Ventas Latin Food - San Jose.xlsx` | Fuente de datos. |

## Cómo se calculan las métricas

- **Revenue**: suma del total de cada línea entregada.
- **Cajas**: cantidad entregada (`Qty Delivered`). Las líneas con cantidad entregada = 0 (pedidos no entregados) se ignoran.
- **Cajas promo**: líneas regaladas (las dos columnas de importe vienen en 0).
- **Inversión promos**: cajas promo valoradas al precio real por caja de su pedido.
- **ROI promo**: `Revenue ÷ Inversión promos`.
- **% Vol. en promo**: cajas promo ÷ cajas totales.

> **Nota sobre el Excel:** el archivo trae dos formatos de columnas mezclados.
> Hasta abril 2026, `Total` es el total de la línea y `Unit Price` el precio por caja;
> desde mayo 2026 esas dos columnas vienen **intercambiadas**. El script lo resuelve
> automáticamente tomando `total_de_línea = max(Total, Unit Price)` y
> `precio_por_caja = total_de_línea ÷ cantidad`, así que funciona con ambos formatos.

## Verificar los datos

Después de generar, revisa que los totales tengan sentido (por ejemplo, que la inversión
en promos sea un porcentaje razonable del revenue, no mayor que él). El script imprime un
resumen al terminar:

```
✓ data.js generado: Oct 2025 – May 2026 (8 meses)
  Revenue total: $26,429.76
  Cajas: 756
  % Promo: 23.9%
```
