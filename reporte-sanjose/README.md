# Reporte Ventas - San José

Dashboard de ventas LatinFood · Florida & Nueva York.

## 🔗 Ver el reporte

**[https://cataguzmang.github.io/getup/reporte-sanjose/](https://cataguzmang.github.io/getup/reporte-sanjose/)**

El dashboard incluye KPIs, ingresos mensuales, desempeño por cliente y proveedor, productos por SKU, efectividad de campañas (ROI) y análisis por estado.

## Uso mensual

San José ahora se alimenta del pipeline central: el Excel mensual del distribuidor se
separa por marca con `split_excel.py` (en la raíz), que deja
`fuentes/San-Jose-<AAAA-MM>.xlsx` (formato canónico) y regenera este `data.js`.

1. Desde la raíz del repo, deja el Excel del distribuidor en `entrada/` y corre
   `python split_excel.py` (regenera San José y las demás marcas listas).
2. O, para regenerar solo San José desde sus `fuentes/`: `python generar_data.py`.
3. Sube el `data.js` actualizado y el sitio se actualiza en GitHub Pages.

El histórico Oct 2025 – May 2026 vive migrado en
`fuentes/San-Jose-historico-hasta-2026-05.xlsx` (generado una vez con
`migrar_historico.py`). El Excel `Historico Ventas Latin Food - San Jose.xlsx` original
queda solo como respaldo.

> El estado (Florida / Nueva York) se deriva de la columna `Company`.

## Archivos

| Archivo | Descripción |
| --- | --- |
| `index.html` | Página del dashboard. |
| `data.js` | Datos de ventas (autogenerado desde `fuentes/`, no editar a mano). |
| `generar_data.py` | Convierte los `fuentes/*.xlsx` canónicos en `data.js`. |
| `migrar_historico.py` | Migración única del Historico viejo → `fuentes/` (ya ejecutada). |
| `fuentes/` | Excel canónicos por mes (no se suben al repo). |
| `Historico Ventas Latin Food - San Jose.xlsx` | Fuente vieja, ahora solo respaldo. |

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
