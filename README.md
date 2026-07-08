# GetUp — Dashboards

Dashboards operativos de Get Up, publicados vía GitHub Pages.

## Proyectos

| Dashboard | Descripción | Link |
|---|---|---|
| **GOA Inventory** | Inventario y distribución Garden of the Andes · Latinfoods | [Ver →](https://cataguzmang.github.io/getup/goa-inventory/dashboard.html) |
| **Reporte San José** | Ventas LatinFood · Florida & Nueva York | [Ver →](https://cataguzmang.github.io/getup/reporte-sanjose/) |
| **Mapa Cousiño** | Mapa de cuentas y pipeline Get Up NYC | [Ver →](https://cataguzmang.github.io/getup/mapa-cousino/) |

## Cómo actualizar los dashboards

El distribuidor manda un solo Excel mensual (`GET <Mes> <Año>.xlsx`) con una hoja por
estado + marca. El flujo es un comando:

1. Deja ese Excel en `entrada/` (carpeta privada, no se sube al repo).
2. Ejecuta el dispatcher:
   ```
   python split_excel.py
   ```
   - Separa cada marca a su `fuentes/<Marca>-<AAAA-MM>.xlsx` (formato canónico limpio).
   - Regenera el `data.js` de las marcas listas (hoy: GOA).
   - Los SKUs que no reconozca van a `unmatched/` con aviso.
3. Sube los `data.js` cambiados:
   ```
   git add */data.js && git commit -m "data: actualizar reportes" && git push
   ```

Mapeo de marcas (por prefijo de SKU): `GET` → San José · `GOA` → Garden of the Andes ·
`KOM` → Kombuchacha · `ROB` → Robinson Crusoe.

Cada dashboard conserva su propio generador (`fuentes/ → data.js`); `split_excel.py` los
llama al final. Ver el README de cada carpeta para el detalle de cada reporte.

> **Requisitos:** `pip install openpyxl`
