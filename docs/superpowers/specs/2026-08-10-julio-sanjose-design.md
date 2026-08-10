# Actualización a julio 2026 — Fase 1: plomería compartida + San José

**Fecha:** 2026-08-10
**Alcance:** `split_excel.py` (compartido, 4 marcas) + `reporte-sanjose/`
**Fases siguientes (fuera de este spec):** GOA, Kombuchacha, Robinson Crusoe

## Contexto

El distribuidor (LatinFood) manda un Excel mensual cuyo layout cambia sin aviso. El de
julio (`07. GET Jul 2026.xlsx`) trae cuatro cambios respecto al de junio:

| | Junio | Julio |
|---|---|---|
| Columna `Company` | en las 5 hojas | **ausente en las 4 hojas NY** |
| Valores de `Company` (MIA) | `LatinFood Florida` | `Florida - LF` / `Florida- LatinFood FL Inc.` |
| Nombres de hoja | `NY Jun - KOM`, `- ROB` | `NY Jul - Kombuchacha`, `- Salmon y Mejillones` |
| Hojas | 5 | 7 (agregan MIA GOA y MIA Kombuchacha) |

El parsing header-aware introducido el 2026-07-10 absorbe el reordenamiento de columnas:
`split_excel.py` reparte julio sin errores y sin SKUs huérfanos. Lo que **no** absorbe es
la pérdida del dato de estado.

### El problema

`reporte-sanjose/generar_data.py` deriva el estado (Florida / Nueva York) únicamente de la
columna `Company`, y **descarta en silencio** la fila que no mapea
([generar_data.py:120-123](../../../reporte-sanjose/generar_data.py)). Verificado contra el
archivo canónico de julio: las 16 filas dan estado `None`. San José saldría en **$0** con un
warning fácil de perder entre el resto del output.

El estado sí viene en el Excel, pero ahora solo en el nombre de la hoja (`NY Jul - …` /
`MIA Jul - …`). `split_excel.py` ya lo sabe deducir así — pero escribe en `fuentes/` el
`Company` crudo, no el resuelto.

### Qué no está roto (verificado, no requiere cambios)

- **Orientación de montos**: `_detect_orientation` la deduce por mes; julio da `normal`.
- **Filas con `qty 0`**: ya se descartan en `generar_data.py:124` (caso `S45942`).
- **SKUs nuevos** (`GOA10`, `GOA11`, `ROB04`): los dashboards de GOA y ROB no tienen listas
  de SKU hardcodeadas; entran solos en sus fases.
- **Estado en las otras marcas**: GOA, KOM y ROB no usan estado. `COL_COMPANY` está definido
  en `report_core.py` pero nunca se lee, y ninguno de sus tres dashboards menciona estados.
  El impacto visible del arreglo es exclusivo de San José.

## Decisiones tomadas

### D1 — La normalización del estado vive en `split_excel.py`

`fuentes/` pasa a garantizar **valores** canónicos, no solo **orden** de columnas.
Los cuatro generadores quedan intactos: cero riesgo de regresión sobre los nueve meses
históricos de San José, y un solo lugar que entiende los caprichos del distribuidor.

Alternativa descartada: arreglarlo en `generar_data.py`. El nombre de la hoja no viaja al
archivo canónico, así que San José no tendría de dónde deducir el estado sin agregar otro
canal de datos.

### D2 — Una hoja sin estado resoluble aborta la corrida

Falla temprano, en el punto que sí tiene contexto (archivo + hoja), y antes de escribir
nada. Como el estado sale del nombre de la hoja, un fallo nunca es "una fila rara": es una
hoja entera. O falla todo o no falla nada, así que abortar no bloquea el mes por un caso
aislado.

Alternativa descartada: procesar igual y marcarlo en el dashboard. Publicar un número
incompleto con una advertencia encima es exactamente el modo de fallo que este spec existe
para eliminar.

### D3 — Un mes sin ventas se registra como cero explícito

Hoy `split_excel.py` salta la marca sin filas: no escribe archivo, no regenera nada. Así
**"vendió $0" y "no cargué el mes" se ven idénticos**. Kombuchacha en julio dice literalmente
`No sales reported` en ambos estados; sin este cambio su dashboard seguiría diciendo
"Mayo – Junio 2026" sin explicación.

Se implementa en `split_excel.py` en esta fase (escribe el archivo); el render del mes en
cero es trabajo de la fase de Kombuchacha.

### D4 — El cross-check de `SOLD` acepta dos bases

El `SOLD` que reporta LatinFood cambió de significado entre meses:

| | Junio | Julio |
|---|---:|---:|
| Cajas entregadas (FL) | 115 | 61 |
| Cajas gratis | 21 | 10 |
| Cajas pagadas | 94 | 51 |
| `SOLD` reportado | **115** = entregadas | **51** = pagadas |
| `FREE` reportado | 21 ✓ | 10 ✓ |

Ninguna lectura es un error de cuadratura: ambas dan un número exacto contra su base. Lo que
cambió es qué cuenta. Con la regla actual (`SOLD` = entregadas, aprendida en SP3b) julio
dispararía una alarma falsa de `51 vs 61`.

El cross-check pasa a validar contra entregadas **o** pagadas y registra cuál calzó. Si un mes
no calza con ninguna, ahí sí hay alarma real.

Alternativa descartada: fijar `SOLD = pagadas` como regla nueva. Apuesta a que el distribuidor
no vuelva a cambiar, y ya cambió de convención dos veces en dos meses.

## Diseño

### Fase 1 (este spec)

Incluye toda la plomería compartida de `split_excel.py` aunque D3 solo la necesite
Kombuchacha: las tres decisiones tocan el mismo archivo, y partirlo en dos pasadas obliga a
reabrirlo y re-testearlo. Se corre con `--only GET` para que GOA/KOM/ROB no se regeneren
todavía.

### Cambios en `split_excel.py`

**C1 — Normalizar `Company` al escribir el archivo canónico.**
`state_of()` ya resuelve el estado con hoja + `Company`. Nuevo: `canonical_row` recibe el
estado resuelto y escribe en la columna `Company` el valor canónico correspondiente
(`LatinFood Florida` para Florida, `LatinFood US Corp.` para Nueva York), descartando lo que
haya mandado el distribuidor. No es inventar dato: la información venía en el nombre de la hoja.

**C2 — Validar antes de escribir.**
`collect()` pasa a registrar las hojas cuyas filas de transacción quedaron en estado `?`.
`main()` valida el resultado completo **antes** del primer `write_canonical`. Si hay alguna,
imprime un error que nombra archivo, hoja y el valor de `Company` que no reconoció, y retorna
código de salida distinto de cero sin escribir nada ni correr generadores.

**C3 — Mes en cero para la marca sin ventas.**
Al recorrer las hojas, registrar qué marcas *aparecen* (por nombre de hoja) aunque no aporten
filas. El mes que se les atribuye es el **mes dominante del archivo** — el más frecuente entre
todas sus transacciones, que es como `split_excel.py` ya resuelve el mes de un bloque de
incentivos. Si el archivo no tuviera ninguna transacción, no se escribe nada: no hay mes que
atribuir. Para cada marca presente-pero-vacía se escribe `<Marca>-<AAAA-MM>.xlsx` solo con el
encabezado.

Si el archivo abarcara más de un mes, la marca vacía se registra solo en el dominante. Es una
simplificación deliberada: el distribuidor manda un archivo por mes, y una marca sin ventas no
tiene datos que permitan distinguir en cuál de los meses estuvo vacía.

El mapeo de nombre de hoja a marca debe cubrir los nombres nuevos: `Kombuchacha` → `KOM`,
`Salmon y Mejillones` → `ROB`, `Jack Mackerel` (y el `Jack Mckarel` de junio) → `GET`,
`GOA` → `GOA`. El match es por subcadena, insensible a mayúsculas y a acentos.

### Cambios en `reporte-sanjose/generar_data.py`

**C4 — El monto del descuento se lee de su columna.**
`_parse_incentive_segment` usa hoy `_last_number(row)`, que barre la fila completa. En julio la
fila `In Brine == > 0 cases` trae la celda de monto vacía, así que agarra el `5` de la columna
`FREE`: descuento de `$5` donde debe ser `$0`, y total de `$305` donde debe ser `$300`.

El monto pasa a buscarse solo desde la columna de montos en adelante (`row[4:]`, o sea después
de la etiqueta de descuento en col D). Sin número en ese rango → `$0`.

**C5 — Cross-check de `SOLD` con dos bases.**
El cross-check compara el `SOLD` reportado contra entregadas y contra pagadas. Pasa si calza
con cualquiera, y `crossCheck` gana un campo que indica cuál (`"entregadas"` / `"pagadas"`).
Falla solo si no calza con ninguna.

**C6 — `index.html` muestra la base usada.**
Único cambio en el dashboard: el texto del cross-check indica contra qué base validó
(p. ej. `SOLD 51 = cajas pagadas`). Sigue sin calcular nada — la etiqueta viene resuelta desde
`generar_data.py`, según la regla de SP3b.

### Fuera de alcance

- `report_core.py`, `goa-inventory/parse_excel.py` y los dashboards de GOA/KOM/ROB.
- El render del mes en cero (fase Kombuchacha).
- Los pendientes de SP4 (Cloudflare Access, GitHub Pages público).
- Rebranding visual de Robinson Crusoe.

## Verificación

1. **Regresión histórica (la prueba que manda):** regenerar San José sin el archivo de julio y
   comparar contra el `data.js` en `main`. Debe salir **idéntico**. Cualquier diferencia en los
   nueve meses de Oct 2025 – Jun 2026 significa que la plomería no es fiel.
2. **Tests nuevos**, siguiendo lo que ya existe en `tests/`:
   - `Company` ausente + hoja `NY …` → se escribe `LatinFood US Corp.`
   - `Company` con basura (`Florida - LF`) + hoja `MIA …` → se escribe `LatinFood Florida`
   - Hoja con transacciones y estado irresoluble → aborta, `fuentes/` sin tocar, rc ≠ 0
   - Hoja de marca presente sin ventas → archivo solo con encabezado
   - Descuento con etiqueta válida y celda de monto vacía → `$0`, no el `FREE`
   - Cross-check de `SOLD` contra entregadas (junio) y contra pagadas (julio), y fallo cuando
     no calza con ninguna
3. **Números de julio** contra la tabla de abajo.
4. **Dashboard en el navegador** (`.claude/launch.json` → `sanjose-dashboard`, puerto 8137),
   sin errores de consola, con julio en el selector de meses.

### Números esperados de julio 2026 (San José)

| Estado | Revenue | Cajas | Gratis | Pedidos | Clientes |
|---|---:|---:|---:|---:|---:|
| Nueva York | $820.80 | 24 | 6 | 3 | 3 |
| Florida | $2,371.20 | 61 | 10 | 3 | 3 |
| **Total** | **$3,192.00** | **85** | **16** | **6** | **6** |

Descuentos (solo Florida): **$300.00** → revenue neto **$2,892.00**.
Orientación de montos detectada: `normal`.
Cross-check Florida: `SOLD 51` calza con **pagadas** (61 entregadas − 10 gratis).
Nueva York no trae bloque de incentivos, igual que en junio.

### Riesgo conocido

Julio trae en MIA GOA un vendedor agrupado como `Distribuidores` (Getessentialshub, 24 cajas,
Luisa Fernanda), un canal distinto a tienda. No afecta a San José y se resuelve en la fase de
GOA, pero queda anotado para no descubrirlo dos veces.
