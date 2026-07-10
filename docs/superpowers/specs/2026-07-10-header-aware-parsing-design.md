# Parsing por nombre de encabezado (header-aware) — Diseño

**Fecha:** 2026-07-10
**Estado:** aprobado (alcance "ambos archivos" confirmado con la usuaria)

## Problema

Los lectores de Excel del pipeline mapean columnas por **posición fija** (índices
0-10 del formato canónico). El distribuidor "manda los datos como quiere": el
2026-07-10 llegó `kombuchacha/fuentes/Sales Kombuchacha.xlsx` con una columna
**Address insertada** en la posición 4 y **sin Company**:

```
Canónico:  Order Date | Order | Product Variant | Customer | Salesperson | Company | Qty...
Nuevo:     Order Date | Order | Product Variant | Customer | Address     | Salesperson | Qty...
```

Resultado real observado: `report_core` leyó las **direcciones como vendedores**
("2 Orchard St, Morristown, NJ…" apareció en el ranking de vendedores). Los demás
campos calzaron por coincidencia. Corrupción **silenciosa** — el generador terminó
"✓" sin error.

## Objetivo

Que los dos puntos de entrada de Excel ubiquen cada campo lógico **por el nombre
de su encabezado**, tolerando columnas insertadas, quitadas o reordenadas, con
fallback posicional cuando no hay encabezado, y **avisando** (no fallando ni
corrompiendo) cuando falta una columna.

## Alcance

- **`report_core.py`** — lee `fuentes/*.xlsx` de KOM/ROB (donde pegó el bug).
- **`split_excel.py`** — lee el Excel mensual del distribuidor en `entrada/`
  (la puerta principal; blindarlo protege aguas abajo a GOA y San José, que
  consumen el canónico que él escribe).
- **NO se tocan** `goa-inventory/parse_excel.py` ni `reporte-sanjose/generar_data.py`:
  leen archivos canónicos producidos por `split_excel.py` (o históricos ya
  migrados), que siempre traen el layout canónico. Migrarlos al helper queda
  como mejora futura opcional.

## Diseño

### Helper compartido: `column_resolver.py` (raíz)

Módulo nuevo y chico, importado por `report_core.py` y `split_excel.py`.

**Campos lógicos** (claves internas) y sus nombres de encabezado:

| Campo lógico | Encabezado (normalizado) |
|---|---|
| `date` | `order date` |
| `order` | `order` |
| `variant` | `product variant` |
| `customer` | `customer` |
| `salesperson` | `salesperson` |
| `company` | `company` |
| `qty_delivered` | `qty delivered` |
| `qty_invoiced` | `qty invoiced` |
| `qty_ordered` | `qty ordered` |
| `unit_price` | `unit price` |
| `total` | `total` |

**API:**

```python
CANONICAL_POSITIONS = {"date": 0, "order": 1, "variant": 2, "customer": 3,
                       "salesperson": 4, "company": 5, "qty_delivered": 6,
                       "qty_invoiced": 7, "qty_ordered": 8, "unit_price": 9,
                       "total": 10}

def find_header_row(rows, max_scan=10) -> int | None
    # Índice de la primera fila (entre las primeras `max_scan`) que contenga,
    # normalizados, al menos "order date" y "product variant" (las dos columnas
    # imprescindibles). None si no aparece.

def resolve_columns(header_row, label="") -> dict[str, int]
    # {campo_lógico: índice} para cada encabezado reconocido en la fila.
    # Los campos NO encontrados no aparecen en el dict; por cada faltante se
    # imprime "⚠ {label}: columna 'X' no encontrada; ..." una sola vez.
    # Normalización: str(celda).strip().lower() con espacios colapsados.

def get(row, cols, field, default=None)
    # row[cols[field]] con tolerancia: campo ausente en cols o índice fuera
    # de rango -> default.
```

**Semántica de resolución (regla de seguridad):**

1. `find_header_row` encuentra encabezado → se confía **solo** en los nombres
   (`resolve_columns`). Columnas extra (Address) se ignoran; columnas ausentes
   (Company) → warning + `get(...)` devuelve default.
2. No hay encabezado → `CANONICAL_POSITIONS` completo (comportamiento actual,
   compatibilidad con hojas sin header si existieran).

Como el formato canónico mapea 1:1 con `CANONICAL_POSITIONS`, los archivos
existentes producen **exactamente el mismo resultado** → regresión limpia.

### Cambios en `report_core.py`

- `read_sheet_rows` no cambia (sigue devolviendo filas crudas).
- En `load_all_sources`, tras leer las filas de cada archivo: localizar el
  encabezado y resolver columnas **por archivo** (cada `.xlsx` puede venir con
  layout distinto), con `label=path.name` para que el warning diga qué archivo.
- `is_transaction_row` y `parse_transactions` reciben el mapa de columnas y usan
  `get(row, cols, campo)` en lugar de índices `COL_*` fijos.
- `parse_incentives` no cambia (el bloque manual de incentivos no es tabular;
  su ancla 'SOLD' y heurísticas actuales siguen igual).
- Filas más cortas que el índice resuelto → `get` devuelve default (hoy una fila
  corta rompería con IndexError; de regalo queda tolerante).

### Cambios en `split_excel.py`

- En `collect`, por cada **hoja** de cada workbook: localizar encabezado y
  resolver columnas (`label=f"{archivo}/{hoja}"`).
- `is_transaction_row`, `state_of` y la construcción de la fila canónica de
  salida usan el mapa. La fila canónica de salida se **rearma campo por campo**
  en el orden canónico (`HEADER` de 11 columnas), no copiando la fila cruda:
  así el canónico escrito en `fuentes/` siempre queda bien aunque la entrada
  venga desordenada. `Company` ausente → celda vacía y el estado cae al respaldo
  por prefijo de hoja (NY/MIA), que ya existe.
- `find_incentive_block` no cambia (busca el ancla 'SOLD' en cualquier celda).

## Caso de validación real (end-to-end)

Con `kombuchacha/fuentes/Sales Kombuchacha.xlsx` presente (formato con Address):

- `python kombuchacha/build_data.py` debe producir vendedores = personas
  (Mireya Fernandez, Katherine Osorio Duque, Luisa Fernanda), **ninguna
  dirección** en el ranking.
- Totales esperados (verificados a mano del archivo): 2 meses (2026-05,
  2026-06), dedup de S44804, 22 líneas, ingresos $641.00, unidades 12.5.
- Un warning visible por la columna `Company` ausente en ese archivo.

## Estrategia de pruebas

- **`tests/test_column_resolver.py`** (nuevo): find_header_row (posición 0, en
  medio, ausente), resolve_columns (canónico exacto; con Address insertada; con
  Company faltante → warning y clave ausente), get (campo presente, ausente,
  fila corta).
- **`tests/test_report_core.py`** (ampliar): un `.xlsx` con el layout nuevo
  (Address, sin Company) → salespeople correctos; mezcla de un archivo canónico
  + uno con layout nuevo en la misma carpeta → agregados coherentes.
- **`tests/` de split_excel** (ampliar): workbook de entrada con columnas
  reordenadas/extra → el canónico escrito en `fuentes/` queda en orden canónico
  con los valores correctos; hoja sin Company → estado por prefijo de hoja.
- **Regresión:** toda la suite existente (72 tests) debe pasar sin cambios —
  los formatos canónicos existentes no cambian de resultado. En particular
  `reporte-sanjose/tests/test_regresion.py` y los datos de GOA quedan idénticos.

## No-objetivos

- No migrar `goa-inventory/parse_excel.py` ni `reporte-sanjose/generar_data.py`
  al helper (leen canónico garantizado).
- No inferencia difusa de encabezados (sinónimos, typos): matching exacto
  normalizado. Si el distribuidor renombra una columna, el warning lo dirá y
  se decide ahí.
- No cambiar el formato canónico de salida ni el objeto `reportData`.
