# SP1 — `split_excel.py`: dispatcher de limpieza y reparto por marca

**Fecha:** 2026-07-08
**Estado:** propuesto (pendiente de revisión de la usuaria)
**Sub-proyecto:** 1 de 4 (ver "Contexto y alcance")

---

## Contexto y alcance

El distribuidor entrega un único Excel mensual "sucio" (`GET <Mes> <Año>.xlsx`) con
**una hoja por estado + marca** (ej. `NY Jun - GOA`, `MIA Jun - Jack Mckarel`).
Ese archivo alimenta varios reportes, cada uno con su propio dashboard y su propio
`data.js`.

**Este spec cubre solo el SP1: el script de reparto.** Su trabajo es tomar el Excel
sucio y depositar, en la carpeta `fuentes/` de cada marca, un `.xlsx` **canónico y
limpio** con las filas de esa marca — y luego regenerar el `data.js` de las marcas
que ya sepan leer ese formato.

Queda **fuera de alcance** (otros sub-proyectos):

- **SP2** — Refactor de San José a `fuentes/` (parser nuevo + migración del histórico).
- **SP3** — Dashboards nuevos de Kombuchacha (KOM) y Robinson Crusoe (ROB) + núcleo
  compartido `report_core.py`.
- **SP4** — Migración del repo al GitHub de GetUp + subdominio privado.

Este SP1 **no modifica** el parser de GOA ni el dashboard de ningún reporte.

---

## Objetivo

Un comando:

```bash
python split_excel.py
```

que:

1. Lea todos los Excel sucios de `entrada/`.
2. Separe sus filas por **marca** (según el SKU) hacia `fuentes/<Marca>-<YYYY-MM>.xlsx`
   dentro de la carpeta de cada marca.
3. Regenere el `data.js` de cada marca cuyo generador ya entienda el formato canónico
   (en SP1: solo GOA).
4. Imprima un resumen claro de lo que hizo y avise de cualquier cosa que no supo mapear.

Resultado esperable al terminar SP1: **GOA queda actualizado con junio**; San José, KOM
y ROB quedan con su archivo de junio **depositado en `fuentes/`**, esperando su
sub-proyecto.

---

## Formato canónico (el contrato de salida)

El formato canónico **es el mismo export transaccional del distribuidor**, sin la
basura. Un archivo por marca-mes, con **una sola hoja**, columnas idénticas al Excel
original (fila de encabezado + filas de pedido):

```
Order Date | Order | Product Variant | Customer | Salesperson | Company |
Qty Delivered | Qty Invoiced | Qty Ordered | Unit Price | Total
```

Reglas del formato canónico:

- **Passthrough fiel.** Los valores de cada celda se copian **tal cual** (fechas como
  fecha, `0.5` como `0.5`, sin redondear ni castear a int). El script solo **filtra**
  (quita subtotales) y **reparte** (por marca). Cualquier normalización numérica es
  responsabilidad del generador de cada marca (aguas abajo).
- **El estado NO se agrega como columna.** Es derivable de `Company`
  (`LatinFood Florida` → Florida, `LatinFood US Corp.` → Nueva York); el generador de
  San José lo mapeará en SP2. En SP1 el estado se usa solo para el resumen impreso.
- **Nombre de la hoja** = código de marca (para GOA se llama exactamente `GOA`, que es
  lo que su parser prefiere).
- **Bloque de incentivos preservado.** Si una hoja trae el bloque manual del final
  (SOLD / FREE / Descuentos / NEW CUSTOMERS), se copia **verbatim** al archivo canónico
  de esa marca, tras una fila en blanco separadora. No se pierde nada; el parser de GOA
  ya sabe leerlo.

---

## Entrada

- Carpeta `entrada/` en la raíz. Se procesan **todos** los `entrada/*.xlsx` que no
  empiecen por `~$` (temporales de Excel).
- Cada workbook puede traer varias hojas y varias marcas/estados. El script no depende
  del nombre de la hoja para clasificar (ver "Clasificación").

---

## Clasificación de filas

### ¿Es una fila de pedido real?

Misma regla que ya usa el parser de GOA, generalizada:

- Columna A (`Order Date`) es un `datetime`, **y**
- Columna C (`Product Variant`) es un string que matchea el patrón de SKU
  `^\s*\[([A-Z]{2,4}\d+)\]`.

Esto descarta automáticamente:

- Filas-subtotal por vendedor (`Katherine Osorio Duque (6)` — texto en A, sin fecha).
- Filas-subtotal por producto de la hoja de MIA (`[GET01] Jack Mackerel... (8)`,
  `    Alexandra J (2)` — texto en A, no `datetime`).
- Filas en blanco y el bloque de incentivos.

### Marca (por SKU, robusto)

Del SKU embebido en `Product Variant` se extrae el **prefijo alfabético**:

- `[GOA02] Pure Chamoline...` → `GOA`
- `[GET01] Jack Mackerel...` → `GET`
- `[KOM01] Kombuchacha...` → `KOM`
- `[ROB05] Mussel en Brine...` → `ROB`

Mapa de marcas (config del script):

| Prefijo SKU | Marca | `code` | Carpeta destino | Generador | ¿Listo en SP1? |
|---|---|---|---|---|---|
| `GET` | San José | `San-Jose` | `reporte-sanjose` | `generar_data.py` | ❌ (SP2) |
| `GOA` | Garden of the Andes | `GOA` | `goa-inventory` | `parse_excel.py` | ✅ |
| `KOM` | Kombuchacha | `KOM` | `kombuchacha` | `build_data.py` | ❌ (SP3) |
| `ROB` | Robinson Crusoe | `ROB` | `robinson-crusoe` | `build_data.py` | ❌ (SP3) |

- Un prefijo **desconocido** no se descarta: sus filas van a `unmatched/` y se avisa
  en el resumen (nada se pierde en silencio).

### Estado (solo para el resumen)

- Derivado de `Company`: `{"LatinFood Florida": "Florida", "LatinFood US Corp.": "Nueva York"}`.
- Respaldo si `Company` viene vacío: prefijo del nombre de la hoja
  (`{"MIA": "Florida", "NY": "Nueva York"}`).
- Valores no reconocidos se reportan como `?` en el resumen (no bloquea).

### Mes

- `YYYY-MM` de la fecha de cada fila. Una marca puede quedar repartida en varios meses;
  se emite **un archivo canónico por (marca, mes)**.

---

## Algoritmo

1. **Descubrir entrada:** todos los `entrada/*.xlsx` (sin `~$`).
2. **Recolectar:** por cada workbook → hoja → fila:
   - Si es fila de pedido: extraer prefijo SKU → marca; guardar `(marca, fila, mes, estado)`.
   - Prefijo desconocido → lista `unmatched`.
   - Detectar bloque de incentivos de la hoja (ancla: alguna celda == `SOLD`); atribuirlo
     a la marca dominante de esa hoja y a su mes dominante.
3. **Emitir canónicos:** por cada `(marca, mes)` con filas:
   - Crear la carpeta de la marca y su `fuentes/` si no existen.
   - Escribir `<carpeta>/fuentes/<code>-<YYYY-MM>.xlsx`: encabezado + filas de pedido
     (passthrough fiel) + (fila en blanco + bloque de incentivos, si aplica).
   - **Sobrescribe** si el archivo ya existe (idempotente: reprocesar el mismo mes lo
     reemplaza, no duplica).
4. **Emitir no-mapeados:** si hubo filas sin marca, escribir `unmatched/unmatched.xlsx`
   (encabezado + filas) para revisión manual.
5. **Regenerar `data.js`** (salvo `--no-build`): por cada marca con `listo_en_sp1 = True`
   y cuyo generador exista, correrlo (`python <generador>` con cwd = carpeta de la marca)
   vía `subprocess`. Reportar código de salida.
6. **Resumen** (ver abajo).

---

## Interfaz de línea de comandos

```bash
python split_excel.py [--input entrada] [--no-build] [--only GOA,KOM]
```

- `--input` — carpeta de entrada (default `entrada`).
- `--no-build` — solo deposita en `fuentes/`, no regenera ningún `data.js`.
- `--only` — limita a una lista de códigos de marca (para pruebas o updates parciales).

### Resumen impreso (ejemplo)

```
✓ split_excel.py — entrada/GET Jun 2026.xlsx

  GOA  (Garden of the Andes)
    → goa-inventory/fuentes/GOA-2026-06.xlsx   (8 líneas · NY:8)
    ↻ data.js regenerado ✓
  GET  (San José)
    → reporte-sanjose/fuentes/San-Jose-2026-06.xlsx   (3 líneas · NY:1 FL:2 · +incentivos)
    ⏸ data.js NO regenerado (pendiente SP2)
  KOM  (Kombuchacha)
    → kombuchacha/fuentes/KOM-2026-06.xlsx   (2 líneas · NY:2)
    ⏸ data.js NO regenerado (pendiente SP3)
  ROB  (Robinson Crusoe)
    → robinson-crusoe/fuentes/ROB-2026-06.xlsx   (1 línea · NY:1)
    ⏸ data.js NO regenerado (pendiente SP3)

  ⚠ SKUs no mapeados: (ninguno)
```

Si hay filas no mapeadas:
```
  ⚠ 4 líneas sin marca → unmatched/unmatched.xlsx  (prefijos: XYZ, ABC)
```

---

## Privacidad y `.gitignore`

Se agrega un `.gitignore` en la raíz:

```gitignore
# Datos comerciales — nunca subir al repo
*.xlsx
*.xls
entrada/
unmatched/

__pycache__/
*.pyc
.pytest_cache/
```

- `*.xlsx` (sin slash) matchea en **cualquier** subcarpeta → todos los `fuentes/*.xlsx`
  y los Excel sucios quedan fuera del repo, en todas las marcas.
- Los `data.js` **siguen versionándose** (son el artefacto publicado). Que hoy sean
  públicos es un problema conocido que resuelve SP4; SP1 no lo cambia.

---

## Robustez y casos borde

- **Excel sin hojas reconocibles / corrupto:** se avisa y se omite ese archivo, no
  rompe el resto (igual que hace hoy el parser de GOA).
- **Hoja sin filas de pedido:** se ignora (ej. hojas casi vacías).
- **Fila con SKU pero sin fecha, o fecha sin SKU:** no es fila de pedido → se ignora.
- **Cantidades fraccionarias (KOM `0.5`):** se copian tal cual (passthrough). El casteo
  correcto es problema del generador de KOM (SP3).
- **Marca repartida en varias hojas/estados** (ej. Jack Mackerel en `NY` y `MIA`): se
  combinan en un único `San-Jose-<mes>.xlsx`; el estado queda implícito en `Company`.
- **Varios archivos en `entrada/` que tocan el mismo (marca, mes):** se combinan las
  filas de todos antes de emitir (se acumulan, no se pisan entre archivos distintos del
  mismo mes). Reprocesar **el mismo** archivo produce el mismo resultado.
- **Carpeta de marca inexistente** (KOM, ROB en SP1): se crea la carpeta y su `fuentes/`;
  el build se salta con aviso (aún no hay generador).
- **Generador que falla:** se captura el código de salida y se reporta; no aborta las
  demás marcas.

---

## Pruebas (pytest)

Andamiaje espejo del de GOA (`tests/`, `conftest.py`). Los fixtures **construyen un
workbook sintético con openpyxl** (encabezado + filas de pedido de varias marcas/estados +
filas-subtotal + un bloque de incentivos); **no** se commitea data real.

Casos:

1. Clasifica por prefijo de SKU (GET/GOA/KOM/ROB).
2. Filtra subtotales: por vendedor **y** los de producto de MIA (`[GET01]...(8)`,
   `    Alexandra J (2)`).
3. Deriva estado desde `Company`; usa respaldo por prefijo de hoja si `Company` vacío.
4. Agrupa por mes (`YYYY-MM`) y emite un archivo por (marca, mes).
5. Passthrough fiel: `0.5` se conserva; las fechas siguen siendo fechas.
6. Prefijo desconocido → va a `unmatched/`, no se pierde.
7. Bloque de incentivos se copia al canónico de la marca correcta.
8. Idempotencia: reprocesar el mismo mes sobrescribe (no duplica).
9. `--only` y `--no-build` respetados.
10. El archivo GOA emitido es **legible por el parser de GOA existente** (test de
    integración: emitir `GOA-<mes>.xlsx`, correr `parse_excel.py` apuntado a un `fuentes/`
    temporal, y verificar que produce un `data.js` con las líneas esperadas).

---

## Verificación de aceptación (junio 2026)

Con el `GET Jun 2026.xlsx` real en `entrada/`, tras `python split_excel.py`:

- Se crean/actualizan:
  `goa-inventory/fuentes/GOA-2026-06.xlsx` (8 líneas),
  `reporte-sanjose/fuentes/San-Jose-2026-06.xlsx` (3 líneas GET + bloque incentivos),
  `kombuchacha/fuentes/KOM-2026-06.xlsx` (2 líneas),
  `robinson-crusoe/fuentes/ROB-2026-06.xlsx` (1 línea).
- `goa-inventory/data.js` se regenera e incluye junio 2026 en su selector de meses.
- El resumen no reporta SKUs no mapeados.
- No queda ningún `.xlsx` trackeado por git (todos ignorados).

---

## Decisiones ya tomadas

- Nombres de carpeta nuevas: `kombuchacha/`, `robinson-crusoe/`.
- El split **sí** regenera los `data.js` automáticamente (con `--no-build` para saltarlo).
- San José migra al modelo `fuentes/` (en SP2); su histórico se migra una vez.
- El Excel sucio se deja en `entrada/`.
- Robinson Crusoe (`ROB`) es marca propia (no parte de San José) y tendrá dashboard en SP3.
