# Rebranding visual del dashboard Kombuchacha — Diseño

**Fecha:** 2026-07-10
**Estado:** aprobado (opción "fotos línea regular como aproximación" confirmada)

## Problema

`kombuchacha/dashboard.html` es un clon de GOA y heredó su identidad (verdes,
dorado, hoja 🌿-style). Kombuchacha tiene identidad propia y opuesta:
https://www.kombuchacha.cl/ — fondo blanco, rojo de marca, amarillo de acento,
tipografía bold juguetona, y **un color de lata por sabor**.

## Alcance

- Solo `kombuchacha/dashboard.html` + carpeta nueva `kombuchacha/img/`.
- **Cero cambios** de estructura, datos, lógica de degradación o `data.js`.
  Es piel: CSS, colores de gráficos, logo y miniaturas.
- Robinson Crusoe, GOA y San José no se tocan.

## Identidad (extraída de la web y las latas reales)

Paleta de marca (hex muestreados de los assets descargados):

| Token | Valor | Uso |
|---|---|---|
| `--rojo` | `#E80D0D` | header, acentos primarios, línea de tendencia |
| `--amarillo` | `#FFD22E` | acentos secundarios, KPI destacado, badges |
| `--negro` | `#1E1E1E` | texto principal |
| `--crema` | `#FFF9F0` | fondo de página |
| `--blanco` | `#FFFFFF` | tarjetas |

Colores por sabor (color de la lata de cada SKU):

| SKU | Producto (línea Zero) | Color | Hex |
|---|---|---|---|
| KOM01 | Zero Blueberry | morado | `#8E2E88` |
| KOM02 | Zero Raspberry | rosa fucsia | `#E91E7B` |
| KOM03 | Zero Ginger | turquesa | `#2BBFCB` |
| KOM04 | Zero Mate | negro | `#2A2A2A` |

SKU desconocido (sabor futuro) → color de una paleta de respaldo derivada de la
marca (rojo/amarillo/rosas/turquesas), nunca falla.

## Assets (`kombuchacha/img/`)

Descargados de kombuchacha.cl (marketing público de la marca que GetUp importa):

- `logo.png` (45 KB, 800×800, círculo rojo con lettering negro)
- `kom01.webp` = lata Arándano (morada) · `kom02.webp` = Frambuesa (rosa) ·
  `kom03.webp` = Jengibre (turquesa) · `kom04.webp` = Mate (negra)
  (~130 KB c/u, 1027×1804, fondo transparente)

**Nota asumida:** nuestros SKUs son línea **Zero** (naming de exportación); la web
chilena solo muestra la línea regular. Mismos sabores y colores → se usan como
aproximación. Cuando la marca entregue las fotos Zero, se reemplazan los 4
archivos con el mismo nombre y no se toca código.

## Diseño visual

1. **Header**: banda roja `--rojo` plana (sin degradado), logo circular real
   (`img/logo.png`, ~52px) a la izquierda, `<h1>Kombuchacha</h1>` +
   subtítulo "la bebida de todos · Latin Foods" en blanco. Se elimina el
   `.leaf-icon` con emoji.
2. **Página**: fondo `--crema`; tarjetas blancas con borde suave gris-cálido y
   sombra sutil roja-tibia. Títulos de sección en negro bold; el estilo general
   pasa de "orgánico sereno" (GOA) a "pop limpio".
3. **KPIs**: barra de acento izquierda por tarjeta rotando la paleta de marca
   (rojo, amarillo, turquesa, rosa) en vez de verde/dorado. KPI "Ingresos
   totales" destacado con `--amarillo` (reemplaza al `gold` de GOA).
4. **Gráficos por sabor**: en los gráficos de SKU (ingresos y unidades), cada
   barra usa el color de su lata vía mapa `SKU_COLORS = {KOM01:…, KOM02:…, …}`
   con fallback a paleta de marca por índice. Ranking de vendedores: paleta de
   marca. Tendencia: línea `--rojo` con relleno amarillo translúcido y puntos
   amarillos.
5. **Miniaturas de lata**: en la tabla "Detalle por producto", una `<img>` de
   la lata (~34px de alto, `img/kom0N.webp` vía mapa `SKU_IMAGES`) junto al
   nombre. SKU sin imagen → no se renderiza `<img>` (sin hueco ni icono roto).
6. **Tipografía**: Google Fonts **Poppins** (700/800 para títulos y KPIs, 400/600
   para texto) — sans geométrica bold, el matiz juguetón de la marca sin ser
   caricaturesca. Fallback `system-ui, sans-serif`.
7. **Selector de mes / badges / cross-elementos**: activos en `--rojo`; badges
   `.free-badge`/`.pct-badge` en amarillo pálido con texto oscuro; `.sku-badge`
   con el color del sabor correspondiente (texto blanco, salvo mate).
8. **`<title>`**: "Kombuchacha – Reporte de Ventas" (ya está). Favicon: el logo
   (`<link rel="icon" href="img/logo.png">`).

## Regla del proyecto

Sin cálculos nuevos en el front: los mapas `SKU_COLORS`/`SKU_IMAGES` son
constantes de presentación (estética, no datos derivados) y viven en el
`dashboard.html`, igual que la paleta actual de GOA.

## Verificación

- Dashboard en preview (puerto 8138): header rojo con logo, barras con el color
  correcto por sabor (Blueberry morada, Raspberry rosa, Ginger turquesa, Mate
  negra), miniaturas visibles en la tabla, sin errores de consola.
- La degradación sigue intacta (probar vista "Junio" y "Todos").
- Peso total añadido ≤ ~600 KB (aceptable para GitHub Pages).
- Sin rastros de la paleta GOA (verdes `#2d5a27`/`#4a8c3f`/dorado `#c8a84b`) en
  el archivo.

## No-objetivos

- No rediseñar layout ni añadir/quitar secciones.
- No tocar Robinson Crusoe (queda con el look heredado hasta su propio rebranding).
- No optimizar/recomprimir imágenes (ya son livianas).
