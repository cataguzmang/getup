# Rebranding visual dashboard Kombuchacha — Plan de Implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reemplazar la piel heredada de GOA en `kombuchacha/dashboard.html` por la identidad real de Kombuchacha (rojo/amarillo/crema, Poppins, color de lata por SKU, logo y miniaturas de lata).

**Architecture:** Solo presentación: assets en `kombuchacha/img/`, reemplazo del bloque `<style>`, header con logo, y mapas de presentación `SKU_COLORS`/`SKU_IMAGES` en el JS del dashboard. Cero cambios en datos, estructura o degradación.

**Tech Stack:** HTML/CSS + Chart.js 4.4 (sin build). Google Fonts Poppins. Imágenes webp/png ya descargadas en el scratchpad.

---

## File Structure

- **Crear** `kombuchacha/img/logo.png`, `kom01.webp`, `kom02.webp`, `kom03.webp`, `kom04.webp` (copiados del scratchpad `komimg/`).
- **Modificar** `kombuchacha/dashboard.html` (head, style, header, JS de colores/miniaturas).

---

## Task 1: Assets de marca

- [ ] **Step 1: Copiar imágenes al repo**

```bash
mkdir -p kombuchacha/img
SCRATCH="C:/Users/caguz/AppData/Local/Temp/claude/C--Users-caguz-proyectos-claude-work-getup-public/f5263440-3946-4ec0-a9a5-502c64e4fef4/scratchpad/komimg"
cp "$SCRATCH/logo.png"      kombuchacha/img/logo.png
cp "$SCRATCH/arandano.webp" kombuchacha/img/kom01.webp
cp "$SCRATCH/frambuesa.webp" kombuchacha/img/kom02.webp
cp "$SCRATCH/jengibre.webp" kombuchacha/img/kom03.webp
cp "$SCRATCH/mate.webp"     kombuchacha/img/kom04.webp
ls -la kombuchacha/img/
```

Expected: 5 archivos (~560 KB total).

- [ ] **Step 2: Commit**

```bash
git add kombuchacha/img/
git commit -m "assets: logo y latas Kombuchacha (línea regular como aproximación de Zero)"
```

---

## Task 2: Head, header y CSS de marca

- [ ] **Step 1: Favicon + Poppins en el `<head>`**

En `kombuchacha/dashboard.html`, reemplazar:

```html
<title>Kombuchacha – Reporte de Ventas</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
```

por:

```html
<title>Kombuchacha – Reporte de Ventas</title>
<link rel="icon" type="image/png" href="img/logo.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
```

- [ ] **Step 2: Header con el logo real**

Reemplazar:

```html
  <div class="header-brand">
    <div class="leaf-icon"><span>🫧</span></div>
    <div class="header-title">
      <h1>Kombuchacha</h1>
      <p>Latin Foods</p>
    </div>
  </div>
```

por:

```html
  <div class="header-brand">
    <img class="brand-logo" src="img/logo.png" alt="Kombuchacha">
    <div class="header-title">
      <h1>Kombuchacha</h1>
      <p>la bebida de todos · Latin Foods</p>
    </div>
  </div>
```

- [ ] **Step 3: Reemplazar TODO el contenido del bloque `<style>…</style>`**

El contenido nuevo completo (mismos selectores y layout, valores de marca; se
elimina `.leaf-icon`, se agregan `.brand-logo` y `.can-thumb`; se conservan
`.charts-grid` y `.sparse-note` re-coloreados):

```css
  :root {
    --rojo:      #E80D0D;
    --amarillo:  #FFD22E;
    --am-palido: #FFF3C9;
    --negro:     #1E1E1E;
    --crema:     #FFF9F0;
    --blanco:    #ffffff;
    --borde:     #F0E4D2;
    --texto:     #1E1E1E;
    --texto-suave: #8A8177;
    --morado:    #8E2E88;
    --rosa:      #E91E7B;
    --turquesa:  #2BBFCB;
    --shadow:    0 2px 12px rgba(232,13,13,.07);
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Poppins', system-ui, sans-serif; background: var(--crema); color: var(--texto); min-height: 100vh; }

  header {
    background: var(--rojo);
    color: #fff;
    padding: 20px 32px 18px;
    display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px;
  }
  .header-brand { display: flex; align-items: center; gap: 14px; }
  .brand-logo { width: 54px; height: 54px; border-radius: 50%; flex-shrink: 0; background: #fff; }
  .header-title h1 { font-size: 1.4rem; font-weight: 800; letter-spacing: .3px; }
  .header-title p  { font-size: .82rem; opacity: .88; margin-top: 2px; font-weight: 600; }
  .header-meta { text-align: right; font-size: .82rem; opacity: .88; line-height: 1.6; }
  .header-meta strong { opacity: 1; font-size: .95rem; }

  main { max-width: 1280px; margin: 0 auto; padding: 28px 24px 48px; }

  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(185px, 1fr));
    gap: 16px; margin-bottom: 28px;
  }
  .kpi-card {
    background: var(--blanco);
    border: 1px solid var(--borde);
    border-radius: 14px;
    padding: 20px 22px;
    box-shadow: var(--shadow);
    display: flex; flex-direction: column; gap: 6px;
    position: relative; overflow: hidden;
  }
  .kpi-card::before {
    content: ''; position: absolute; top: 0; left: 0;
    width: 4px; height: 100%; background: var(--turquesa);
  }
  .kpi-card:nth-child(4n+1)::before { background: var(--rojo); }
  .kpi-card:nth-child(4n+2)::before { background: var(--morado); }
  .kpi-card:nth-child(4n+3)::before { background: var(--turquesa); }
  .kpi-card:nth-child(4n)::before   { background: var(--rosa); }
  .kpi-card.gold { background: var(--am-palido); border-color: var(--amarillo); }
  .kpi-card.gold::before { background: var(--amarillo); }
  .kpi-label { font-size: .72rem; font-weight: 700; color: var(--texto-suave); text-transform: uppercase; letter-spacing: .6px; }
  .kpi-value { font-size: 2.0rem; font-weight: 800; color: var(--negro); line-height: 1.1; }
  .kpi-value.sm { font-size: 1.25rem; }   /* valores de texto (ej. nombre del vendedor) */
  .kpi-card.gold .kpi-value { color: var(--negro); }
  .kpi-sub   { font-size: .80rem; color: var(--texto-suave); }

  .section-title {
    font-size: .72rem; font-weight: 800; color: var(--negro);
    text-transform: uppercase; letter-spacing: 1px; margin-bottom: 14px;
  }

  .month-selector {
    display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 22px;
  }
  .month-selector button {
    padding: 7px 16px; border-radius: 22px; border: 1px solid var(--borde);
    background: var(--blanco); font-size: .82rem; font-weight: 600; cursor: pointer;
    color: var(--texto-suave); transition: all .15s; box-shadow: var(--shadow);
    font-family: inherit;
  }
  .month-selector button:hover { border-color: var(--rojo); color: var(--rojo); }
  .month-selector button.active {
    background: var(--rojo); color: #fff; border-color: var(--rojo);
  }

  .chart-card {
    background: var(--blanco); border: 1px solid var(--borde);
    border-radius: 14px; padding: 22px 24px; box-shadow: var(--shadow);
    margin-bottom: 16px;
  }
  /* Contenedor de altura fija: evita el bucle de redimensionado infinito de
     Chart.js cuando se usa maintainAspectRatio:false. */
  .chart-wrap { position: relative; height: 260px; width: 100%; }
  .chart-wrap.rank { height: 330px; }              /* barras de ranking (más series) */
  @media (max-width: 560px) { .chart-wrap.rank { height: 280px; } }
  .chart-toolbar { display: flex; gap: 6px; margin-bottom: 14px; }
  .chart-toolbar button {
    padding: 3px 12px; border-radius: 20px; border: 1px solid var(--borde);
    background: none; font-size: .74rem; font-weight: 600; cursor: pointer;
    color: var(--texto-suave); transition: all .15s; font-family: inherit;
  }
  .chart-toolbar button.active { background: var(--rojo); color: #fff; border-color: var(--rojo); }

  .table-card {
    background: var(--blanco); border: 1px solid var(--borde);
    border-radius: 14px; box-shadow: var(--shadow);
    overflow: hidden; margin-bottom: 28px;
  }
  .table-card-header {
    padding: 16px 22px; border-bottom: 1px solid var(--borde);
    display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 8px;
  }

  .table-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  table { width: 100%; border-collapse: collapse; font-size: .875rem; min-width: 520px; }
  thead tr { background: var(--am-palido); }
  th {
    padding: 11px 16px; text-align: left;
    font-size: .7rem; font-weight: 700; text-transform: uppercase;
    letter-spacing: .6px; color: var(--negro); white-space: nowrap;
  }
  th.num, td.num { text-align: right; }
  td { padding: 11px 16px; border-bottom: 1px solid var(--borde); vertical-align: middle; }
  tbody tr:last-child td { border-bottom: none; }

  .product-row { cursor: pointer; transition: background .12s; }
  .product-row:hover { background: var(--am-palido); }
  .product-row td:first-child { font-weight: 600; }

  .toggle-btn {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 2px 8px; border-radius: 20px; border: 1px solid var(--borde);
    background: none; font-size: .72rem; cursor: pointer;
    color: var(--texto-suave); transition: background .15s; font-family: inherit;
  }
  .toggle-btn:hover { background: var(--am-palido); }
  .arrow { transition: transform .2s; display: inline-block; }
  .arrow.open { transform: rotate(90deg); }

  .sp-row { display: none; }
  .sp-row.visible { display: table-row; }
  .sp-row td { background: #FFFCF5; font-size: .82rem; color: var(--texto-suave); padding: 9px 16px; }
  .sp-row td:first-child { padding-left: 36px; font-style: italic; }

  .sku-badge {
    display: inline-block; padding: 2px 8px;
    border-radius: 10px; font-size: .68rem; font-weight: 700;
    background: var(--rojo); color: #fff;
    margin-right: 8px;
  }
  .can-thumb { height: 34px; width: auto; vertical-align: middle; margin-right: 8px; }

  .bar-inline { display: flex; align-items: center; gap: 10px; }
  .bar-track {
    flex: 1; height: 7px; background: var(--borde);
    border-radius: 4px; overflow: hidden; min-width: 70px;
  }
  .bar-fill { height: 100%; background: var(--turquesa); border-radius: 4px; }

  .total-row td {
    font-weight: 700; background: var(--am-palido);
    border-top: 2px solid var(--rojo); color: var(--negro);
  }

  .sp-rank { font-weight: 700; color: var(--texto-suave); }
  .pct-badge {
    display: inline-block; padding: 1px 7px; border-radius: 10px;
    font-size: .72rem; font-weight: 600;
    background: var(--am-palido); color: #8A6A00;
  }
  .free-badge {
    display: inline-block; padding: 1px 7px; border-radius: 10px;
    font-size: .72rem; font-weight: 600;
    background: var(--am-palido); color: #8A6A00;
  }

  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 28px; }
  @media (max-width: 760px) { .two-col { grid-template-columns: 1fr; } }

  .incentive-list { list-style: none; }
  .incentive-list li {
    display: flex; justify-content: space-between; align-items: center;
    padding: 10px 0; border-bottom: 1px solid var(--borde); font-size: .87rem;
  }
  .incentive-list li:last-child { border-bottom: none; }
  .incentive-list .amount { font-weight: 700; color: var(--rojo); }
  .incentive-total {
    display: flex; justify-content: space-between; padding: 12px 0 2px;
    margin-top: 6px; border-top: 2px solid var(--rojo);
    font-weight: 700; color: var(--negro);
  }
  .muted-note { font-size: .74rem; color: var(--texto-suave); margin-top: 10px; font-style: italic; }

  /* ── Responsive: tablet y móvil ──────────────────────────────────── */
  @media (max-width: 768px) {
    header { padding: 16px 20px 14px; }
    .header-meta { text-align: left; }
    main { padding: 20px 14px 40px; }
    .kpi-grid { gap: 12px; }
    .kpi-value { font-size: 1.7rem; }
    .chart-card { padding: 18px 16px; }
  }
  @media (max-width: 480px) {
    .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    .kpi-card { padding: 14px 14px; }
    .kpi-value { font-size: 1.5rem; }
    .header-title h1 { font-size: 1.15rem; }
  }

  /* Grilla de gráficos: auto-fit para que los gráficos ocultos no dejen huecos */
  .charts-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
    gap: 16px; margin-bottom: 16px;
  }
  @media (max-width: 400px) { .charts-grid { grid-template-columns: 1fr; } }

  /* Nota discreta cuando hay pocos datos (marca nueva) */
  .sparse-note {
    background: var(--am-palido); color: #8A6A00;
    border: 1px solid var(--amarillo); border-radius: 10px;
    padding: 10px 16px; font-size: .82rem; margin-bottom: 20px;
  }
```

---

## Task 3: JS — colores por sabor, miniaturas y tendencia

- [ ] **Step 1: Paleta de marca + mapas por SKU**

Reemplazar:

```javascript
const palette = ['#2d5a27','#3d7a32','#4a8c3f','#5ea050','#7ab648','#94c96a','#aad98a','#c0e8a8','#d4f0c2'];
```

por:

```javascript
// Paleta de respaldo (marca): para vendedores, clientes y SKUs sin color propio
const palette = ['#E80D0D','#FFD22E','#2BBFCB','#E91E7B','#8E2E88','#F2802E','#7CC576','#B85AA6','#5AB0C0'];

// Identidad por SKU: color e imagen de la lata (presentación; nuevos sabores caen a `palette`)
const SKU_COLORS = { KOM01: '#8E2E88', KOM02: '#E91E7B', KOM03: '#2BBFCB', KOM04: '#2A2A2A' };
const SKU_IMAGES = { KOM01: 'img/kom01.webp', KOM02: 'img/kom02.webp', KOM03: 'img/kom03.webp', KOM04: 'img/kom04.webp' };
const skuColor = (sku, i) => SKU_COLORS[sku] || palette[i % palette.length];
```

- [ ] **Step 2: `horizontalBarChart` acepta color por ítem**

Reemplazar:

```javascript
        backgroundColor: labels.map((_, i) => palette[i % palette.length]),
```

por:

```javascript
        backgroundColor: sorted.map((it, i) => it.color || palette[i % palette.length]),
```

- [ ] **Step 3: Gráficos de SKU con color de lata**

Reemplazar:

```javascript
  // Ventas por SKU — Ingresos
  if (show.revenue) horizontalBarChart('revenueChart',
    view.products.map(p => ({ label: shortName(p.name), value: p.revenue })),
    money0, ctx => ` ${money(ctx.parsed.x)}`);

  // Ventas por SKU — Unidades
  if (show.units) horizontalBarChart('unitsChart',
    view.products.map(p => ({ label: shortName(p.name), value: p.units })),
    qty, ctx => ` ${qty(ctx.parsed.x)} unidades`);
```

por:

```javascript
  // Ventas por SKU — Ingresos (cada barra con el color de su lata)
  if (show.revenue) horizontalBarChart('revenueChart',
    view.products.map((p, i) => ({ label: shortName(p.name), value: p.revenue, color: skuColor(p.sku, i) })),
    money0, ctx => ` ${money(ctx.parsed.x)}`);

  // Ventas por SKU — Unidades (cada barra con el color de su lata)
  if (show.units) horizontalBarChart('unitsChart',
    view.products.map((p, i) => ({ label: shortName(p.name), value: p.units, color: skuColor(p.sku, i) })),
    qty, ctx => ` ${qty(ctx.parsed.x)} unidades`);
```

- [ ] **Step 4: Tendencia en rojo/amarillo**

Reemplazar (dentro del gráfico de tendencia):

```javascript
        borderColor: '#2d5a27', backgroundColor: 'rgba(122,182,72,.18)',
        fill: true, tension: .3, pointRadius: 4, pointBackgroundColor: '#c8a84b', borderWidth: 2,
```

por:

```javascript
        borderColor: '#E80D0D', backgroundColor: 'rgba(255,210,46,.25)',
        fill: true, tension: .3, pointRadius: 4, pointBackgroundColor: '#FFD22E', borderWidth: 2,
```

- [ ] **Step 5: Grillas de ejes y etiquetas de barras en tonos cálidos**

(a) Reemplazar (en `barValueLabel`):

```javascript
    ctx.font = '600 11px "Segoe UI", system-ui, sans-serif';
    ctx.fillStyle = '#1e2e1a';
```

por:

```javascript
    ctx.font = '600 11px "Poppins", system-ui, sans-serif';
    ctx.fillStyle = '#1E1E1E';
```

(b) Reemplazar las DOS ocurrencias de `grid: { color: '#e8f5e0' }` por `grid: { color: '#F3E9DB' }` (una en `horizontalBarChart` escala x, otra en la tendencia escala y).

- [ ] **Step 6: Miniatura de lata + badge coloreado en la tabla de productos**

Reemplazar:

```javascript
    tr.innerHTML = `
      <td><span class="sku-badge">${p.sku}</span>${p.name}</td>
```

por:

```javascript
    tr.innerHTML = `
      <td><span class="sku-badge" style="background:${skuColor(p.sku, idx)}">${p.sku}</span>${SKU_IMAGES[p.sku] ? `<img class="can-thumb" src="${SKU_IMAGES[p.sku]}" alt="">` : ''}${p.name}</td>
```

---

## Task 4: Verificación y commit

- [ ] **Step 1: Sin rastros de la paleta GOA**

```bash
grep -niE "2d5a27|4a8c3f|7ab648|c8a84b|e8f5e0|d4e4cc|green-|gold-|leaf-icon|f4f8f1" kombuchacha/dashboard.html || echo "LIMPIO"
```

Expected: `LIMPIO`.

- [ ] **Step 2: Verificación en navegador (preview puerto 8138)**

- Header: fondo rojo `rgb(232, 13, 13)` (preview_inspect), logo visible.
- Barras de SKU: Blueberry morada, Raspberry rosa, Ginger turquesa, Mate negra
  (leer `Chart.getChart('revenueChart').data.datasets[0].backgroundColor`).
- Miniaturas `.can-thumb` presentes en la tabla (4, una por SKU, `naturalWidth > 0`).
- Tipografía computada de body incluye Poppins.
- Degradación intacta: selector Todos/Mayo/Junio funciona; sin errores de consola.
- Screenshot final para compartir.

- [ ] **Step 3: Commit**

```bash
git add kombuchacha/dashboard.html
git commit -m "feat: rebranding visual Kombuchacha (rojo/amarillo, color por sabor, latas y logo)"
```

---

## Cierre

- superpowers:finishing-a-development-branch (merge a `main` → GitHub Pages).
