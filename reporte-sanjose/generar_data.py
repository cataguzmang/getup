"""
generar_data.py
Lee Historico Ventas Latin Food - San Jose.xlsx y genera data.js
para que el reporte HTML tenga datos actualizados cada mes.

Uso: python generar_data.py
"""

import re
import json
from pathlib import Path
from collections import defaultdict
from datetime import datetime

import openpyxl

HERE = Path(__file__).parent
SOURCES_DIR = HERE / "fuentes"
OUTPUT_FILE = HERE / "data.js"

MES_CORTO = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
             7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
MES_LARGO = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',
             6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',
             10:'Octubre',11:'Noviembre',12:'Diciembre'}

# Formato canónico: Order Date · Order · Product Variant · Customer · Salesperson ·
# Company · Qty Delivered · Qty Invoiced · Qty Ordered · Unit Price · Total
SKU_RE = re.compile(r"^\s*\[[A-Z]{2,4}\d+\]\s*")
STATE_BY_COMPANY = {
    "latinfood florida": "Florida",
    "latinfood us corp.": "Nueva York",
    "latinfood us corp": "Nueva York",
}


def month_key(date):
    return (date.year, date.month)

def month_label_short(ym):
    y, m = ym
    return f"{MES_CORTO[m]} {str(y)[2:]}"

def month_label_long(ym):
    y, m = ym
    return f"{MES_LARGO[m]} {y}"

def month_label_btn(ym):
    y, m = ym
    return f"{MES_CORTO[m]} {y}"


def _source_files():
    if not SOURCES_DIR.exists():
        return []
    return sorted(p for p in SOURCES_DIR.glob("*.xlsx") if not p.name.startswith("~$"))


def _state_from_company(company):
    if not isinstance(company, str):
        return None
    return STATE_BY_COMPANY.get(re.sub(r"\s+", " ", company).strip().lower())


def _is_tx(row):
    """Fila de pedido real: datetime en col A y SKU [XXX##] en col C (Product Variant)."""
    return (
        isinstance(row[0], datetime)
        and isinstance(row[2], str)
        and SKU_RE.match(row[2]) is not None
    )


def _product_name(variant):
    """'[GET01] Jack Mackerel in Brine ' -> 'Jack Mackerel in Brine'."""
    return SKU_RE.sub("", str(variant)).strip()


def load_rows():
    """Lee todos los fuentes/*.xlsx (formato canónico) y devuelve las filas de pedido
    limpias. Ignora el bloque de incentivos, subtotales y separadores."""
    rows = []
    for path in _source_files():
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        for row in ws.iter_rows(values_only=True):
            if not _is_tx(row):
                continue
            date, order, variant, customer, salesperson, company = row[:6]
            qty_del, unit_price, total = row[6], row[9], row[10]

            state = _state_from_company(company)
            if not state:
                continue
            qty = qty_del or 0
            if qty <= 0:                       # entregado 0 → no es venta ni promo real
                continue

            # Regla de montos ACTUAL (se conserva en SP2; su arreglo es SP3):
            # el Excel mezcla dos orientaciones (mayo 2026 viene intercambiado),
            # line_total = max(Total, Unit Price).
            a = total or 0
            b = unit_price or 0
            is_promo = (a == 0 and b == 0)
            line_rev = max(a, b)
            box_price = line_rev / qty if qty > 0 else 0.0

            rows.append({
                'company': company or '',
                'state': state,
                'customer': (customer or '').strip(),
                'order': order or '',
                'date': date,
                'product': _product_name(variant),
                'qty': qty,
                'line_rev': line_rev,
                'box_price': box_price,
                'is_promo': is_promo,
                'salesperson': (salesperson or '').strip(),
            })
    return rows


def get_order_unit_prices(rows):
    """Per order: determine the per-box price from paid (non-promo) lines."""
    order_prices = defaultdict(float)
    for r in rows:
        if not r['is_promo'] and r['box_price'] > 0:
            order_prices[r['order']] = r['box_price']
    # fallback: average price per product across all paid rows
    product_prices = defaultdict(list)
    for r in rows:
        if not r['is_promo'] and r['box_price'] > 0:
            product_prices[r['product']].append(r['box_price'])
    product_avg = {p: sum(v)/len(v) for p, v in product_prices.items()}
    # global average as last-resort fallback so promo boxes are never valued at 0
    all_prices = [r['box_price'] for r in rows if not r['is_promo'] and r['box_price'] > 0]
    global_avg = sum(all_prices) / len(all_prices) if all_prices else 0.0
    return order_prices, product_avg, global_avg


def promo_box_price(r, order_prices, product_avg, global_avg):
    """Per-box price used to value a promo (free) box."""
    return order_prices.get(r['order']) or product_avg.get(r['product']) or global_avg


def build_monthly_state_data(rows, months, order_prices, product_avg, global_avg):
    """Build D.fl and D.ny arrays (one value per month per metric)."""
    state_month = defaultdict(lambda: defaultdict(lambda: {
        'rev': 0.0, 'cajas': 0, 'pc': 0, 'pv': 0.0, 'orders': set()
    }))

    for r in rows:
        ym = month_key(r['date'])
        if ym not in months:
            continue
        s = 'fl' if r['state'] == 'Florida' else 'ny'
        d = state_month[s][ym]
        d['cajas'] += r['qty']
        if not r['is_promo']:
            d['rev'] += r['line_rev']
            d['orders'].add(r['order'])
        else:
            # promo box: free (both amount columns are 0)
            d['pc'] += r['qty']
            d['pv'] += r['qty'] * promo_box_price(r, order_prices, product_avg, global_avg)
            d['orders'].add(r['order'])

    result = {}
    for state in ('fl', 'ny'):
        result[state] = {
            'rev': [], 'cajas': [], 'pv': [], 'pc': [], 'ord': []
        }
        for ym in months:
            d = state_month[state][ym]
            result[state]['rev'].append(round(d['rev'], 2))
            result[state]['cajas'].append(d['cajas'])
            result[state]['pv'].append(round(d['pv'], 2))
            result[state]['pc'].append(d['pc'])
            result[state]['ord'].append(len(d['orders']))
    return result


def build_period_data(rows, months, order_prices, product_avg, global_avg, month_set=None):
    """Build summary, states, customerChart, salesChart, customerTable, productsCards for a given month set."""
    if month_set is None:
        month_set = set(months)

    filtered = [r for r in rows if month_key(r['date']) in month_set]

    # --- summary & states ---
    state_data = defaultdict(lambda: {'rev': 0.0, 'cajas': 0, 'pc': 0, 'pv': 0.0, 'orders': set()})
    for r in filtered:
        s = r['state']
        d = state_data[s]
        d['cajas'] += r['qty']
        if not r['is_promo']:
            d['rev'] += r['line_rev']
            d['orders'].add(r['order'])
        else:
            # promo box: free (both amount columns are 0)
            d['pc'] += r['qty']
            d['pv'] += r['qty'] * promo_box_price(r, order_prices, product_avg, global_avg)
            d['orders'].add(r['order'])

    total_rev = sum(d['rev'] for d in state_data.values())
    total_cajas = sum(d['cajas'] for d in state_data.values())
    total_pc = sum(d['pc'] for d in state_data.values())
    total_pv = sum(d['pv'] for d in state_data.values())
    total_ord = len(set(r['order'] for r in filtered if r['order']))
    paid_cajas = total_cajas - total_pc
    avg_price = round(total_rev / paid_cajas, 2) if paid_cajas > 0 else 0

    summary = {
        'rev': round(total_rev, 2),
        'cajas': total_cajas,
        'pc': total_pc,
        'pv': round(total_pv, 2),
        'ord': total_ord,
        'avg_price': avg_price
    }

    states = {}
    for state_name, d in state_data.items():
        states[state_name] = {
            'rev': round(d['rev'], 2),
            'cajas': d['cajas'],
            'pc': d['pc'],
            'pv': round(d['pv'], 2),
            'ord': len(d['orders'])
        }
    if 'Florida' not in states:
        states['Florida'] = {'rev': 0.0, 'cajas': 0, 'pc': 0, 'pv': 0.0, 'ord': 0}
    if 'Nueva York' not in states:
        states['Nueva York'] = {'rev': 0.0, 'cajas': 0, 'pc': 0, 'pv': 0.0, 'ord': 0}

    # --- customer aggregation ---
    cust_data = defaultdict(lambda: {'rev': 0.0, 'cajas': 0, 'pc': 0, 'state': ''})
    for r in filtered:
        c = r['customer']
        cust_data[c]['state'] = r['state']
        cust_data[c]['cajas'] += r['qty']
        if not r['is_promo']:
            cust_data[c]['rev'] += r['line_rev']
        else:
            cust_data[c]['pc'] += r['qty']

    custs_sorted = sorted(cust_data.items(), key=lambda x: -x[1]['rev'])

    customer_chart = []
    for name, d in custs_sorted[:10]:
        s = 'fl' if d['state'] == 'Florida' else 'ny'
        customer_chart.append({'n': name, 'v': round(d['rev'], 2), 's': s})

    customer_table = []
    for name, d in custs_sorted:
        total_c = d['cajas']
        promo_c = d['pc']
        promo_pct = round(promo_c / total_c * 100, 1) if total_c > 0 else 0.0
        customer_table.append({
            'name': name,
            'state': d['state'],
            'revenue': round(d['rev'], 2),
            'cajas': total_c,
            'promo_qty': promo_c,
            'promo_pct': promo_pct
        })

    # --- salesperson aggregation ---
    sales_data = defaultdict(float)
    for r in filtered:
        if not r['is_promo']:
            sales_data[r['salesperson']] += r['line_rev']
    sales_sorted = sorted(sales_data.items(), key=lambda x: -x[1])
    sales_chart = [{'n': n, 'v': round(v, 2)} for n, v in sales_sorted[:10]]

    # --- products aggregation ---
    prod_data = defaultdict(lambda: {'rev': 0.0, 'cajas': 0, 'pc': 0})
    for r in filtered:
        p = r['product']
        prod_data[p]['cajas'] += r['qty']
        if not r['is_promo']:
            prod_data[p]['rev'] += r['line_rev']
        else:
            prod_data[p]['pc'] += r['qty']

    products_sorted = sorted(prod_data.items(), key=lambda x: -x[1]['rev'])
    products_cards = []
    for name, d in products_sorted:
        total_c = d['cajas']
        promo_pct = round(d['pc'] / total_c * 100, 1) if total_c > 0 else 0.0
        products_cards.append({
            'name': name,
            'revenue': round(d['rev'], 2),
            'cajas': total_c,
            'promo_pct': promo_pct
        })

    return {
        'summary': summary,
        'states': states,
        'customerChart': customer_chart,
        'salesChart': sales_chart,
        'customerTable': customer_table,
        'productsCards': products_cards
    }


def find_product_first_months(rows, months):
    """Return the first month (label) each product appears in."""
    prod_first = {}
    for ym in months:
        for r in rows:
            if month_key(r['date']) == ym:
                p = r['product']
                if p not in prod_first:
                    prod_first[p] = month_label_short(ym)
    return prod_first


def main():
    print("Leyendo fuentes/...")
    rows = load_rows()

    # Determine sorted month list
    all_months = sorted(set(month_key(r['date']) for r in rows))
    n = len(all_months)

    order_prices, product_avg, global_avg = get_order_unit_prices(rows)

    # Build D (time series)
    D = build_monthly_state_data(rows, all_months, order_prices, product_avg, global_avg)

    # Build GENERAL_DATA
    general_data = {}

    # ALL
    general_data['ALL'] = build_period_data(rows, all_months, order_prices, product_avg, global_avg,
                                            month_set=set(all_months))

    # Per month
    for ym in all_months:
        label = month_label_short(ym)
        general_data[label] = build_period_data(rows, all_months, order_prices, product_avg, global_avg,
                                                month_set={ym})

    # Find first month per product (for dynamic note in product cards)
    prod_first = find_product_first_months(rows, all_months)
    first_month_label = month_label_short(all_months[0])
    prod_first_js = {p: m for p, m in prod_first.items() if m != first_month_label}

    # MONTHS, PAIRS, PLBLS
    months_short = [month_label_short(ym) for ym in all_months]
    months_btn = [month_label_btn(ym) for ym in all_months]

    pairs = [[i, i-1] for i in range(n-1, 0, -1)]
    plbls = [[month_label_long(all_months[i]), month_label_long(all_months[i-1])]
             for i in range(n-1, 0, -1)]

    # Period pill text
    period_text = f"{month_label_btn(all_months[0])} – {month_label_btn(all_months[-1])}"

    # Serialize to JS
    js = f"""// AUTO-GENERADO por generar_data.py — no editar manualmente
// Fuente: fuentes/*.xlsx
// Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}

const MONTHS = {json.dumps(months_short, ensure_ascii=False)};
const MONTHS_BTN = {json.dumps(months_btn, ensure_ascii=False)};
const PAIRS = {json.dumps(pairs)};
const PLBLS = {json.dumps(plbls, ensure_ascii=False)};
const PERIOD_TEXT = {json.dumps(period_text, ensure_ascii=False)};
const PROD_FIRST_MONTH = {json.dumps(prod_first_js, ensure_ascii=False)};
const D = {json.dumps(D, ensure_ascii=False)};
const GENERAL_DATA = {json.dumps(general_data, ensure_ascii=False)};
"""

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(js)

    print(f"✓ {OUTPUT_FILE} generado: {period_text} ({n} meses)")

    # Quick validation
    all_summary = general_data['ALL']['summary']
    print(f"  Revenue total: ${all_summary['rev']:,.2f}")
    print(f"  Cajas: {all_summary['cajas']}")
    print(f"  % Promo: {round(all_summary['pc']/all_summary['cajas']*100,1) if all_summary['cajas'] else 0}%")


if __name__ == '__main__':
    main()
