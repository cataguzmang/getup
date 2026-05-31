"""
generar_data.py
Lee Historico Ventas Latin Food - San Jose.xlsx y genera data.js
para que el reporte HTML tenga datos actualizados cada mes.

Uso: python generar_data.py
"""

import openpyxl
import json
from collections import defaultdict
from datetime import datetime

EXCEL_FILE = "Historico Ventas Latin Food - San Jose.xlsx"
OUTPUT_FILE = "data.js"

MES_CORTO = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
             7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
MES_LARGO = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',
             6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',
             10:'Octubre',11:'Noviembre',12:'Diciembre'}


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


def load_rows():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb['Historico']
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        company, state, customer, order, date, cod_prod, product, qty_del, qty_inv, qty_ord, salesperson, total, unit_price = row
        if not date or not state:
            continue
        rows.append({
            'company': company or '',
            'state': state,
            'customer': (customer or '').strip(),
            'order': order or '',
            'date': date if isinstance(date, datetime) else datetime(date.year, date.month, date.day),
            'product': (product or '').strip(),
            'qty': qty_del or 0,
            'total': total or 0,
            'unit_price': unit_price or 0,
            'salesperson': (salesperson or '').strip(),
        })
    return rows


def get_order_unit_prices(rows):
    """Per order: determine the unit_price from paid lines (total > 0)."""
    order_prices = defaultdict(float)
    for r in rows:
        if r['total'] > 0 and r['unit_price'] > 0:
            order_prices[r['order']] = r['unit_price']
    # fallback: average price per product across all paid rows
    product_prices = defaultdict(list)
    for r in rows:
        if r['unit_price'] > 0:
            product_prices[r['product']].append(r['unit_price'])
    product_avg = {p: sum(v)/len(v) for p, v in product_prices.items()}
    return order_prices, product_avg


def build_monthly_state_data(rows, months, order_prices, product_avg):
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
        if r['total'] > 0:
            d['rev'] += r['total']
            d['orders'].add(r['order'])
        elif r['qty'] > 0 and r['unit_price'] == 0:
            # promo box: total=0, unit_price=0, qty>0
            d['pc'] += r['qty']
            price = order_prices.get(r['order'], product_avg.get(r['product'], 0))
            d['pv'] += r['qty'] * price
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


def build_period_data(rows, months, order_prices, product_avg, month_set=None):
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
        if r['total'] > 0:
            d['rev'] += r['total']
            d['orders'].add(r['order'])
        elif r['qty'] > 0 and r['unit_price'] == 0:
            # promo box: total=0, unit_price=0, qty>0
            d['pc'] += r['qty']
            price = order_prices.get(r['order'], product_avg.get(r['product'], 0))
            d['pv'] += r['qty'] * price
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
        if r['total'] > 0:
            cust_data[c]['rev'] += r['total']
        elif r['qty'] > 0:
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
        if r['total'] > 0:
            sales_data[r['salesperson']] += r['total']
    sales_sorted = sorted(sales_data.items(), key=lambda x: -x[1])
    sales_chart = [{'n': n, 'v': round(v, 2)} for n, v in sales_sorted[:10]]

    # --- products aggregation ---
    prod_data = defaultdict(lambda: {'rev': 0.0, 'cajas': 0, 'pc': 0})
    for r in filtered:
        p = r['product']
        prod_data[p]['cajas'] += r['qty']
        if r['total'] > 0:
            prod_data[p]['rev'] += r['total']
        elif r['qty'] > 0:
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
    print(f"Leyendo {EXCEL_FILE}...")
    rows = load_rows()

    # Determine sorted month list
    all_months = sorted(set(month_key(r['date']) for r in rows))
    n = len(all_months)

    order_prices, product_avg = get_order_unit_prices(rows)

    # Build D (time series)
    D = build_monthly_state_data(rows, all_months, order_prices, product_avg)

    # Build GENERAL_DATA
    general_data = {}

    # ALL
    general_data['ALL'] = build_period_data(rows, all_months, order_prices, product_avg,
                                            month_set=set(all_months))

    # Per month
    for ym in all_months:
        label = month_label_short(ym)
        general_data[label] = build_period_data(rows, all_months, order_prices, product_avg,
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
// Fuente: {EXCEL_FILE}
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
