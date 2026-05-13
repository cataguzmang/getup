import re
import json
import openpyxl
from pathlib import Path

EXCEL_FILE = Path(__file__).parent / "Garden of the Andes_Latinfood_report2026.xlsx"
OUTPUT_FILE = Path(__file__).parent / "data.js"

SKU_PATTERN = re.compile(r"^\[GOA\d+\]")

def val(cell):
    v = cell.value
    return int(v) if isinstance(v, (int, float)) and not (v != v) else 0

def parse():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows())

    # Row 0 → date in col A, channel headers in cols B-D
    # Row 1 → "Qty Ordered" labels
    date_cell = rows[0][0].value
    report_date = str(date_cell) if date_cell else "2026-05-08"

    # Extract channel names from row 0, cols B(1), C(2), D(3)
    channels = [
        str(rows[0][1].value or "Vendors"),
        str(rows[0][2].value or "Team NY"),
        str(rows[0][3].value or "Inside Sales NY"),
    ]

    products = []
    current_product = None

    for row in rows[2:]:  # skip header rows
        label = row[0].value
        if label is None:
            continue

        label = str(label).strip()

        if SKU_PATTERN.match(label):
            # Save previous product
            if current_product:
                products.append(current_product)

            # Parse SKU and name: "[GOA01] Green Tea Herbal Tea"
            bracket_end = label.index("]")
            sku = label[1:bracket_end]
            name = label[bracket_end + 1:].strip()

            current_product = {
                "id": sku,
                "sku": sku,
                "name": name,
                "vendors": val(row[1]),
                "teamNY": val(row[2]),
                "insideSalesNY": val(row[3]),
                "total": val(row[4]),
                "salespeople": [],
            }

        elif current_product and label and not label.lower().startswith("total"):
            # Salesperson row under current product
            sp_total = val(row[1]) + val(row[2]) + val(row[3])
            if sp_total > 0 or val(row[4]) > 0:
                current_product["salespeople"].append({
                    "name": label,
                    "vendors": val(row[1]),
                    "teamNY": val(row[2]),
                    "insideSalesNY": val(row[3]),
                    "total": val(row[4]) if val(row[4]) > 0 else sp_total,
                })

    if current_product:
        products.append(current_product)

    totals = {
        "vendors": sum(p["vendors"] for p in products),
        "teamNY": sum(p["teamNY"] for p in products),
        "insideSalesNY": sum(p["insideSalesNY"] for p in products),
        "grand": sum(p["total"] for p in products),
    }

    report_data = {
        "meta": {
            "date": report_date,
            "brand": "Garden of the Andes",
            "distributor": "Latinfoods",
        },
        "channels": channels,
        "products": products,
        "totals": totals,
    }

    js_content = "const reportData = " + json.dumps(report_data, indent=2, ensure_ascii=False) + ";\n"
    OUTPUT_FILE.write_text(js_content, encoding="utf-8")
    print(f"✓ data.js generado — {len(products)} productos, {totals['grand']} unidades totales")

if __name__ == "__main__":
    parse()
