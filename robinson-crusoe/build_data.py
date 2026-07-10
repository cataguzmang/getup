"""Genera robinson-crusoe/data.js desde robinson-crusoe/fuentes/ usando el core compartido."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # raíz del repo
import report_core

if __name__ == "__main__":
    report_core.build_and_write(
        here=Path(__file__).parent,
        brand="Robinson Crusoe",
        distributor="LatinFood US Corp",
        sku_prefix="ROB",
        sheet_candidates=("ROB", "Sheet1"),
    )
