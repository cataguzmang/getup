"""Genera kombuchacha/data.js desde kombuchacha/fuentes/ usando el core compartido."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # raíz del repo
import report_core

if __name__ == "__main__":
    report_core.build_and_write(
        here=Path(__file__).parent,
        brand="Kombuchacha",
        distributor="LatinFood US Corp",
        sku_prefix="KOM",
        sheet_candidates=("KOM", "Sheet1"),
    )
