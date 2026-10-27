"""
parse_altagama.py — Sell-through de Alta Gama para Garden of the Andes.

Wrapper del núcleo compartido `altagama_core.py` (raíz del repo): importa solo
los SKU `GA …`, combina el histórico `altagama-historico.json` (ene–jul 2026,
reconstruido desde lo publicado porque el Excel YTD original se perdió) con los
Excel de `Inventarios-GOA-altagama/` y escribe `data-altagama.js`.

Uso:
    python parse_altagama.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # raíz del repo
import altagama_core as core
from altagama_core import *  # noqa: F401,F403 — helpers puros, los usan los tests

HERE = Path(__file__).parent
SOURCE_DIR = core.SOURCE_DIR
HISTORY_FILE = HERE / "altagama-historico.json"
OUTPUT_FILE = HERE / "data-altagama.js"

BRAND = "Garden of the Andes"
SKU_PREFIX = "GA "             # el resto de las filas son de otras marcas


def find_sources(folder=SOURCE_DIR):
    return core.find_sources(folder)


def parse_sheet(rows):
    return core.parse_sheet(rows, SKU_PREFIX, BRAND)


def load_excel(path):
    return core.load_excel(path, SKU_PREFIX, BRAND)


def load_history(path=HISTORY_FILE):
    return core.load_history(path, SKU_PREFIX)


def build_report(sources):
    return core.build_report(sources, BRAND)


if __name__ == "__main__":
    core.run(BRAND, SKU_PREFIX, OUTPUT_FILE, history_file=HISTORY_FILE)
