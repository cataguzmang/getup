"""Genera kombuchacha/data-altagama.js (sell-through de Alta Gama) con el núcleo compartido.

Alta Gama vende Kombuchacha desde agosto de 2026 (factura 1086, 31-jul). No hay
histórico previo: todo sale de los Excel de la carpeta compartida de Alta Gama.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # raíz del repo
import altagama_core

if __name__ == "__main__":
    altagama_core.run(
        brand="Kombuchacha",
        sku_prefix="Kombuchacha ",
        output_file=Path(__file__).parent / "data-altagama.js",
    )
