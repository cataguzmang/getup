import sys
from pathlib import Path

# Hace importable generar_data.py y migrar_historico.py (viven en el padre de tests/)
SANJOSE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SANJOSE))
