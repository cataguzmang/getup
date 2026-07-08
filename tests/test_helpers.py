import datetime

from helpers import tx, subtotal
import split_excel as sx

D = datetime.datetime


def test_sku_prefix():
    assert sx.sku_prefix("[GOA02] Pure Chamoline Herbal Tea") == "GOA"
    assert sx.sku_prefix("[GET01] Jack Mackerel in Brine") == "GET"
    assert sx.sku_prefix("  [ROB05] Mussel en Brine") == "ROB"
    assert sx.sku_prefix("Katherine Osorio Duque (6)") is None
    assert sx.sku_prefix(None) is None


def test_is_transaction_row():
    good = tx(D(2026, 6, 24), "S45227", "[GOA02] x", "cliente", "sp")
    assert sx.is_transaction_row(good) is True
    # subtotal por vendedor
    assert sx.is_transaction_row(subtotal("Katherine (6)")) is False
    # subtotal por producto de MIA (texto con SKU en col A, sin fecha)
    assert sx.is_transaction_row(subtotal("[GET01] Jack Mackerel (8)")) is False
    # encabezado del bloque de incentivos
    inc = [D(2026, 5, 1), "SOLD", "FREE", "Descuentos", None, None, None, None, None, None, None]
    assert sx.is_transaction_row(inc) is False


def test_month_key():
    assert sx.month_key(D(2026, 6, 24, 11, 4)) == "2026-06"
    assert sx.month_key(D(2026, 12, 1)) == "2026-12"


def test_state_of():
    fl = tx(D(2026, 6, 1), "S1", "[GET01] x", "c", "sp", company="LatinFood Florida")
    ny = tx(D(2026, 6, 1), "S2", "[GET01] x", "c", "sp", company="LatinFood US Corp.")
    assert sx.state_of(fl, "MIA Jun - Jack Mckarel") == "Florida"
    assert sx.state_of(ny, "NY Jun - GOA") == "Nueva York"
    # Company vacío → respaldo por prefijo de hoja
    blank = tx(D(2026, 6, 1), "S3", "[GET01] x", "c", "sp", company=None)
    assert sx.state_of(blank, "MIA Jun - Jack Mckarel") == "Florida"
    # desconocido
    unk = tx(D(2026, 6, 1), "S4", "[GET01] x", "c", "sp", company="Otra Cosa")
    assert sx.state_of(unk, "?? Jun") == "?"
