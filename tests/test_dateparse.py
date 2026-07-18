from datetime import date, datetime
from dateparse import parse_fecha


def test_iso():
    assert parse_fecha("2026-07-14") == date(2026, 7, 14)
    assert parse_fecha("2026-07-14 00:00:00") == date(2026, 7, 14)


def test_espanol():
    assert parse_fecha("vie 09-ene-26") == date(2026, 1, 9)
    assert parse_fecha("vie 10-abr-26") == date(2026, 4, 10)
    assert parse_fecha("mar 23-dic-25") == date(2025, 12, 23)


def test_datetime_y_date_pasan_directo():
    assert parse_fecha(datetime(2026, 5, 29, 10, 30)) == date(2026, 5, 29)
    assert parse_fecha(date(2026, 5, 29)) == date(2026, 5, 29)


def test_invalidos():
    assert parse_fecha(None) is None
    assert parse_fecha("") is None
    assert parse_fecha("salidas") is None      # un argumento de carpeta no es fecha
