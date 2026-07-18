"""Regresion de datos de mercado con el grid real del 2026-07-14 (fixture en el
repo). Los valores esperados provienen del maestro Excel validado al centavo."""
import os
from datetime import date
import pytest
from load_market import load_market

GRID = os.path.join(os.path.dirname(__file__), "data",
                    "Grid Vol Bloomberg Creasys 2026-07-14.xlsx")


@pytest.fixture(scope="module")
def md():
    return load_market(GRID, fecha_proceso=date(2026, 7, 14))


def test_spot(md):
    assert md.spot == pytest.approx(926.78)


def test_nodos_curvas(md):
    assert len(md.clp_dias) == 23 and md.clp_dias[-1] == 7311   # CLP hasta 20Y
    assert len(md.usd_dias) == 17 and md.usd_dias[-1] == 1832   # USD hasta 5Y


def test_forward_contra_maestro(md):
    # OPT100192: vencimiento 2026-07-22 -> forward cacheado en el maestro
    assert md.forward_at(date(2026, 7, 22)) == pytest.approx(926.712662262186, abs=1e-9)


def test_vol_contra_maestro(md):
    # OPT100192: 8 dias, strike 930 -> vol cacheada en el maestro
    assert md.surface.get_vol(8, 930) == pytest.approx(12.90101412099385, abs=1e-9)


def test_strikes_superficie_contra_maestro(md):
    # nodo 90 dias de 'Volatilidades' (V:Z cacheado): 10P,25P,ATM,25C,10C
    k = md.vol_tenor_dias.index(90)
    esperados = [859.83, 891.50, 928.29, 972.87, 1026.46]
    for j, e in enumerate(esperados):
        assert md.surface.strikes[j][k] == pytest.approx(e, abs=0.01)
