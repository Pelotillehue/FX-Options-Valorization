import pytest
from load_market import completar_fila_curva

SPOT = 890.68   # caso real 2026-05-29


def test_datos_completos_pasan_intactos():
    clp, usd = completar_fila_curva(SPOT, 34, 4.4983, 4.5317, 4.515,
                                    890.449999999999, 4.78965647904396)
    assert clp == 4.515 and usd == 4.78965647904396


def test_mid_clp_desde_bid_ask():
    clp, _ = completar_fila_curva(SPOT, 1, 4.4, 4.6, None, 890.67, 4.9)
    assert clp == pytest.approx(4.5)


def test_mid_usd_derivado_de_fwd_spot_clp():
    # 1d con Mid USD vacio: debe reproducir la tasa implicita del archivo real
    _, usd = completar_fila_curva(SPOT, 1, 4.5, 4.5, None, 890.67, None)
    assert usd == pytest.approx(4.904240627841, abs=1e-9)


def test_mid_usd_derivado_otro_plazo():
    # el "plazo faltante" puede ser cualquiera, no solo 1d
    _, usd = completar_fila_curva(SPOT, 65, None, None, 4.535, 890.38, None)
    assert usd == pytest.approx(4.72313803786444, abs=1e-9)


def test_no_derivable_queda_none():
    clp, usd = completar_fila_curva(SPOT, 2194, None, None, 5.045, None, None)
    assert clp == 5.045 and usd is None      # sin FWD no hay tasa USD


# ---------- asimetria con el grid real (se salta sin data/) ----------
import os
GRID = os.path.join(os.path.dirname(__file__), "data",
                    "Grid Vol Bloomberg Creasys 2026-07-14.xlsx")

def test_curvas_asimetricas_grid_real():
    from load_market import load_market
    md = load_market(GRID)
    assert len(md.clp_dias) > len(md.usd_dias)           # CLP llega mas lejos
    assert md.clp_dias[-1] > 7000                        # ~20Y
    assert md.usd_dias[-1] < 2000                        # ~5Y
    # tasa CLP a 10Y usa nodos largos reales (no extrapolacion plana de 5Y)
    assert md.rate_clp_std(3659) == pytest.approx(5.325, abs=1e-9)
