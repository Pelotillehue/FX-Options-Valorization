from datetime import date
import math
import pytest
from fx_options_valuation import (MarketData, Operacion, lin_interpol,
                                  prima_call, prima_put, mtm_option_fx,
                                  liquidar_vencimiento, dias_vencimiento)


def md_plano(spot=900.0, r_clp=5.0, r_usd=5.0, vol=10.0):
    """Mercado sintetico: curvas y superficie planas."""
    dias = [1, 30, 90, 180, 365]
    return MarketData(fecha_proceso=date(2026, 1, 1), spot=spot,
                      raw_dias=dias, raw_clp=[r_clp] * 5, raw_usd=[r_usd] * 5,
                      vol_tenor_dias=dias,
                      vol_smile=[(vol, vol, vol, vol, vol)] * 5).build()


# ---------- interpolacion ----------
def test_lin_interpol_interior():
    assert lin_interpol([0, 10], [0.0, 100.0], 5) == pytest.approx(50.0)

def test_lin_interpol_extrapola_plano():
    assert lin_interpol([10, 20], [1.0, 2.0], 5) == 1.0     # bajo el rango
    assert lin_interpol([10, 20], [1.0, 2.0], 99) == 2.0    # sobre el rango


# ---------- forward ----------
def test_forward_igual_spot_con_tasas_iguales():
    md = md_plano(r_clp=5.0, r_usd=5.0)
    fwd = md.forward_at(date(2026, 7, 1))
    assert fwd == pytest.approx(900.0, abs=1e-9)

def test_forward_manual():
    md = md_plano(spot=900.0, r_clp=6.0, r_usd=4.0)
    dias = dias_vencimiento(md.fecha_proceso, date(2026, 6, 30))  # 180
    df_clp = 1 / (1 + 0.06 * dias / 360)
    df_usd = 1 / (1 + 0.04 * dias / 360)
    assert md.forward_at(date(2026, 6, 30)) == pytest.approx(900.0 * df_usd / df_clp)


# ---------- superficie de vol ----------
def test_vol_plana_devuelve_vol_plana():
    md = md_plano(vol=12.5)
    assert md.surface.get_vol(45, 900) == pytest.approx(12.5)
    assert md.surface.get_vol(45, 800) == pytest.approx(12.5)   # cualquier strike

def test_strikes_de_superficie_ordenados():
    md = md_plano()
    k = md.vol_tenor_dias.index(90)
    strikes = [md.surface.strikes[j][k] for j in range(5)]
    assert strikes == sorted(strikes)          # 10P < 25P < ATM < 25C < 10C


# ---------- Black-Scholes ----------
def test_paridad_put_call():
    md = md_plano()
    dias = 90
    fwd = 905.0; k = 900.0; vol = 12.0
    df = md.df_clp_at(dias)
    c = prima_call(fwd, k, vol, df, dias)
    p = prima_put(fwd, k, vol, df, dias)
    assert c - p == pytest.approx(df * (fwd - k), rel=1e-12)

def test_signos_mtm():
    args = dict(fx_fwd=910.0, strike=900.0, vol=12.0, df=0.99, plazo_dias=90)
    compra = mtm_option_fx(1000, tipo="CALL", lado="Compra", **args)
    venta = mtm_option_fx(1000, tipo="CALL", lado="Venta", **args)
    assert compra > 0 and venta == pytest.approx(-compra)


# ---------- liquidacion ----------
def _op(tipo, lado, strike=900.0, nominal=1000.0):
    return Operacion("X", date(2026, 1, 1), date(2026, 2, 1), tipo, lado,
                     nominal, strike)

def test_call_itm():
    liq = liquidar_vencimiento(_op("CALL", "Compra"), fixing_usdobs=920.0)
    assert liq["itm"] and liq["flujo_pago_clp"] == pytest.approx(20.0 * 1000)

def test_call_otm():
    liq = liquidar_vencimiento(_op("CALL", "Compra"), fixing_usdobs=890.0)
    assert not liq["itm"] and liq["flujo_pago_clp"] == 0.0

def test_put_itm_venta_paga():
    liq = liquidar_vencimiento(_op("PUT", "Venta"), fixing_usdobs=880.0)
    assert liq["itm"] and liq["flujo_pago_clp"] == pytest.approx(-20.0 * 1000)
