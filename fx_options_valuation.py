"""
Motor de valorizacion de Opciones FX (USD/CLP) - reimplementacion en Python
del modelo VBA del archivo maestro MtM.xlsm de Nevasa.

Replica exactamente:
  - Interpolacion lineal de curvas (LinInterpol)
  - Construccion de curva de descuento y forwards (hoja Curvas)
  - Superficie de volatilidad: strikes desde delta (Mids BGN / Volatilidades)
  - GetVol  -> vol interpolada por plazo y por strike (smile)
  - GetFxForward -> forward USDCLP por paridad de tasas
  - MtM_Option_FX -> Black-Scholes (Garman-Kohlhagen, forward) x nominal

Convenciones tomadas del VBA original:
  - Descuento Act/360 ; tiempo de la opcion Act/365
  - Delta con ajuste por prima: -NORMSINV(delta*exp(r_usd*t))
"""

from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime, date
from math import log, sqrt, exp
from statistics import NormalDist

_N = NormalDist()


# --------------------------------------------------------------------------
# Utilidades numericas (equivalentes VBA)
# --------------------------------------------------------------------------
def norm_s_dist(x: float) -> float:
    return _N.cdf(x)


def normsinv(p: float) -> float:
    return _N.inv_cdf(p)


def lin_interpol(plazos, tasas, plazo):
    """LinInterpol del VBA: plano fuera de rango, lineal en el interior."""
    xs = list(plazos); ys = list(tasas)
    n = len(xs)
    if n == 0 or n != len(ys):
        return None
    if plazo >= xs[-1]:
        return ys[-1]
    if plazo <= xs[0]:
        return ys[0]
    x1, y1, x2, y2 = xs[0], ys[0], xs[1], ys[1]
    for i in range(n - 1):
        if plazo > xs[i]:
            x1, y1 = xs[i], ys[i]
            x2, y2 = xs[i + 1], ys[i + 1]
    return y1 + (y2 - y1) / (x2 - x1) * (plazo - x1)


def dias_vencimiento(fecha_proceso, fecha_vencimiento) -> int:
    return (_as_date(fecha_vencimiento) - _as_date(fecha_proceso)).days


def _as_date(d):
    if isinstance(d, datetime):
        return d.date()
    if isinstance(d, date):
        return d
    if isinstance(d, str):
        return datetime.fromisoformat(d[:19]).date()
    raise TypeError(f"fecha no reconocida: {d!r}")


# --------------------------------------------------------------------------
# Datos de mercado del dia
# --------------------------------------------------------------------------
STD_TENORS_DAYS = [1, 7, 14, 21, 30, 60, 90, 120, 180, 270, 365,
                   540, 730, 1095, 1460, 1825]
STD_TENOR_LABELS = ["1D", "1W", "2W", "3W", "1M", "2M", "3M", "4M", "6M",
                    "9M", "1Y", "18M", "2Y", "3Y", "4Y", "5Y"]


@dataclass
class MarketData:
    fecha_proceso: date
    spot: float
    raw_dias: list      # nodos reales de la curva (dias)
    raw_clp: list       # tasa CLP % en cada nodo
    raw_usd: list       # tasa USD implicita % en cada nodo
    vol_tenor_dias: list
    vol_smile: list     # lista de (v10P, v25P, vATM, v25C, v10C) en %

    surface: "VolSurface" = None

    def build(self):
        self.surface = VolSurface.from_market(self)
        return self

    def rate_clp_std(self, dias):
        return lin_interpol(self.raw_dias, self.raw_clp, dias)

    def rate_usd_std(self, dias):
        return lin_interpol(self.raw_dias, self.raw_usd, dias)

    def forward_at(self, fecha_venc):
        dias = dias_vencimiento(self.fecha_proceso, fecha_venc)
        r_clp = lin_interpol(self.raw_dias, self.raw_clp, dias)
        r_usd = lin_interpol(self.raw_dias, self.raw_usd, dias)
        df_clp = 1.0 / (1.0 + (r_clp / 100.0) * dias / 360.0)
        df_usd = 1.0 / (1.0 + (r_usd / 100.0) * dias / 360.0)
        return self.spot * df_usd / df_clp

    def df_clp_at(self, dias):
        r = lin_interpol(self.raw_dias, self.raw_clp, dias)
        return 1.0 / (1.0 + (r / 100.0) * dias / 360.0)


# --------------------------------------------------------------------------
# Superficie de volatilidad (strikes desde delta)
# --------------------------------------------------------------------------
class VolSurface:
    def __init__(self, tenor_dias, vols_by_delta, strikes_by_delta):
        self.tenor_dias = tenor_dias
        self.vols = vols_by_delta
        self.strikes = strikes_by_delta

    @classmethod
    def from_market(cls, md: "MarketData"):
        tenor_dias = list(md.vol_tenor_dias)
        vols = [[], [], [], [], []]
        strikes = [[], [], [], [], []]
        for k, dias in enumerate(tenor_dias):
            v10p, v25p, vatm, v25c, v10c = md.vol_smile[k]
            r_clp = md.rate_clp_std(dias) / 100.0
            r_usd = md.rate_usd_std(dias) / 100.0
            df_clp = 1.0 / (1.0 + r_clp * dias / 360.0)
            df_usd = 1.0 / (1.0 + r_usd * dias / 360.0)
            fwd = md.spot * df_usd / df_clp
            t = dias / 365.0
            j25 = -normsinv(0.25 * exp(r_usd * dias / 365.0))
            k10 = -normsinv(0.10 * exp(r_usd * dias / 365.0))
            s10p = fwd * exp(-k10 * (v10p / 100.0) * sqrt(t) + 0.5 * (v10p / 100.0) ** 2 * t)
            s25p = fwd * exp(-j25 * (v25p / 100.0) * sqrt(t) + 0.5 * (v25p / 100.0) ** 2 * t)
            satm = fwd * exp(0.5 * (vatm / 100.0) ** 2 * t)
            s25c = fwd * exp(j25 * (v25c / 100.0) * sqrt(t) + 0.5 * (v25c / 100.0) ** 2 * t)
            s10c = fwd * exp(k10 * (v10c / 100.0) * sqrt(t) + 0.5 * (v10c / 100.0) ** 2 * t)
            for j, val in enumerate((v10p, v25p, vatm, v25c, v10c)):
                vols[j].append(val)
            for j, val in enumerate((s10p, s25p, satm, s25c, s10c)):
                strikes[j].append(val)
        return cls(tenor_dias, vols, strikes)

    def get_vol(self, dias, strike):
        vols_at_t = [lin_interpol(self.tenor_dias, self.vols[j], dias) for j in range(5)]
        strikes_at_t = [lin_interpol(self.tenor_dias, self.strikes[j], dias) for j in range(5)]
        return lin_interpol(strikes_at_t, vols_at_t, strike)


# --------------------------------------------------------------------------
# Black-Scholes FX (forward)
# --------------------------------------------------------------------------
def _d1(fx_fwd, strike, vol, plazo_dias):
    yf = plazo_dias / 365.0
    return (log(fx_fwd / strike) + 0.5 * yf * (vol / 100.0) ** 2) / ((vol / 100.0) * sqrt(yf))


def _d2(fx_fwd, strike, vol, plazo_dias):
    yf = plazo_dias / 365.0
    return _d1(fx_fwd, strike, vol, plazo_dias) - (vol / 100.0) * sqrt(yf)


def prima_call(fx_fwd, strike, vol, df, plazo_dias):
    d1 = _d1(fx_fwd, strike, vol, plazo_dias)
    d2 = _d2(fx_fwd, strike, vol, plazo_dias)
    return df * (fx_fwd * norm_s_dist(d1) - strike * norm_s_dist(d2))


def prima_put(fx_fwd, strike, vol, df, plazo_dias):
    d1 = _d1(fx_fwd, strike, vol, plazo_dias)
    d2 = _d2(fx_fwd, strike, vol, plazo_dias)
    return df * (strike * norm_s_dist(-d2) - fx_fwd * norm_s_dist(-d1))


def mtm_option_fx(nominal, fx_fwd, strike, vol, df, plazo_dias, tipo, lado):
    tipo = (tipo or "").strip().upper()
    lado = (lado or "").strip().capitalize()
    if tipo == "CALL":
        p = prima_call(fx_fwd, strike, vol, df, plazo_dias)
    elif tipo == "PUT":
        p = prima_put(fx_fwd, strike, vol, df, plazo_dias)
    else:
        return 0.0
    if lado == "Venta":
        p = -p
    elif lado != "Compra":
        return 0.0
    return nominal * p


# --------------------------------------------------------------------------
# Operacion
# --------------------------------------------------------------------------
@dataclass
class Operacion:
    folio: str
    fecha_inicio: object
    fecha_vencimiento: object
    tipo: str
    lado: str
    nominal: float
    strike: float
    prima_clp: float = 0.0

    def valorizar(self, md: MarketData):
        dias = dias_vencimiento(md.fecha_proceso, self.fecha_vencimiento)
        fwd = md.forward_at(self.fecha_vencimiento)
        vol = md.surface.get_vol(dias, self.strike)
        df = md.df_clp_at(dias)
        mtm = mtm_option_fx(self.nominal, fwd, self.strike, vol, df,
                            dias, self.tipo, self.lado)
        pl = (self.prima_clp + mtm) if abs(self.prima_clp) > 0 else 0.0
        return {"folio": self.folio, "dias": dias, "spot": md.spot,
                "forward": fwd, "vol": vol, "mtm": mtm, "pl_no_realizado": pl}


def liquidar_vencimiento(op: "Operacion", fixing_usdobs: float):
    tipo = (op.tipo or "").strip().upper()
    lado = (op.lado or "").strip().capitalize()
    if tipo == "CALL":
        intrinseco = max(fixing_usdobs - op.strike, 0.0)
    elif tipo == "PUT":
        intrinseco = max(op.strike - fixing_usdobs, 0.0)
    else:
        intrinseco = 0.0
    flujo = intrinseco * op.nominal
    if lado == "Venta":
        flujo = -flujo
    return {"folio": op.folio, "fixing": fixing_usdobs, "itm": intrinseco > 0,
            "flujo_pago_clp": flujo}
