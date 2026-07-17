"""Carga MarketData desde 'Grid Vol Bloomberg Creasys.xlsx' (reemplaza el notebook)."""
import openpyxl
from datetime import datetime
from fx_options_valuation import MarketData, STD_TENORS_DAYS

# etiqueta de tenor -> dias (para mapear filas de Mids BGN a dias estandar)
TENOR_DIAS = dict(zip(
    ["1D","1W","2W","3W","1M","2M","3M","4M","6M","9M","1Y","18M","2Y","3Y","4Y","5Y"],
    STD_TENORS_DAYS))

def load_market(grid_path, fecha_proceso=None):
    wb = openpyxl.load_workbook(grid_path, data_only=True)
    # ---- Curva cruda + spot desde 'FWD pts DEPOs' ----
    ws = wb["FWD pts DEPOs"]
    spot = ws["H28"].value            # 926.78
    fecha_hdr = ws["B28"].value       # fecha de proceso
    raw_dias, raw_clp, raw_usd = [], [], []
    r = 30
    while True:
        c = ws.cell(r, 3).value       # C = Dias
        g = ws.cell(r, 7).value       # G = Mid CLP
        i = ws.cell(r, 9).value       # I = Mid USD implicito
        if not isinstance(c, (int, float)):
            break
        if isinstance(g, (int, float)) and isinstance(i, (int, float)):
            raw_dias.append(float(c)); raw_clp.append(float(g)); raw_usd.append(float(i))
        r += 1
    # ---- Superficie de vol desde 'Mids BGN' (bloque 2, cols M:R) ----
    ws = wb["Mids BGN"]
    vol_tenor_dias, vol_smile = [], []
    for rr in range(3, 3 + 16):       # 1D .. 5Y
        term = ws.cell(rr, 13).value  # M = Term
        if term is None:
            continue
        term = str(term).strip()
        if term not in TENOR_DIAS:
            continue
        vatm = ws.cell(rr, 14).value  # N
        v25c = ws.cell(rr, 15).value  # O 25 Call
        v25p = ws.cell(rr, 16).value  # P 25 Put
        v10c = ws.cell(rr, 17).value  # Q 10 Call
        v10p = ws.cell(rr, 18).value  # R 10 Put
        vol_tenor_dias.append(TENOR_DIAS[term])
        vol_smile.append((float(v10p), float(v25p), float(vatm), float(v25c), float(v10c)))
    wb.close()
    fp = fecha_proceso or (fecha_hdr.date() if isinstance(fecha_hdr, datetime) else fecha_hdr)
    md = MarketData(fecha_proceso=fp, spot=float(spot),
                    raw_dias=raw_dias, raw_clp=raw_clp, raw_usd=raw_usd,
                    vol_tenor_dias=vol_tenor_dias, vol_smile=vol_smile)
    return md.build()
