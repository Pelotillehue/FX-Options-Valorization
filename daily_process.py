"""
Pipeline diario de valorizacion de Opciones FX (USD/CLP).

Entradas (los 2 adjuntos del correo diario):
  1. Grid Vol Bloomberg Creasys.xlsx           -> datos de mercado
  2. Operaciones Diarias ... al YYYY_MM_DD.xlsx -> cartera / altas / vencimientos

Reconstruye la cartera vigente del dia como {inicio <= fecha_proceso < vencimiento}
(auto-consistente, sin estado acumulado), valoriza cada operacion, liquida las
que vencen en la fecha de proceso usando el fixing USDOBS, y genera:
  - MtM_YYYYMMDD.xlsx                (cartera vigente valorizada)
  - MtM_YYYYMMDD_vencimientos.xlsx   (operaciones que vencen + flujo de pago)
  - resumen_YYYYMMDD.txt             (conteos tipo correo)

Uso:
  python daily_process.py <grid.xlsx> <operaciones.xlsx> [YYYY-MM-DD] [dir_salida]
"""
import sys, os
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from dateparse import parse_fecha
from load_market import load_market
from fx_options_valuation import Operacion, liquidar_vencimiento


# --------------------------------------------------------------------------
def leer_operaciones(path):
    """Lee la hoja 'Operaciones Diarias' -> lista de dicts con todos los campos."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Operaciones Diarias"]
    ops, avisos = [], []
    for r in range(2, ws.max_row + 1):
        folio = ws.cell(r, 1).value
        if not folio:
            continue
        inicio = parse_fecha(ws.cell(r, 2).value)
        venc = parse_fecha(ws.cell(r, 3).value)
        if inicio is None or venc is None:
            avisos.append(f"fila {r} ({folio}): fecha no reconocida "
                          f"inicio={ws.cell(r,2).value!r} venc={ws.cell(r,3).value!r}")
            continue
        ops.append(dict(
            folio=folio, inicio=inicio, venc=venc,
            contraparte=ws.cell(r, 6).value, lado=ws.cell(r, 7).value,
            tipo=ws.cell(r, 8).value, modalidad=ws.cell(r, 9).value,
            entrega=ws.cell(r, 10).value, moneda=ws.cell(r, 11).value,
            nominal=ws.cell(r, 12).value, strike=ws.cell(r, 13).value,
            prima_clp_usd=ws.cell(r, 15).value, prima_clp=ws.cell(r, 16).value,
            usdobs=ws.cell(r, 25).value, itm_src=ws.cell(r, 26).value,
        ))
    wb.close()
    return ops, avisos


def clasificar(ops, fecha_proceso):
    """vigentes (a valorizar) y vencimientos (a liquidar en la fecha de proceso)."""
    vigentes = [o for o in ops if o["inicio"] <= fecha_proceso < o["venc"]]
    vencen = [o for o in ops if o["venc"] == fecha_proceso]
    altas = [o for o in ops if o["inicio"] == fecha_proceso]
    return vigentes, vencen, altas


def _op_obj(o):
    return Operacion(folio=o["folio"], fecha_inicio=o["inicio"],
                     fecha_vencimiento=o["venc"], tipo=o["tipo"], lado=o["lado"],
                     nominal=o["nominal"] or 0, strike=o["strike"] or 0,
                     prima_clp=o["prima_clp"] or 0)


# --------------------------------------------------------------------------
# Estilos Excel
# --------------------------------------------------------------------------
_HDR_FILL = PatternFill("solid", fgColor="1F2A44")
_HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
_TIT_FONT = Font(bold=True, size=13, color="1F2A44")
_THIN = Side(style="thin", color="D0D0D0")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_MONEY = '#,##0;[Red]-#,##0'
_NUM2 = '#,##0.00'


def _estilo_encabezado(ws, fila, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(fila, c)
        cell.fill = _HDR_FILL; cell.font = _HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDER


def escribir_mtm(vigentes, resultados, md, path, fecha_proceso):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "MtM"
    ws["A1"] = f"Mark-to-Market Opciones FX  ·  al {fecha_proceso:%d-%b-%Y}"
    ws["A1"].font = _TIT_FONT
    ws["A2"] = f"Spot USDCLP: {md.spot:,.2f}   ·   Operaciones vigentes: {len(vigentes)}"
    ws["A2"].font = Font(italic=True, size=9, color="555555")
    cols = ["Folio","Fecha Inicio","Fecha Vencimiento","Plazo (días)","Contraparte",
            "Lado","Tipo","Moneda","Spot","Volatilidad (%)","Forward","Monto Principal",
            "Strike","Prima (CLP)","MtM (CLP)","P&L No Realizado (CLP)"]
    hr = 4
    for i, h in enumerate(cols, 1):
        ws.cell(hr, i, h)
    _estilo_encabezado(ws, hr, len(cols))
    tot_mtm = tot_pl = tot_prima = 0.0
    for k, (o, res) in enumerate(zip(vigentes, resultados)):
        row = hr + 1 + k
        vals = [o["folio"], o["inicio"], o["venc"], res["dias"], o["contraparte"],
                o["lado"], o["tipo"], o["moneda"], md.spot, res["vol"], res["forward"],
                o["nominal"], o["strike"], o["prima_clp"] or 0, res["mtm"], res["pl_no_realizado"]]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(row, i, v); cell.border = _BORDER
        for i in (2, 3):
            ws.cell(row, i).number_format = 'yyyy-mm-dd'
        ws.cell(row, 9).number_format = _NUM2
        ws.cell(row, 10).number_format = _NUM2
        ws.cell(row, 11).number_format = _NUM2
        ws.cell(row, 13).number_format = _NUM2
        for i in (12, 14, 15, 16):
            ws.cell(row, i).number_format = _MONEY
        tot_mtm += res["mtm"]; tot_pl += res["pl_no_realizado"]; tot_prima += (o["prima_clp"] or 0)
    tr = hr + 1 + len(vigentes)
    ws.cell(tr, 1, "TOTAL").font = Font(bold=True)
    for i, v in ((14, tot_prima), (15, tot_mtm), (16, tot_pl)):
        c = ws.cell(tr, i, v); c.font = Font(bold=True); c.number_format = _MONEY
        c.fill = PatternFill("solid", fgColor="EEF1F7")
    _auto_width(ws, len(cols))
    ws.freeze_panes = "A5"
    wb.save(path)


def escribir_vencimientos(vencen, liqs, path, fecha_proceso):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Vencimientos"
    ws["A1"] = f"Vencimientos  ·  fixing {fecha_proceso:%d-%b-%Y}"
    ws["A1"].font = _TIT_FONT
    cols = ["Folio","Fecha Inicio","Fecha Vencimiento","Contraparte","Lado","Tipo",
            "Monto Principal","Strike","Fixing USDOBS","ITM","Prima (CLP)",
            "Flujo de Pago (CLP)"]
    hr = 3
    for i, h in enumerate(cols, 1):
        ws.cell(hr, i, h)
    _estilo_encabezado(ws, hr, len(cols))
    tot_flujo = tot_prima = 0.0
    for k, (o, liq) in enumerate(zip(vencen, liqs)):
        row = hr + 1 + k
        vals = [o["folio"], o["inicio"], o["venc"], o["contraparte"], o["lado"], o["tipo"],
                o["nominal"], o["strike"], liq["fixing"], "ITM" if liq["itm"] else "OTM",
                o["prima_clp"] or 0, liq["flujo_pago_clp"]]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row, i, v); c.border = _BORDER
        for i in (2, 3):
            ws.cell(row, i).number_format = 'yyyy-mm-dd'
        ws.cell(row, 9).number_format = _NUM2
        ws.cell(row, 8).number_format = _NUM2
        for i in (7, 11, 12):
            ws.cell(row, i).number_format = _MONEY
        if liq["itm"]:
            ws.cell(row, 10).font = Font(bold=True, color="B00000")
        tot_flujo += liq["flujo_pago_clp"]; tot_prima += (o["prima_clp"] or 0)
    tr = hr + 1 + len(vencen)
    ws.cell(tr, 1, "TOTAL").font = Font(bold=True)
    for i, v in ((11, tot_prima), (12, tot_flujo)):
        c = ws.cell(tr, i, v); c.font = Font(bold=True); c.number_format = _MONEY
        c.fill = PatternFill("solid", fgColor="EEF1F7")
    _auto_width(ws, len(cols))
    wb.save(path)


def _auto_width(ws, ncols):
    for c in range(1, ncols + 1):
        letra = get_column_letter(c)
        largo = 10
        for cell in ws[letra]:
            if cell.value is not None:
                largo = max(largo, len(str(cell.value)) + 2)
        ws.column_dimensions[letra].width = min(largo, 22)


# --------------------------------------------------------------------------
def procesar(grid_path, ops_path, fecha_proceso=None, out_dir="."):
    ops, avisos = leer_operaciones(ops_path)
    md = load_market(grid_path, fecha_proceso=fecha_proceso)
    fecha_proceso = md.fecha_proceso
    vigentes, vencen, altas = clasificar(ops, fecha_proceso)

    resultados = [_op_obj(o).valorizar(md) for o in vigentes]
    liqs = [liquidar_vencimiento(_op_obj(o), o["usdobs"]) for o in vencen]

    tag = fecha_proceso.strftime("%Y%m%d")
    p_mtm = os.path.join(out_dir, f"MtM_{tag}.xlsx")
    p_ven = os.path.join(out_dir, f"MtM_{tag}_vencimientos.xlsx")
    p_res = os.path.join(out_dir, f"resumen_{tag}.txt")
    escribir_mtm(vigentes, resultados, md, p_mtm, fecha_proceso)
    escribir_vencimientos(vencen, liqs, p_ven, fecha_proceso)

    n_itm = sum(1 for l in liqs if l["itm"])
    tot_mtm = sum(r["mtm"] for r in resultados)
    tot_pl = sum(r["pl_no_realizado"] for r in resultados)
    tot_flujo = sum(l["flujo_pago_clp"] for l in liqs)
    resumen = (
        f"Proceso de valorizacion Opciones FX — {fecha_proceso:%d-%m-%Y}\n"
        f"{'='*52}\n"
        f"Spot USDCLP: {md.spot:,.2f}\n\n"
        f"Hubo {len(altas)} nuevas operaciones.\n"
        f"Vencieron {len(vencen)} operaciones, {n_itm} ITM por lo que hay flujos de pago.\n"
        f"Hay {len(vigentes)} operaciones vigentes.\n\n"
        f"MtM total cartera:         {tot_mtm:,.0f} CLP\n"
        f"P&L no realizado total:    {tot_pl:,.0f} CLP\n"
        f"Flujo de pago vencimientos:{tot_flujo:,.0f} CLP\n"
    )
    if altas:
        resumen += f"\nAltas del dia: {', '.join(o['folio'] for o in altas)}\n"
    if avisos:
        resumen += "\nAvisos:\n" + "\n".join("  - " + a for a in avisos) + "\n"
    with open(p_res, "w", encoding="utf-8") as f:
        f.write(resumen)
    return dict(fecha=fecha_proceso, vigentes=len(vigentes), vencen=len(vencen),
                itm=n_itm, altas=len(altas), tot_mtm=tot_mtm, tot_pl=tot_pl,
                tot_flujo=tot_flujo, avisos=avisos,
                archivos=[p_mtm, p_ven, p_res], resumen=resumen)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    grid, opsf = sys.argv[1], sys.argv[2]
    fp = parse_fecha(sys.argv[3]) if len(sys.argv) > 3 else None
    outd = sys.argv[4] if len(sys.argv) > 4 else "."
    res = procesar(grid, opsf, fp, outd)
    print(res["resumen"])
    print("Archivos generados:")
    for a in res["archivos"]:
        print("  ", a)
