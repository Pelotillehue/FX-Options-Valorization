"""
Marcha blanca: procesa automaticamente los adjuntos del correo diario.

Flujo completo:
  1. Power Automate guarda los 2 adjuntos del correo en una carpeta de OneDrive.
  2. OneDrive sincroniza esa carpeta a este PC (CARPETA_ENTRADA).
  3. Este script (programado en el Programador de tareas de Windows) la revisa:
     cuando estan los 2 archivos de una fecha aun no procesada, genera los
     informes en salidas/ y registra el resultado en el log.

NO envia nada por correo: los informes quedan en salidas/ para revision manual.

Es idempotente: cada fecha se procesa una sola vez (salidas/procesados.json).
Se puede correr las veces que sea; si no hay nada nuevo, termina en silencio.

Uso manual:  python procesar_correo.py
"""
import json
import os
import re
import shutil
import sys
from datetime import date, datetime

import openpyxl

from daily_process import procesar
from dateparse import parse_fecha

# ----------------------------------------------------------------------------
# CONFIGURACION — ajustar CARPETA_ENTRADA a tu ruta real de OneDrive
# ----------------------------------------------------------------------------
BASE = os.path.dirname(os.path.abspath(__file__))
CARPETA_ENTRADA = r"C:\Users\fabia\OneDrive - CREASYS S.A\FX-Opciones\Entrada"
CARPETA_DATA = os.path.join(BASE, "data")
CARPETA_SALIDAS = os.path.join(BASE, "salidas")
ARCHIVO_ESTADO = os.path.join(CARPETA_SALIDAS, "procesados.json")
ARCHIVO_LOG = os.path.join(CARPETA_SALIDAS, "proceso_correo.log")

NOMBRE_GRID = "grid vol bloomberg creasys.xlsx"
PAT_OPS = re.compile(r"operaciones diarias.*al (\d{4})_(\d{2})_(\d{2})\.xlsx$", re.I)


# ----------------------------------------------------------------------------
def log(msg):
    linea = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(linea)
    os.makedirs(CARPETA_SALIDAS, exist_ok=True)
    with open(ARCHIVO_LOG, "a", encoding="utf-8") as f:
        f.write(linea + "\n")


def cargar_estado():
    if os.path.exists(ARCHIVO_ESTADO):
        with open(ARCHIVO_ESTADO, encoding="utf-8") as f:
            return json.load(f)
    return {}


def guardar_estado(estado):
    os.makedirs(CARPETA_SALIDAS, exist_ok=True)
    with open(ARCHIVO_ESTADO, "w", encoding="utf-8") as f:
        json.dump(estado, f, indent=2, ensure_ascii=False)


def fecha_del_grid(path):
    """Fecha de proceso declarada dentro del grid (celda B28 de 'FWD pts DEPOs')."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        v = wb["FWD pts DEPOs"]["B28"].value
        wb.close()
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        if v is not None:
            return parse_fecha(str(v))
    except Exception as e:
        log(f"AVISO: no pude leer la fecha del grid {os.path.basename(path)}: {e}")
    return None


def buscar_entrada():
    """Recorre CARPETA_ENTRADA (recursivo): devuelve ({fecha: ruta_ops}, [rutas_grid])."""
    ops, grids = {}, []
    for raiz, _dirs, archivos in os.walk(CARPETA_ENTRADA):
        for nombre in archivos:
            ruta = os.path.join(raiz, nombre)
            m = PAT_OPS.search(nombre.lower())
            if m:
                f = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                ops[f] = ruta
            elif nombre.lower().endswith(NOMBRE_GRID):
                grids.append(ruta)
    return ops, grids


def elegir_grid(grids, fecha, ruta_ops):
    """Regla de emparejamiento: el grid que esta en el MISMO directorio que el
    archivo de cartera (llegaron en el mismo correo); la fecha de proceso la fija
    el nombre de la cartera. La fecha interna del grid (B28) puede diferir
    (Bloomberg la actualiza) — se avisa en el log, sin bloquear.
    Si no hay grid en el mismo directorio, se busca uno cuya fecha interna
    coincida con la fecha de proceso."""
    misma_carpeta = [g for g in grids
                     if os.path.dirname(g) == os.path.dirname(ruta_ops)]
    if misma_carpeta:
        g = misma_carpeta[0]
        interna = fecha_del_grid(g)
        if interna != fecha:
            log(f"  aviso: fecha interna del grid ({interna}) difiere de la fecha "
                f"de proceso ({fecha}); se usa igual por venir en el mismo correo.")
        return g
    for g in grids:
        if fecha_del_grid(g) == fecha:
            return g
    return None


def main():
    if not os.path.isdir(CARPETA_ENTRADA):
        log(f"ERROR: no existe la carpeta de entrada: {CARPETA_ENTRADA}")
        return 1

    estado = cargar_estado()
    ops, grids = buscar_entrada()
    pendientes = sorted(f for f in ops if f.isoformat() not in estado)
    if not pendientes:
        return 0                                   # nada nuevo: silencio

    for fecha in pendientes:
        ruta_ops = ops[fecha]
        grid = elegir_grid(grids, fecha, ruta_ops)
        if grid is None:
            log(f"{fecha}: llego el archivo de operaciones pero no encontre un "
                f"grid para emparejar; reintentare en la proxima pasada.")
            for g in grids:
                log(f"  grid visto: {g}  fecha interna={fecha_del_grid(g)}")
            continue

        log(f"{fecha}: procesando…")
        log(f"  grid: {grid}")
        log(f"  ops : {ruta_ops}")
        # copia a data/ para trazabilidad y re-ejecucion manual (no bloqueante)
        try:
            os.makedirs(CARPETA_DATA, exist_ok=True)
            shutil.copy2(grid, os.path.join(CARPETA_DATA, "Grid Vol Bloomberg Creasys.xlsx"))
            shutil.copy2(ruta_ops, os.path.join(CARPETA_DATA, os.path.basename(ruta_ops)))
        except OSError as e:
            log(f"  aviso: no pude copiar insumos a data/: {e}")

        try:
            res = procesar(grid, ruta_ops, fecha, CARPETA_SALIDAS)
        except Exception as e:
            log(f"{fecha}: ERROR al procesar: {e!r}")
            continue                                # queda pendiente para reintento

        estado[fecha.isoformat()] = {
            "procesado_en": datetime.now().isoformat(timespec="seconds"),
            "vigentes": res["vigentes"], "altas": res["altas"],
            "vencen": res["vencen"], "itm": res["itm"],
            "archivos": res["archivos"],
        }
        guardar_estado(estado)
        log(f"{fecha}: OK — {res['altas']} altas, {res['vencen']} vencimientos "
            f"({res['itm']} ITM), {res['vigentes']} vigentes. "
            f"Informes en {CARPETA_SALIDAS}\\")
        if res.get("avisos"):
            for a in res["avisos"]:
                log(f"  aviso: {a}")
    return 0


if __name__ == "__main__":
    sys.exit(main())