"""Regresion con archivos reales. Se salta si no estan en data/ (no van al repo)."""
import os
from datetime import date
import pytest

DATA = os.path.join(os.path.dirname(__file__), "..", "data")
GRID = os.path.join(DATA, "Grid Vol Bloomberg Creasys.xlsx")
OPS = os.path.join(DATA, "Operaciones Diarias Opciones y Cartera Vigente al 2026_07_14.xlsx")

requiere_datos = pytest.mark.skipif(
    not (os.path.exists(GRID) and os.path.exists(OPS)),
    reason="archivos reales no disponibles en data/")


@requiere_datos
def test_pipeline_2026_07_14(tmp_path):
    import daily_process as dp
    res = dp.procesar(GRID, OPS, date(2026, 7, 14), str(tmp_path))
    assert res["vigentes"] == 64
    assert res["altas"] == 4
    assert res["vencen"] == 4
    assert res["itm"] == 2
    assert res["tot_mtm"] == pytest.approx(0.0, abs=1.0)        # cartera calzada
    assert res["tot_pl"] == pytest.approx(71960550, abs=2)
    for a in res["archivos"]:
        assert os.path.exists(a)
