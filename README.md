# Valorización automática de Opciones FX (USD/CLP)

Reemplaza el proceso manual diario (notebook de extracción + pegado de CSV en el
maestro Excel + altas/vencimientos a mano) por un proceso en Python. La
valorización **coincide al centavo** con el modelo VBA original (validado sobre
las 64 operaciones del 2026-07-14: error máximo 0,000000 CLP en MtM, forwards,
vols y strikes).

## Estructura

```
fx-options-valorizacion/
├── fx_options_valuation.py   # Motor: curvas, superficie de vol, Black-Scholes FX, MtM
├── load_market.py            # Lee 'Grid Vol Bloomberg Creasys.xlsx' (reemplaza el notebook)
├── dateparse.py              # Parser de fechas robusto (incl. formato español)
├── daily_process.py          # Orquestador diario (único punto de entrada)
├── requirements.txt          # Dependencias (solo openpyxl)
├── data/                     # (ignorado por git) los 2 adjuntos del correo
└── salidas/                  # (ignorado por git) resultados generados
```

Solo se ejecuta `daily_process.py`; los otros módulos los usa por dentro.

## Requisitos

- Python 3.9 o superior.
- Única dependencia externa: `openpyxl` (el resto es librería estándar).

## Instalación (entorno virtual)

Desde la carpeta del proyecto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1        # PowerShell  (CMD: .\.venv\Scripts\activate.bat)
pip install -r requirements.txt
```

## Uso diario

1. Guarda los 2 adjuntos del correo en `data/`.
2. Con el entorno activo, corre:

```powershell
python daily_process.py "data\Grid Vol Bloomberg Creasys.xlsx" "data\Operaciones Diarias Opciones y Cartera Vigente al 2026_07_14.xlsx" 2026-07-14 salidas
```

(La fecha y la carpeta de salida son opcionales.)

Genera en `salidas/`:
- `MtM_YYYYMMDD.xlsx` — cartera vigente valorizada.
- `MtM_YYYYMMDD_vencimientos.xlsx` — vencimientos + payoff con fixing USDOBS.
- `resumen_YYYYMMDD.txt` — conteos tipo correo (altas / vencimientos / vigentes).

## Uso como módulo (para automatizar desde el correo)

```python
from daily_process import procesar

res = procesar(grid_path="data/Grid Vol ....xlsx",
               ops_path="data/Operaciones ....xlsx",
               fecha_proceso=None,       # o date(2026,7,14)
               out_dir="salidas")
print(res["resumen"])                    # cuerpo del correo
```

## Lógica de cartera

Cada día reconstruye la cartera vigente desde `Operaciones Diarias` con la regla
`inicio ≤ fecha_proceso < vencimiento` (auto-consistente, sin estado acumulado).
- **Altas** = `inicio == fecha_proceso`
- **Vencimientos** = `vencimiento == fecha_proceso` (payoff con fixing USDOBS)

## Validación

- Los 64 MtM coinciden con el Excel (Δ máx = 0,000000 CLP).
- Forwards, vols y strikes de la superficie coinciden exactamente.
- Los payoffs de vencimientos coinciden con el archivo `_vencimientos` de referencia.

## Nota de seguridad

`data/` y `salidas/` están en `.gitignore`: los archivos de mercado y cartera
**no se suben** al repositorio. Usa un repositorio **privado**.
