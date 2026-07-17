"""Parser de fechas robusto (incluye formato español 'vie 09-ene-26')."""
from datetime import date, datetime

_MESES = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,
          "sep":9,"set":9,"oct":10,"nov":11,"dic":12}

def parse_fecha(x):
    if x is None: return None
    if isinstance(x, datetime): return x.date()
    if isinstance(x, date): return x
    s = str(x).strip()
    if not s: return None
    for f in ("%Y-%m-%d %H:%M:%S","%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y"):
        try: return datetime.strptime(s[:19], f).date()
        except ValueError: pass
    # formato español: opcional 'dia ' + dd-mmm-yy
    t = s.lower().replace(".", "")
    parts = t.split()
    if len(parts) >= 1:
        cand = parts[-1]              # '09-ene-26'
        for sep in ("-", "/", " "):
            if sep in cand:
                p = cand.split(sep)
                if len(p) == 3 and p[1][:3] in _MESES:
                    dd = int(p[0]); mm = _MESES[p[1][:3]]; yy = int(p[2])
                    yy += 2000 if yy < 100 else 0
                    try: return date(yy, mm, dd)
                    except ValueError: return None
    return None
