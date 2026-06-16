import os
import json
import urllib.request
import re
import csv
import io
from datetime import datetime, timezone

# ── Configuración ──────────────────────────────────────────────
TOKEN     = os.environ["SMARTSHEET_TOKEN"]
REPORT_ID = "1111996261945220"
URL       = f"https://api.smartsheet.com/2.0/reports/{REPORT_ID}?pageSize=1000"
MODULO    = "IMC"
INFO_FILE = "info_proyectos.csv"   # archivo de persistencia

# ── Llamada a la API ───────────────────────────────────────────
req = urllib.request.Request(URL, headers={
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
})
with urllib.request.urlopen(req) as resp:
    data = json.loads(resp.read())

# ── Mapear columnas ────────────────────────────────────────────
col_map = {}
for c in data.get("columns", []):
    col_id = c.get("virtualId") or c.get("virtualColumnId") or c.get("id")
    if col_id:
        col_map[str(col_id)] = c.get("title", "")

# ── Extraer filas padre ────────────────────────────────────────
proyectos = []

def fmt_fecha(f):
    if not f: return ""
    f = str(f).split("T")[0]
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try: return datetime.strptime(f, fmt).strftime("%d/%m/%Y")
        except ValueError: continue
    return str(f)

for row in data.get("rows", []):
    cells = {}
    for cell in row.get("cells", []):
        col_id = str(cell.get("virtualColumnId") or cell.get("columnId") or "")
        col_title = col_map.get(col_id, "")
        val = cell.get("displayValue") or cell.get("value") or ""
        if col_title:
            cells[col_title] = val

    sheet_name = str(cells.get("Sheet Name", "")).strip()
    primary    = str(cells.get("Primary", "")).strip()
    avance_raw = cells.get("Avance", "")
    inicio     = cells.get("Fecha de inicio", "")
    fin        = cells.get("Fecha de finalización", "")

    if not sheet_name or not primary: continue
    if sheet_name.lower() not in primary.lower(): continue

    try:
        real = int(float(str(avance_raw).replace("%", "").strip()))
    except (ValueError, TypeError):
        real = 0

    hoy = datetime.today()
    try:
        d_ini = datetime.strptime(fmt_fecha(inicio), "%d/%m/%Y")
        d_fin = datetime.strptime(fmt_fecha(fin), "%d/%m/%Y")
        total = (d_fin - d_ini).days
        trans = (hoy - d_ini).days
        esp   = max(0, min(100, int((trans / total) * 100))) if total > 0 else 0
    except Exception:
        esp = real

    if real >= 100:       sem = "Terminado"
    elif real >= esp:     sem = "A tiempo"
    elif real >= esp - 10: sem = "En riesgo"
    else:                 sem = "Atrasado"

    proyectos.append({
        "nombre": primary,
        "inicio": fmt_fecha(inicio),
        "fin":    fmt_fecha(fin),
        "real":   real,
        "esp":    esp,
        "sem":    sem,
        "area": "", "responsable": "", "solicitante": "", "descripcion": ""
    })

print(f"✅ {len(proyectos)} proyectos IMC de Smartsheet")

# ── Leer info_proyectos.csv (persistencia) ────────────────────
info_extra = {}
if os.path.exists(INFO_FILE):
    with open(INFO_FILE, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            nombre = row.get("nombre", "").strip()
            if nombre:
                info_extra[nombre.lower()] = {
                    "area":        row.get("area", ""),
                    "responsable": row.get("responsable", ""),
                    "solicitante": row.get("solicitante", ""),
                    "descripcion": row.get("descripcion", ""),
                }
    print(f"✅ info_proyectos.csv leído: {len(info_extra)} entradas")
else:
    print("ℹ️  info_proyectos.csv no existe aún — se creará al subir el CSV")

# ── Merge info extra en los proyectos ─────────────────────────
for p in proyectos:
    extra = info_extra.get(p["nombre"].lower(), {})
    if extra.get("area"):        p["area"]        = extra["area"]
    if extra.get("responsable"): p["responsable"] = extra["responsable"]
    if extra.get("solicitante"): p["solicitante"] = extra["solicitante"]
    if extra.get("descripcion"): p["descripcion"] = extra["descripcion"]

# ── Inyectar en index.html ─────────────────────────────────────
with open("index.html", "r", encoding="utf-8") as f:
    html = f.read()

js_rows = []
for p in proyectos:
    nombre_js = p["nombre"].replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"')
    area_js   = p["area"].replace("'", "\\'")
    resp_js   = p["responsable"].replace("'", "\\'")
    soli_js   = p["solicitante"].replace("'", "\\'")
    desc_js   = p["descripcion"].replace("'", "\\'") if p["descripcion"] else "Sincronizado desde Smartsheet"
    js_rows.append(
        f'  {{tipo:"Proyecto",nombre:"{nombre_js}",area:"{area_js}",'
        f'inicio:"{p["inicio"]}",fin:"{p["fin"]}",real:{p["real"]},'
        f'esp:{p["esp"]},sem:"{p["sem"]}",responsable:"{resp_js}",'
        f'solicitante:"{soli_js}",descripcion:"{desc_js}"}}'
    )

new_imc_data = "const IMC_DATA = [\n" + ",\n".join(js_rows) + "\n];"
pattern = r"const IMC_DATA\s*=\s*\[.*?\];"
html = re.sub(pattern, new_imc_data, html, flags=re.DOTALL)

# ── Escribir timestamp de sync (meta tag) ──────────────────────
ts_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
ts_display = datetime.now().strftime("%d/%m/%Y %H:%M")

# Reemplazar o insertar meta last-sync
meta_tag = f'<meta name="last-sync" content="{ts_iso}">'
if 'name="last-sync"' in html:
    html = re.sub(r'<meta name="last-sync"[^>]*>', meta_tag, html)
else:
    html = html.replace('<!-- LAST_SYNC -->\n</head>', f'<!-- LAST_SYNC -->\n{meta_tag}\n</head>')

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"✅ index.html actualizado — {ts_display}")
print(f"✅ Timestamp de sync: {ts_iso}")
