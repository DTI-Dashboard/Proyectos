import os
import json
import urllib.request
import re
from datetime import datetime

# ── Configuración ──────────────────────────────────────────────
TOKEN     = os.environ["SMARTSHEET_TOKEN"]
REPORT_ID = "1111996261945220"
URL       = f"https://api.smartsheet.com/2.0/reports/{REPORT_ID}?pageSize=1000"
MODULO    = "IMC"

# ── Llamada a la API ───────────────────────────────────────────
req = urllib.request.Request(URL, headers={
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
})
with urllib.request.urlopen(req) as resp:
    data = json.loads(resp.read())

# ── Debug: imprimir estructura de columnas ─────────────────────
print("=== COLUMNAS DISPONIBLES ===")
for c in data.get("columns", []):
    print(f"  keys: {list(c.keys())} → {c}")

# ── Mapear columnas: soporta tanto 'id' como 'virtualColumnId' ──
col_map = {}
for c in data.get("columns", []):
    col_id = str(c.get("virtualColumnId") or c.get("id") or "")
    col_map[col_id] = c.get("title", "")

print(f"\n=== COL_MAP ===\n{col_map}\n")

# ── Imprimir primera fila para ver estructura ──────────────────
if data.get("rows"):
    print("=== PRIMERA FILA (cells) ===")
    for cell in data["rows"][0].get("cells", []):
        print(f"  {cell}")

# ── Extraer filas padre ────────────────────────────────────────
proyectos = []

for row in data.get("rows", []):
    cells = {}
    for cell in row.get("cells", []):
        col_id = str(cell.get("virtualColumnId") or cell.get("columnId") or "")
        col_title = col_map.get(col_id, "")
        if col_title:
            cells[col_title] = cell.get("displayValue") or cell.get("value") or ""

    sheet_name = str(cells.get("Sheet Name", "")).strip()
    primary    = str(cells.get("Primary", "")).strip()
    avance_raw = cells.get("Avance", "")
    inicio     = cells.get("Fecha de inicio", "")
    fin        = cells.get("Fecha de finalización", "")

    # Filtro fila padre: Sheet Name está contenido en Primary
    if not sheet_name or not primary:
        continue
    if sheet_name.lower() not in primary.lower():
        continue

    # Convertir avance: "50%" → 50
    try:
        real = int(float(str(avance_raw).replace("%", "").strip()))
    except (ValueError, TypeError):
        real = 0

    # Convertir fechas al formato dd/mm/yyyy
    def fmt_fecha(f):
        if not f:
            return ""
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(str(f), fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return str(f)

    # Calcular esperado basado en tiempo transcurrido
    hoy = datetime.today()
    try:
        d_ini = datetime.strptime(fmt_fecha(inicio), "%d/%m/%Y")
        d_fin = datetime.strptime(fmt_fecha(fin),    "%d/%m/%Y")
        total = (d_fin - d_ini).days
        trans = (hoy  - d_ini).days
        esp   = max(0, min(100, int((trans / total) * 100))) if total > 0 else 0
    except Exception:
        esp = real

    # Semáforo automático
    if real >= 100:
        sem = "Terminado"
    elif real >= esp:
        sem = "A tiempo"
    elif real >= esp - 10:
        sem = "En riesgo"
    else:
        sem = "Atrasado"

    proyectos.append({
        "tipo":        "Proyecto",
        "nombre":      primary,
        "area":        "",
        "inicio":      fmt_fecha(inicio),
        "fin":         fmt_fecha(fin),
        "real":        real,
        "esp":         esp,
        "sem":         sem,
    })

print(f"\n✅ {len(proyectos)} proyectos IMC extraídos")
for p in proyectos:
    print(f"  - {p['nombre']} | {p['real']}% | {p['sem']}")

# ── Inyectar en index.html ─────────────────────────────────────
with open("index.html", "r", encoding="utf-8") as f:
    html = f.read()

js_rows = []
for p in proyectos:
    nombre_js = p["nombre"].replace("'", "\\'").replace('"', '\\"')
    js_rows.append(
        f'  {{tipo:"{p["tipo"]}",nombre:"{nombre_js}",area:"",'
        f'inicio:"{p["inicio"]}",fin:"{p["fin"]}",real:{p["real"]},'
        f'esp:{p["esp"]},sem:"{p["sem"]}",responsable:"",'
        f'solicitante:"",descripcion:"Sincronizado desde Smartsheet"}}'
    )

new_imc_data = "const IMC_DATA = [\n" + ",\n".join(js_rows) + "\n];"

pattern = r"const IMC_DATA\s*=\s*\[.*?\];"
new_html = re.sub(pattern, new_imc_data, html, flags=re.DOTALL)

ts = datetime.now().strftime("%d/%m/%Y %H:%M")
with open("index.html", "w", encoding="utf-8") as f:
    f.write(new_html)

print(f"\n✅ index.html actualizado — {ts}")
