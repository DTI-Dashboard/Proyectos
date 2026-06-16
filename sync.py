import os
import json
import urllib.request
import urllib.error
from datetime import datetime

# ── Configuración ──────────────────────────────────────────────
TOKEN     = os.environ["SMARTSHEET_TOKEN"]
REPORT_ID = "1111996261945220"
URL       = f"https://api.smartsheet.com/2.0/reports/{REPORT_ID}?pageSize=1000&include=columnType"
MODULO    = "IMC"

# ── Llamada a la API ───────────────────────────────────────────
req = urllib.request.Request(URL, headers={
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
})
with urllib.request.urlopen(req) as resp:
    data = json.loads(resp.read())

# ── Mapear columnas por virtualColumnId → título ───────────────
col_map = {str(c["virtualColumnId"]): c["title"] for c in data["columns"]}

# ── Extraer filas padre ────────────────────────────────────────
# La fila padre es donde el campo "Sheet Name" == campo "Primary"
# (así lo filtraba tu Power Query)
proyectos = []

for row in data["rows"]:
    cells = {}
    for cell in row["cells"]:
        col_title = col_map.get(str(cell.get("virtualColumnId", "")), "")
        if col_title:
            cells[col_title] = cell.get("displayValue") or cell.get("value") or ""

    sheet_name = cells.get("Sheet Name", "").strip()
    primary    = cells.get("Primary", "").strip()
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

    # Convertir fechas: "2026-03-01" → "01/03/2026"
    def fmt_fecha(f):
        if not f:
            return ""
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(str(f), fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return str(f)

    # Calcular semáforo y esperado automáticamente
    # esperado = avance teórico basado en tiempo transcurrido
    hoy = datetime.today()
    try:
        d_ini = datetime.strptime(fmt_fecha(inicio), "%d/%m/%Y")
        d_fin = datetime.strptime(fmt_fecha(fin),    "%d/%m/%Y")
        total = (d_fin - d_ini).days
        trans = (hoy  - d_ini).days
        esp   = max(0, min(100, int((trans / total) * 100))) if total > 0 else 0
    except Exception:
        esp = real

    # Semáforo
    if real >= 100:
        sem = "Terminado"
    elif real >= esp:
        sem = "A tiempo"
    elif real >= esp - 10:
        sem = "En riesgo"
    else:
        sem = "Atrasado"

    proyectos.append({
        "modulo":      MODULO,
        "tipo":        "Proyecto",
        "nombre":      primary,
        "area":        "",
        "inicio":      fmt_fecha(inicio),
        "fin":         fmt_fecha(fin),
        "real":        real,
        "esperado":    esp,
        "semaforo":    sem,
        "responsable": "",
        "solicitante": "",
        "descripcion": ""
    })

print(f"✅ {len(proyectos)} proyectos IMC extraídos de Smartsheet")

# ── Generar CSV ────────────────────────────────────────────────
csv_lines = ["modulo,tipo,nombre,area,inicio,fin,real,esperado,semaforo,responsable,solicitante,descripcion"]
for p in proyectos:
    nombre_safe = p["nombre"].replace(",", " ").replace('"', '')
    csv_lines.append(
        f'{p["modulo"]},{p["tipo"]},"{nombre_safe}",{p["area"]},'
        f'{p["inicio"]},{p["fin"]},{p["real"]},{p["esperado"]},'
        f'{p["semaforo"]},{p["responsable"]},{p["solicitante"]},{p["descripcion"]}'
    )

csv_content = "\n".join(csv_lines)

# ── Inyectar en index.html ────────────────────────────────────
# Reemplaza el array IMC_DATA en el HTML con los datos frescos de Smartsheet
with open("index.html", "r", encoding="utf-8") as f:
    html = f.read()

# Construir el nuevo IMC_DATA como array JS
js_rows = []
for p in proyectos:
    nombre_js = p["nombre"].replace("'", "\\'")
    js_rows.append(
        f'  {{tipo:"{p["tipo"]}",nombre:"{nombre_js}",area:"{p["area"]}",'
        f'inicio:"{p["inicio"]}",fin:"{p["fin"]}",real:{p["real"]},'
        f'esp:{p["esperado"]},sem:"{p["semaforo"]}",responsable:"",'
        f'solicitante:"",descripcion:"Sincronizado desde Smartsheet"}}'
    )

new_imc_data = "const IMC_DATA = [\n" + ",\n".join(js_rows) + "\n];"

# Reemplazar bloque IMC_DATA en el HTML
import re
pattern = r"const IMC_DATA\s*=\s*\[.*?\];"
new_html = re.sub(pattern, new_imc_data, html, flags=re.DOTALL)

# Agregar timestamp de última actualización
ts = datetime.now().strftime("%d/%m/%Y %H:%M")
new_html = new_html.replace(
    "<!-- LAST_SYNC -->",
    f"<!-- LAST_SYNC --><span id='last-sync'>Última sync: {ts}</span>"
)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(new_html)

print(f"✅ index.html actualizado con {len(proyectos)} proyectos IMC")
print(f"✅ Timestamp: {ts}")
