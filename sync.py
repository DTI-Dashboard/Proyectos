import os
import json
import urllib.request
import re
import csv
from datetime import datetime, timezone

TOKEN = os.environ["SMARTSHEET_TOKEN"]

def fmt_fecha(f):
    if not f: return ""
    f = str(f).split("T")[0]
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try: return datetime.strptime(f, fmt).strftime("%d/%m/%Y")
        except ValueError: continue
    return str(f)

def calcular_semaforo(real, esp, fin):
    hoy = datetime.today()
    try:
        fin_vencido = datetime.strptime(fmt_fecha(fin), "%d/%m/%Y") < hoy
    except Exception:
        fin_vencido = False
    if real >= 100:       return "Terminado"
    elif fin_vencido:     return "Atrasado"
    elif real >= esp:     return "A tiempo"
    elif real <= esp - 11: return "Atrasado"
    else:                 return "En riesgo"

def fetch_report(report_id, modulo_label, filter_mode="sheet_match"):
    """
    filter_mode:
      'sheet_match' → fila es proyecto si sheet_name está dentro de primary (IMC)
      'parent_row'  → fila es proyecto si parentId es None (DDI/Desarrollo)
    """
    url = f"https://api.smartsheet.com/2.0/reports/{report_id}?pageSize=1000"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/json"
    })
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())

    col_map = {}
    for c in data.get("columns", []):
        col_id = c.get("virtualId") or c.get("virtualColumnId") or c.get("id")
        if col_id:
            col_map[str(col_id)] = c.get("title", "")

    proyectos = []
    seen_sheets = set()
    hoy = datetime.today()

    for row in data.get("rows", []):
        cells = {}
        for cell in row.get("cells", []):
            col_id = str(cell.get("virtualColumnId") or cell.get("columnId") or "")
            col_title = col_map.get(col_id, "")
            val = cell.get("displayValue") or cell.get("value") or ""
            if col_title:
                cells[col_title] = val

        sheet_name = str(cells.get("Sheet Name", "") or cells.get("Nombre de la hoja", "")).strip()
        primary    = str(cells.get("Primary", "")).strip()
        avance_raw = cells.get("Avance", "")
        inicio     = cells.get("Fecha de inicio", "")
        fin        = cells.get("Fecha de finalización", "")
        parent_id  = row.get("parentId")

        if not primary:
            continue

        if filter_mode == "sheet_match":
            # IMC: el nombre del sheet está contenido en el Primary
            if not sheet_name or sheet_name.lower() not in primary.lower():
                continue
        elif filter_mode == "first_per_sheet":
            # DDI: tomar solo la primera fila de cada hoja (= el proyecto resumen)
            if not sheet_name:
                continue
            if sheet_name in seen_sheets:
                continue
            seen_sheets.add(sheet_name)

        try:
            real = int(float(str(avance_raw).replace("%", "").strip()))
        except (ValueError, TypeError):
            real = 0

        try:
            d_ini = datetime.strptime(fmt_fecha(inicio), "%d/%m/%Y")
            d_fin = datetime.strptime(fmt_fecha(fin), "%d/%m/%Y")
            total = (d_fin - d_ini).days
            trans = (hoy - d_ini).days
            esp   = max(0, min(100, int((trans / total) * 100))) if total > 0 else 0
        except Exception:
            esp = real

        sem = calcular_semaforo(real, esp, fin)

        proyectos.append({
            "nombre":      primary,
            "inicio":      fmt_fecha(inicio),
            "fin":         fmt_fecha(fin),
            "real":        real,
            "esp":         esp,
            "sem":         sem,
            "area":        "",
            "responsable": "",
            "solicitante": "",
            "descripcion": ""
        })

    print(f"✅ {len(proyectos)} proyectos {modulo_label} de Smartsheet")
    return proyectos

def load_csv(filepath):
    info = {}
    if os.path.exists(filepath):
        with open(filepath, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                nombre = row.get("nombre", "").strip()
                if nombre:
                    info[nombre.lower()] = {
                        "area":        row.get("area", ""),
                        "responsable": row.get("responsable", ""),
                        "solicitante": row.get("solicitante", ""),
                        "descripcion": row.get("descripcion", ""),
                        "url_plan":    row.get("url_plan", ""),
                    }
        print(f"✅ {filepath}: {len(info)} entradas")
    else:
        print(f"ℹ️  {filepath} no existe aún")
    return info

def merge_info(proyectos, info_extra):
    for p in proyectos:
        extra = info_extra.get(p["nombre"].lower(), {})
        for key in ("area", "responsable", "solicitante", "descripcion", "url_plan"):
            if extra.get(key): p[key] = extra[key]

def proyectos_to_js(proyectos, var_name):
    js_rows = []
    for p in proyectos:
        def esc(s): return s.replace("\\","\\\\").replace("'","\\'").replace('"','\\"')
        desc = p["descripcion"] if p["descripcion"] else "Sincronizado desde Smartsheet"
        js_rows.append(
            f'  {{tipo:"Proyecto",nombre:"{esc(p["nombre"])}",area:"{esc(p["area"])}",'
            f'inicio:"{p["inicio"]}",fin:"{p["fin"]}",real:{p["real"]},'
            f'esp:{p["esp"]},sem:"{p["sem"]}",responsable:"{esc(p["responsable"])}",'
            f'solicitante:"{esc(p["solicitante"])}",descripcion:"{esc(desc)}",'
            f'url_plan:"{esc(p["url_plan"])}"}}'
        )
    return f"const {var_name} = [\n" + ",\n".join(js_rows) + "\n];"

# ── Configuración de módulos ──────────────────────────────────
MODULOS = [
    {
        "report_id":   "1111996261945220",
        "label":       "IMC",
        "var":         "IMC_DATA",
        "csv":         "info_proyectos.csv",
        "filter_mode": "sheet_match"
    },
    {
        "report_id":   "2916326996660100",
        "label":       "Desarrollo",
        "var":         "DEV_DATA",
        "csv":         "info_dev.csv",
        "filter_mode": "first_per_sheet"
    },
]

# ── Leer index.html ───────────────────────────────────────────
with open("index.html", "r", encoding="utf-8") as f:
    html = f.read()

# ── Procesar cada módulo ──────────────────────────────────────
for mod in MODULOS:
    proyectos  = fetch_report(mod["report_id"], mod["label"], mod["filter_mode"])
    info_extra = load_csv(mod["csv"])
    merge_info(proyectos, info_extra)
    new_data   = proyectos_to_js(proyectos, mod["var"])
    pattern    = rf"const {mod['var']}\s*=\s*\[.*?\];"
    html = re.sub(pattern, lambda m: new_data, html, flags=re.DOTALL)

# ── Timestamp ─────────────────────────────────────────────────
ts_iso     = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
ts_display = datetime.now().strftime("%d/%m/%Y %H:%M")
meta_tag   = f'<meta name="last-sync" content="{ts_iso}">'
if 'name="last-sync"' in html:
    html = re.sub(r'<meta name="last-sync"[^>]*>', meta_tag, html)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"✅ index.html actualizado — {ts_display}")
