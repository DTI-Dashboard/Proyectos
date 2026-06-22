import os, json, urllib.request, traceback, sys
from datetime import datetime

TOKEN = os.environ["SMARTSHEET_TOKEN"]
output = []

def log(msg):
    print(msg)
    output.append(msg)

def fetch_report_debug(report_id, label):
    url = f"https://api.smartsheet.com/2.0/reports/{report_id}?pageSize=10000"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"})
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())
    
    total_rows = data.get("totalRowCount", "?")
    returned_rows = len(data.get("rows", []))
    log(f"📊 {label}: totalRowCount={total_rows}, rows devueltas={returned_rows}")
    
    col_map = {}
    for c in data.get("columns", []):
        col_id = c.get("virtualId") or c.get("id")
        if col_id:
            col_map[str(col_id)] = c.get("title", "")
    log(f"   Columnas: {list(col_map.values())}")
    
    proyectos = []
    seen = set()
    for row in data.get("rows", []):
        cells = {}
        for cell in row.get("cells", []):
            col_id = str(cell.get("virtualColumnId") or cell.get("columnId") or "")
            col_title = col_map.get(col_id, "")
            val = cell.get("displayValue") or cell.get("value") or ""
            if col_title: cells[col_title] = val
        
        sheet_name = str(cells.get("Sheet Name", "") or cells.get("Nombre de la hoja", "")).strip()
        primary = str(cells.get("Primary", "")).strip()
        
        if not primary: continue
        if not sheet_name: continue
        if sheet_name.lower() in primary.lower():
            proyectos.append(f"{sheet_name} | {primary[:40]} | {cells.get('Avance','')}")
    
    log(f"   ✅ Proyectos encontrados: {len(proyectos)}")
    for p in proyectos:
        log(f"      - {p}")

try:
    log(f"=== Test sync {datetime.now().strftime('%d/%m/%Y %H:%M')} ===")
    fetch_report_debug("1111996261945220", "IMC Informe General")
    log("✅ Completado sin errores")
except Exception as e:
    log(f"❌ ERROR: {type(e).__name__}: {e}")
    log(traceback.format_exc())

with open("sync_debug.txt", "w") as f:
    f.write("\n".join(output))
print("\n".join(output))
