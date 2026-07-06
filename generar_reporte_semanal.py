"""
generar_reporte_semanal.py
──────────────────────────
Regenera SOLO la base de cierre_base.pptx (slides 1 y 2: KPIs + tabla IMC
con datos actuales) BAJO DEMANDA — ya no corre dentro del sync horario.

La diapositiva 3 (eventualidades) NO se toca aquí: eso lo arma el propio
navegador en exportCierreSemanal() usando localStorage (dti_eventualidades),
que es la fuente de verdad real y siempre está más actualizada que cualquier
copia en el repo.

Se dispara desde el botón "📋 Reporte Semanal (pptx)" del módulo IMC, vía
el workflow .github/workflows/reporte_semanal.yml (workflow_dispatch).

Requiere el secret GH_TOKEN (o PAT_TOKEN) con permisos de escritura sobre
el repo, para poder subir cierre_base.pptx actualizado.
"""
import traceback as _tb
import sys as _sys

_error_file = "reporte_last_error.txt"

try:
    import os, json, re, base64, io
    import urllib.request as _ur
    from datetime import datetime, date, timedelta

    GH_TOKEN = os.environ.get("GH_TOKEN", "") or os.environ.get("PAT_TOKEN", "")
    if not GH_TOKEN:
        raise Exception("Falta el secret GH_TOKEN (o PAT_TOKEN) en el workflow")

    # ── Descargar el index.html actual del repo (ya trae IMC_DATA fresco del último sync) ──
    req_html = _ur.Request(
        "https://api.github.com/repos/DTI-Dashboard/Proyectos/contents/index.html",
        headers={"Authorization": f"token {GH_TOKEN}"}
    )
    with _ur.urlopen(req_html) as r:
        info = json.loads(r.read())
    with _ur.urlopen(info["download_url"]) as r:
        html = r.read().decode("utf-8")

    from pptx import Presentation
    from pptx.util import Inches as I, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    import lxml.etree as etree

    # Leer datos IMC del HTML actual
    idx_imc = html.find('const IMC_DATA = [')
    depth = 0
    i_pos = idx_imc + len('const IMC_DATA = ')
    for j in range(i_pos, i_pos+500000):
        if html[j] == '[': depth += 1
        elif html[j] == ']':
            depth -= 1
            if depth == 0:
                block = html[i_pos:j+1]
                break
    matches = re.findall(
        r'\{tipo:"([^"]*)",nombre:"([^"]*)",area:"([^"]*)",inicio:"([^"]*)",fin:"([^"]*)",real:(\d+),esp:(\d+),sem:"([^"]*)"',
        block
    )
    datos = [{'tipo':m[0],'nombre':m[1],'area':m[2],'inicio':m[3],'fin':m[4],'real':int(m[5]),'esp':int(m[6]),'sem':m[7]} for m in matches]

    # Descargar template
    tpl_url = 'https://raw.githubusercontent.com/DTI-Dashboard/Proyectos/main/template_cierre.pptx'
    with _ur.urlopen(tpl_url) as _r:
        tpl_bytes = _r.read()

    hoy_d = date.today()
    dow = hoy_d.weekday()
    dias_atras = (dow - 3) % 7
    jP = hoy_d - timedelta(days=dias_atras)
    jA = jP + timedelta(days=7)
    MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    rango = f"{jP.day} de {MESES[jP.month-1]} al {jA.day} de {MESES[jA.month-1]}"

    prs = Presentation(io.BytesIO(tpl_bytes))
    NAVY=RGBColor(0x1A,0x23,0x40); GOLD=RGBColor(0xC9,0xA8,0x4C)
    WHITE=RGBColor(0xFF,0xFF,0xFF); GRAY=RGBColor(0x55,0x55,0x55)
    sem_col = {'Terminado':RGBColor(0x22,0xC5,0x5E),'A tiempo':RGBColor(0x22,0xC5,0x5E),'Atrasado':RGBColor(0xEF,0x44,0x44),'Stand By':RGBColor(0x11,0x11,0x11),'Riesgo de atraso':RGBColor(0xF5,0x9E,0x0B)}
    sem_hex = {'Terminado':'22C55E','A tiempo':'22C55E','Atrasado':'EF4444','Stand By':'111111','Riesgo de atraso':'F59E0B'}
    sem_txt = {'Terminado':'Terminado','A tiempo':'A tiempo','Atrasado':'Atrasado','Stand By':'Stand by','Riesgo de atraso':'Riesgo de atraso'}

    def sc(p):
        s=(p['sem'] or '').lower()
        if 'terminado' in s: return 'Terminado'
        if 'tiempo' in s: return 'A tiempo'
        if 'riesgo' in s: return 'Riesgo de atraso'
        if 'stand' in s: return 'Stand By'
        return 'Atrasado'

    def cell_txt(cell, text, sz=9, bold=False, color=None, align=PP_ALIGN.LEFT):
        tf=cell.text_frame; tf.word_wrap=True; tf.clear()
        p=tf.paragraphs[0]; p.alignment=align
        r=p.add_run(); r.text=str(text)
        r.font.size=Pt(sz); r.font.bold=bold; r.font.name='Calibri'
        if color: r.font.color.rgb=color

    def bar_fill(cell, real, sc_str):
        tc=cell._tc
        NS='http://schemas.openxmlformats.org/drawingml/2006/main'
        tcPr=tc.find(f'{{{NS}}}tcPr')
        if tcPr is None:
            tcPr=etree.Element(f'{{{NS}}}tcPr'); tc.insert(0,tcPr)
        for f in list(tcPr):
            if 'Fill' in f.tag or 'noFill' in f.tag: tcPr.remove(f)
        pct=min(float(real or 0),100)
        h=sem_hex.get(sc_str,'888888')
        if pct>=100:
            fill=etree.SubElement(tcPr,f'{{{NS}}}solidFill')
            etree.SubElement(fill,f'{{{NS}}}srgbClr').set('val',h)
        else:
            stop=int(pct*1000); s3=min(stop+500,100000)
            gf=etree.SubElement(tcPr,f'{{{NS}}}gradFill')
            gl=etree.SubElement(gf,f'{{{NS}}}gsLst')
            for pos,col in [(0,h),(stop,h),(s3,'E2E2E2'),(100000,'E2E2E2')]:
                g=etree.SubElement(gl,f'{{{NS}}}gs'); g.set('pos',str(pos))
                etree.SubElement(g,f'{{{NS}}}srgbClr').set('val',col)
            lin=etree.SubElement(gf,f'{{{NS}}}lin'); lin.set('ang','0'); lin.set('scaled','0')

    # SLIDE 1
    s1=prs.slides[0]
    for shape in s1.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if 'semanal del' in run.text:
                        run.text=f'Cierre semanal del {rango}'
                        shape.width=I(7.5); shape.left=I(1.7)

    # SLIDE 2
    s2=prs.slides[1]
    for shape in list(s2.shapes):
        if hasattr(shape,'image'):
            shape._element.getparent().remove(shape._element)

    tot=len(datos); ter=sum(1 for p in datos if sc(p)=='Terminado')
    ati=sum(1 for p in datos if sc(p)=='A tiempo'); rie=sum(1 for p in datos if sc(p)=='Riesgo de atraso')
    atr=sum(1 for p in datos if sc(p)=='Atrasado')
    pcr=sum(1 for p in datos if p['real']>=90 and sc(p)!='Terminado')
    stb=sum(1 for p in datos if sc(p)=='Stand By')
    kpis=[('Total de proyectos',tot),('Terminado',ter),('A tiempo',ati),('En riesgo\nde atraso',rie),('Atrasados',atr),('Por cerrar',pcr),('Stand By',stb)]
    kw=1.345; kh=0.62; ky=0.62; kg=0.055; ksx=0.08
    for ii,(lbl,val) in enumerate(kpis):
        x=ksx+ii*(kw+kg)
        bx=s2.shapes.add_shape(1,I(x),I(ky),I(kw),I(kh))
        bx.fill.solid(); bx.fill.fore_color.rgb=WHITE
        bx.line.color.rgb=GOLD; bx.line.width=Pt(1.2)
        vt=s2.shapes.add_textbox(I(x),I(ky+0.03),I(kw),I(0.32))
        vp=vt.text_frame.paragraphs[0]; vp.alignment=PP_ALIGN.CENTER
        vr=vp.add_run(); vr.text=str(val); vr.font.size=Pt(20); vr.font.bold=True; vr.font.name='Calibri'; vr.font.color.rgb=NAVY
        lt=s2.shapes.add_textbox(I(x),I(ky+0.37),I(kw),I(0.24))
        lt.text_frame.word_wrap=True
        lp=lt.text_frame.paragraphs[0]; lp.alignment=PP_ALIGN.CENTER
        lr=lp.add_run(); lr.text=lbl; lr.font.size=Pt(6.5); lr.font.name='Calibri'; lr.font.color.rgb=GRAY
    tx=s2.shapes.add_textbox(I(1.75),I(0.04),I(6.5),I(0.34))
    p_tx=tx.text_frame.paragraphs[0]; p_tx.alignment=PP_ALIGN.CENTER
    r_tx=p_tx.add_run(); r_tx.text='ÁREA DE IMPLEMENTACIÓN'
    r_tx.font.size=Pt(15); r_tx.font.bold=True; r_tx.font.name='Calibri'; r_tx.font.color.rgb=NAVY
    tx2=s2.shapes.add_textbox(I(0),I(0.38),I(10),I(0.18))
    p2=tx2.text_frame.paragraphs[0]; p2.alignment=PP_ALIGN.CENTER
    r2=p2.add_run(); r2.text=f'Cierre semanal del {rango}'
    r2.font.size=Pt(8.5); r2.font.name='Calibri'; r2.font.color.rgb=GRAY
    colW=[I(0.83),I(2.70),I(0.73),I(0.73),I(0.58),I(0.65),I(2.42),I(0.92)]
    TBL_X=0.08; TBL_Y=1.30; HDR_H=0.22; ROW_H=0.215
    tblW=sum(colW); nrows=len(datos)+1
    tblS=s2.shapes.add_table(nrows,8,I(TBL_X),I(TBL_Y),tblW,I(HDR_H+len(datos)*ROW_H))
    tbl=tblS.table
    for ci,w in enumerate(colW): tbl.columns[ci].width=w
    tbl.rows[0].height=I(HDR_H)
    hdrs=['TIPO','PROYECTO','INICIO','FIN','% REAL','% ESPERADO','BARRA DE AVANCE','SEMÁFORO']
    for ci,h in enumerate(hdrs):
        cell=tbl.cell(0,ci)
        cell_txt(cell,h,sz=7,bold=True,color=NAVY,align=PP_ALIGN.CENTER)
        cell.fill.solid(); cell.fill.fore_color.rgb=GOLD
    for ri,p in enumerate(datos):
        tbl.rows[ri+1].height=I(ROW_H)
        s=sc(p); sc_c=sem_col.get(s,RGBColor(0x88,0x88,0x88))
        bg=RGBColor(0xF5,0xF7,0xFA) if ri%2==0 else WHITE
        vals=[p['tipo'],p['nombre'],p['inicio'],p['fin'],f"{p['real']}%",f"{p['esp']}%",f"{p['real']}%",sem_txt.get(s,s)]
        aligns=[PP_ALIGN.CENTER,PP_ALIGN.LEFT,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.RIGHT,PP_ALIGN.CENTER]
        colors=[NAVY,NAVY,GRAY,GRAY,NAVY,GRAY,sc_c,sc_c]; bolds=[True,False,False,False,True,False,True,True]
        for ci2,(v,al,co,bo) in enumerate(zip(vals,aligns,colors,bolds)):
            cell=tbl.cell(ri+1,ci2); cell_txt(cell,v,sz=7,bold=bo,color=co,align=al)
            if ci2==6: bar_fill(cell,p['real'],s)
            else: cell.fill.solid(); cell.fill.fore_color.rgb=bg

    # Guardar PPTX con slide2 usando slideLayout17 (mismo fondo que slide3)
    buf=io.BytesIO(); prs.save(buf); buf.seek(0)
    pptx_bytes_raw=buf.read()
    import zipfile as _zf_s2, io as _io_s2
    z_in_s2=_zf_s2.ZipFile(_io_s2.BytesIO(pptx_bytes_raw))
    buf_s2=_io_s2.BytesIO()
    z_out_s2=_zf_s2.ZipFile(buf_s2,'w',_zf_s2.ZIP_DEFLATED)
    for item in z_in_s2.infolist():
        d=z_in_s2.read(item.filename)
        if item.filename=='ppt/slides/_rels/slide2.xml.rels':
            d=d.replace(b'slideLayouts/slideLayout12.xml',b'slideLayouts/slideLayout17.xml')
        elif item.filename=='ppt/slides/slide2.xml':
            s2_str=d.decode('utf-8')
            shapes=re.findall(r'<p:sp>[\s\S]*?</p:sp>',s2_str)
            for sh in shapes:
                if 'ctrTitle' in sh:
                    s2_str=s2_str.replace(sh,'',1)
                    break
            d=s2_str.encode('utf-8')
        z_out_s2.writestr(item,d)
    z_out_s2.close()
    pptx_bytes=buf_s2.getvalue()

    # Subir al repo via API
    api_url='https://api.github.com/repos/DTI-Dashboard/Proyectos/contents/cierre_base.pptx'
    headers_gh={'Authorization':f'token {GH_TOKEN}','Content-Type':'application/json'}
    try:
        req_get=_ur.Request(api_url,headers={'Authorization':f'token {GH_TOKEN}'})
        with _ur.urlopen(req_get) as rg: existing_sha=json.loads(rg.read())['sha']
    except: existing_sha=None
    payload={'message':'reporte: actualizar cierre_base.pptx (bajo demanda)','content':base64.b64encode(pptx_bytes).decode()}
    if existing_sha: payload['sha']=existing_sha
    req_put=_ur.Request(api_url,data=json.dumps(payload).encode(),method='PUT',headers=headers_gh)
    with _ur.urlopen(req_put) as rp: json.loads(rp.read())
    print(f"✅ cierre_base.pptx generado bajo demanda: {len(datos)} proyectos, rango: {rango}")

    with open(_error_file, "w") as _f:
        _f.write("OK")

except Exception as _e:
    _err = _tb.format_exc()
    print("❌ EXCEPCIÓN:", _err)
    with open(_error_file, "w") as _f:
        _f.write(f"ERROR: {type(_e).__name__}: {_e}\n\n{_err}")
    _sys.exit(1)
