# IFC TAKEOFF WEB – BIM STYLE
from flask import Flask, request, render_template_string, send_file, jsonify
import os, zipfile
import pandas as pd
import ifcopenshell
import ifcopenshell.util.element as Element
import ifcopenshell.geom
from openpyxl import load_workbook
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

progress_state = {"current":0, "total":1, "text":"idle"}

def set_progress(c,t,txt):
    progress_state["current"]=c
    progress_state["total"]=t
    progress_state["text"]=txt

HTML = """
<!DOCTYPE html>
<html>
<head>
<style>
body{
  background:#0a0e0a;
  color:#39ff14;
  font-family:'Courier New', monospace;
  padding:20px;
}
.panel{
  border:2px solid #39ff14;
  padding:15px;
  max-width:800px;
  box-shadow:0 0 12px #39ff14;
}
pre.logo{ color:#39ff14; line-height:1.1; }
button,input[type=submit]{
  background:#0a0e0a;
  color:#39ff14;
  border:2px solid #39ff14;
  padding:6px 12px;
}
progress{width:100%;height:18px;}
</style>
</head>
<body>
<div class='panel'>
<pre class='logo'>
   ____  _  __  __      ____ ___ __  __ 
  | __ )(_)|  \/  |    |  _ \_ _|  \/  |
  |  _ \| || |\/| |____| |_) | || |\/| |
  | |_) | || |  | |____|  __/| || |  | |
  |____/|_||_|  |_|    |_|  |___|_|  |_|

   +351 IFC TAKEOFF TERMINAL
   BIM → PSets → WBS → BOQ PIPELINE
   MODEL‑BASED QUANTIFICATION ENGINE
</pre>

<form method="post" enctype="multipart/form-data">
IFC Files:<br>
<input type="file" name="ifc" multiple required><br><br>

WBS Excel:<br>
<input type="file" name="excel" required><br><br>

<progress id="bar" value="0" max="100"></progress>
<span id="txt"></span><br><br>

<input type="submit" value="RUN TAKEOFF">
</form>
</div>

<script>
setInterval(async ()=>{
 const r = await fetch('/progress');
 const j = await r.json();
 const p = j.total==0?0:Math.round(j.current/j.total*100);
 document.getElementById('bar').value=p;
 document.getElementById('txt').innerText=j.text;
},700);
</script>

</body></html>
"""

settings = ifcopenshell.geom.settings()
settings.set(settings.USE_WORLD_COORDS, True)

def txt(v): return str(v).strip().lower() if v else ""

def get_pset_value(el,ps,p):
    try:
        psets=Element.get_psets(el)
        if psets and ps in psets:
            return psets[ps].get(p)
    except: pass
    return None

def get_bbox_length(el):
    try:
        s=ifcopenshell.geom.create_shape(settings,el)
        v=s.geometry.verts
        xs,ys,zs=v[0::3],v[1::3],v[2::3]
        return max(max(xs)-min(xs),max(ys)-min(ys),max(zs)-min(zs))
    except: return None

def get_quantity(el,uom):
    uom=txt(uom)
    if uom=="un": return 1

    if uom=="m":
        L=get_pset_value(el,"BaseQuantities","Length")
        if L: return float(L)
        g=get_bbox_length(el)
        if g: return float(g)

    if uom=="m3":
        v=get_pset_value(el,"BaseQuantities","NetVolume")
        if v: return float(v)

    return 0

def run_one(ifc_path, excel_path, output_path):
    model=ifcopenshell.open(ifc_path)
    elements=model.by_type("IfcElement")

    df=pd.read_excel(excel_path,dtype=str)

    out=[]; val=[]
    total_el=len(elements)
    step=0

    for _,row in df.iterrows():
        wbs=txt(row["WBS"])
        desc=row["Description"]
        uom=row["UN"]
        cost=float(str(row["Cost"]).replace(",","."))

        by={}

        for el in elements:
            step+=1
            set_progress(step,total_el,os.path.basename(ifc_path))

            el_wbs=txt(get_pset_value(el,"Pset_+351","WBS"))
            piso=get_pset_value(el,"Pset_+351","Piso")

            if not piso:
                val.append({"GlobalId":el.GlobalId,"IfcClass":el.is_a(),"WBS":el_wbs,"Reason":"Missing Piso"})
                piso="NO PISO"

            if el_wbs!=wbs: continue

            q=get_quantity(el,uom)
            by[piso]=by.get(piso,0)+q

        tq=round(sum(by.values()),3)
        tc=round(tq*cost,3)

        out.append([row["WBS"],desc+" - TOTAL",uom,"",tq,cost,tc])

    wb=load_workbook(excel_path)
    ws=wb.active
    ws.delete_rows(2,ws.max_row)

    for r,data in enumerate(out,2):
        for c,v in enumerate(data,1): ws.cell(row=r,column=c).value=v

    wb.save(output_path)

@app.route('/progress')
def prog(): return jsonify(progress_state)

@app.route('/',methods=['GET','POST'])
def index():
    if request.method=='POST':
        excel=request.files['excel']
        excel_path=os.path.join(UPLOAD_FOLDER,secure_filename(excel.filename))
        excel.save(excel_path)

        results=[]
        for ifc in request.files.getlist('ifc'):
            p=os.path.join(UPLOAD_FOLDER,secure_filename(ifc.filename))
            ifc.save(p)

            o=os.path.join(OUTPUT_FOLDER,f"result_{secure_filename(ifc.filename)}.xlsx")
            run_one(p,excel_path,o)
            results.append(o)

        if len(results)==1:
            return send_file(results[0],as_attachment=True)

        z=os.path.join(OUTPUT_FOLDER,"batch_results.zip")
        with zipfile.ZipFile(z,'w') as zz:
            for r in results: zz.write(r,os.path.basename(r))
        return send_file(z,as_attachment=True)

    return render_template_string(HTML)

if __name__=='__main__':
    app.run(debug=True)
