import json, shutil, zipfile, os, tempfile, re

SRC = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\PowerBi\PL.pbix"
tmp = os.path.join(tempfile.gettempdir(), "_pl_copy.pbix")
shutil.copy2(SRC, tmp)
z = zipfile.ZipFile(tmp)
def readj(name):
    raw = z.read(name)
    for enc in ("utf-8","utf-16-le","utf-16"):
        try: return json.loads(raw.decode(enc))
        except Exception: continue

ACTIVE = "7045c2e030dcd0a4d264"
n = f"Report/definition/pages/{ACTIVE}/visuals/ffdd5fcc216ac35bc2e8/visual.json"
j = readj(n)

def lits(obj):
    out=[]
    def w(o):
        if isinstance(o,dict):
            if "Literal" in o and isinstance(o["Literal"],dict):
                out.append(o["Literal"].get("Value"))
            for v in o.values(): w(v)
        elif isinstance(o,list):
            for v in o: w(v)
    w(obj); return out

for f in (j.get("filterConfig",{}) or {}).get("filters",[]):
    blob = json.dumps(f, ensure_ascii=False)
    if re.search(r"(year_month|month_name|date_|Period)", blob):
        prop = re.findall(r'"Property":"([^"]+)"', blob)
        sel = lits(f.get("filter",{}))
        print(f"\nFILTER on {prop} type={f.get('type')} howCreated={f.get('howCreated')}")
        print("  selected values:", sel)
