import json, shutil, zipfile, os, tempfile, re

SRC = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\PowerBi\PL.pbix"
tmp = os.path.join(tempfile.gettempdir(), "_pl_copy.pbix")
shutil.copy2(SRC, tmp)
z = zipfile.ZipFile(tmp)
def readtxt(name):
    raw = z.read(name)
    for enc in ("utf-8","utf-16-le","utf-16"):
        try: return raw.decode(enc)
        except Exception: continue

ACTIVE = "7045c2e030dcd0a4d264"
names = z.namelist()

# Dump full slicer + the pivot visual's column/row fields & the date-ish filters verbatim
for n in sorted(names):
    m = re.match(rf"Report/definition/pages/{ACTIVE}/visuals/([^/]+)/visual\.json$", n)
    if not m: continue
    txt = readtxt(n)
    j = json.loads(txt)
    vt = j.get("visual",{}).get("visualType")
    print(f"\n########## {vt}  {m.group(1)} ##########")
    if vt == "slicer":
        # print the whole thing (slicers are small)
        print(txt[:4000])
    else:
        # pivot: show row/column/values projections + any Calendar/date filter blob
        vq = j.get("visual",{}).get("query",{})
        print("PROJECTIONS:", json.dumps(vq, ensure_ascii=False)[:1500])
        # date filters
        for f in (j.get("filterConfig",{}) or {}).get("filters",[]):
            blob = json.dumps(f, ensure_ascii=False)
            if re.search(r"(Calendar|date_|year_month|Period|month_)", blob):
                print("  DATE-FILTER:", blob[:900])
