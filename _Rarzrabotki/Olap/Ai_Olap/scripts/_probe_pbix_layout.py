import json, shutil, zipfile, os, tempfile, re

SRC = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\PowerBi\PL.pbix"
tmp = os.path.join(tempfile.gettempdir(), "_pl_copy.pbix")
shutil.copy2(SRC, tmp)

z = zipfile.ZipFile(tmp)

def readj(name):
    raw = z.read(name)
    for enc in ("utf-8", "utf-16-le", "utf-16"):
        try:
            return json.loads(raw.decode(enc))
        except Exception:
            continue
    return None

names = z.namelist()

# page folder -> displayName + order
pages_meta = readj("Report/definition/pages/pages.json")
active = pages_meta.get("activePageName") if pages_meta else None
order = pages_meta.get("pageOrder", []) if pages_meta else []
print("activePage:", active)

def page_folder(pj):
    return pj  # name in pages.json equals folder name

# map folder -> displayName
folder_disp = {}
for n in names:
    m = re.match(r"Report/definition/pages/([^/]+)/page\.json$", n)
    if m:
        pj = readj(n)
        folder_disp[m.group(1)] = pj.get("displayName") if pj else None

print("\n=== PAGES (in order) ===")
for f in order:
    print(f"  {folder_disp.get(f, '?')!r:45}  folder={f}{'   <== ACTIVE' if f==active else ''}")

KW = re.compile(r"(year_month|Period|date_|Calendar|Месяц|RelativeDate|2025|2026|year_)", re.I)
DATEISH = re.compile(r"(Calendar|Period|year_month|date_|Месяц)", re.I)

def find_fields(obj):
    """collect Entity/Property pairs in a visual/filter json"""
    out = set()
    def w(o):
        if isinstance(o, dict):
            ent = o.get("Entity");
            for k,v in o.items():
                if isinstance(v, dict) and "Property" in v and isinstance(v.get("Expression"),dict):
                    pass
                w(v)
            # column ref pattern {"Column":{"Expression":{"SourceRef":{"Entity":..}},"Property":..}}
        elif isinstance(o, list):
            for v in o: w(v)
    # simpler: regex on raw
    return out

print("\n=== SCAN: page-level filters + slicers + any date-ish filter w/ selection ===")
for n in sorted(names):
    if not n.startswith("Report/definition/pages/"):
        continue
    if not n.endswith(".json"):
        continue
    raw = z.read(n)
    try:
        txt = raw.decode("utf-8")
    except Exception:
        txt = raw.decode("utf-16-le", "ignore")
    # figure out page label
    pm = re.search(r"pages/([^/]+)/", n)
    plabel = folder_disp.get(pm.group(1), "?") if pm else "?"
    low = txt.lower()
    is_visual = n.endswith("visual.json")
    is_slicer = '"slicer"' in low or '"visualtype":"slicer"' in low
    has_filter_sel = ('"filterconfig"' in low) and ('"filter"' in low or '"conditions"' in low or 'values' in low)
    mentions_date = bool(DATEISH.search(txt))

    if is_slicer or (mentions_date and ('filter' in low)):
        # extract entity.property refs
        refs = re.findall(r'"Entity"\s*:\s*"([^"]+)"', txt)
        props = re.findall(r'"Property"\s*:\s*"([^"]+)"', txt)
        # extract any literal date/string values present
        lits = re.findall(r'"(20\d\d[-./]\d\d[-./]\d\d[^"]*)"', txt)
        datelits = re.findall(r'datetime\'([^\']+)\'', txt, re.I)
        rel = re.findall(r'"RelativeDate[^}]{0,120}', txt)
        tag = "SLICER" if is_slicer else "FILTER"
        if DATEISH.search(" ".join(refs+props)) or is_slicer or rel or datelits or any('20' in l for l in lits):
            print(f"\n--- [{tag}] page={plabel!r}  {n.split('/')[-1]} ---")
            print("   entities:", sorted(set(refs)))
            print("   props   :", sorted(set(props)))
            if datelits: print("   datetimeLits:", datelits[:10])
            if lits: print("   dateStrLits:", lits[:10])
            if rel: print("   relativeDate:", rel[:3])
