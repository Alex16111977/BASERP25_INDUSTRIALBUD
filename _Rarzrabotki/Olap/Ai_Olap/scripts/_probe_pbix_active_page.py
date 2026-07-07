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
    return None

names = z.namelist()
ACTIVE = readj("Report/definition/pages/pages.json").get("activePageName")
print("ACTIVE page folder:", ACTIVE)
pj = readj(f"Report/definition/pages/{ACTIVE}/page.json")
print("ACTIVE page displayName:", pj.get("displayName"))

def dump_filterconfig(fc, indent="    "):
    """fc = filterConfig dict with 'filters' list"""
    if not fc: return
    for f in fc.get("filters", []):
        field = f.get("field", {})
        # extract entity.property
        s = json.dumps(field, ensure_ascii=False)
        ent = re.findall(r'"Entity":"([^"]+)"', s)
        prop = re.findall(r'"Property":"([^"]+)"', s)
        ftype = f.get("type")
        # selection / conditions
        fl = f.get("filter", {})
        conds = json.dumps(fl, ensure_ascii=False) if fl else ""
        # shorten but keep date/value literals
        lits = re.findall(r'datetime\'([^\']+)\'|"(20\d\d[^"]{0,8})"|"Value":"([^"]+)"', conds)
        litflat = [x for tup in lits for x in tup if x]
        hide = f.get("isHiddenInViewMode")
        print(f"{indent}field={ent}.{prop} type={ftype} hidden={hide}")
        if litflat:
            print(f"{indent}   values/dates: {litflat[:20]}")
        elif conds and conds not in ('{}',):
            print(f"{indent}   cond: {conds[:300]}")

print("\n=== PAGE-LEVEL filters (active page) ===")
dump_filterconfig(pj.get("filterConfig"))

print("\n=== VISUALS on active page ===")
for n in sorted(names):
    m = re.match(rf"Report/definition/pages/{ACTIVE}/visuals/([^/]+)/visual\.json$", n)
    if not m: continue
    v = readj(n)
    vc = v.get("visual", {}) or {}
    vtype = vc.get("visualType")
    # title
    title = ""
    try:
        objs = vc.get("objects", {})
        t = objs.get("title", [{}])[0].get("properties", {}).get("text", {})
        title = json.dumps(t, ensure_ascii=False)[:80]
    except Exception: pass
    # fields used in the query
    q = json.dumps(vc.get("query", {}), ensure_ascii=False)
    ent = re.findall(r'"Entity":"([^"]+)"', q)
    prop = re.findall(r'"Property":"([^"]+)"', q)
    print(f"\n  [{vtype}] {m.group(1)}  title={title}")
    if prop:
        print(f"     query fields: {sorted(set(zip(ent,prop)) if len(ent)==len(prop) else set(prop))}")
    # visual-level filters
    fc = v.get("filterConfig")
    if fc and fc.get("filters"):
        print("     VISUAL FILTERS:")
        dump_filterconfig(fc, indent="        ")
