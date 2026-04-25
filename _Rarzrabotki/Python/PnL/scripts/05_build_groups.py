"""Step 05: build groups list from Excel + config → 03_groups_pl.json."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.names_cleaner import normalize


def main():
    # Collect groups actually seen in Excel
    data = json.loads((config.JSON_DIR / "01_raw_sheets.json").read_text(encoding="utf-8"))
    seen = {}
    for period in data:
        for sh in period["sheets"]:
            for r in sh["rows"]:
                g = r.get("group")
                if g:
                    seen.setdefault(g, 0)
                    seen[g] += 1

    # Also include all configured groups even if never seen
    all_groups = dict(seen)
    for g in config.PL_GROUPS:
        all_groups.setdefault(g, 0)

    # De-duplicate by normalized name (variants like "проданой" / "проданной")
    dedup = {}
    for name, cnt in all_groups.items():
        key = normalize(name)
        if key not in dedup or cnt > dedup[key]["occurrences"]:
            dedup[key] = {"name": name, "occurrences": cnt}

    items = []
    for entry in dedup.values():
        items.append({
            "name": entry["name"],
            "occurrences": entry["occurrences"],
            "uuid": "",
        })

    items.sort(key=lambda x: (-x["occurrences"], x["name"]))

    dst = config.JSON_DIR / "03_groups_pl.json"
    dst.write_text(json.dumps({"total": len(items), "groups": items}, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    print(f"Wrote {dst}  (groups: {len(items)})")
    for it in items:
        print(f"  {it['occurrences']:4}  {it['name']}")


if __name__ == "__main__":
    main()
