"""Step 02: unique PL article names (universal across all files) → 02_unique_articles.json.

Dedupes by normalized name (casefold + collapsed whitespace), keeps most-frequent variant
as canonical; alternative spellings stored in `variants`. Also flags possible fuzzy dupes.
"""
import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from rapidfuzz import fuzz
import config
from utils.names_cleaner import normalize


def main():
    src = config.JSON_DIR / "01_raw_sheets.json"
    data = json.loads(src.read_text(encoding="utf-8"))

    # Алиасы Excel-имён → каноничное имя (исправление опечаток).
    aliases = {k: v for k, v in getattr(config, "ARTICLE_NAME_ALIASES", {}).items()}

    # 1) Collect per raw name (с применением алиасов)
    raw = defaultdict(lambda: {"occurrences": 0, "sample_sheets": set(), "groups": set(), "periods": set()})
    for period in data:
        for sh in period["sheets"]:
            for r in sh["rows"]:
                name = aliases.get(r["article"], r["article"])
                raw[name]["occurrences"] += 1
                if len(raw[name]["sample_sheets"]) < 5:
                    raw[name]["sample_sheets"].add(sh["sheet_name"])
                if r.get("group"):
                    raw[name]["groups"].add(r["group"])
                raw[name]["periods"].add(period["label"])

    # 2) Dedupe by normalized form; canonical = most frequent variant
    clusters = defaultdict(list)
    for name, st in raw.items():
        clusters[normalize(name)].append((name, st))

    merged = []
    for key, variants in clusters.items():
        variants.sort(key=lambda x: -x[1]["occurrences"])
        canonical_name, canon_st = variants[0]
        total_occ = sum(v[1]["occurrences"] for v in variants)
        all_groups = set()
        all_sheets = set()
        all_periods = set()
        for _, s in variants:
            all_groups |= s["groups"]
            all_sheets |= s["sample_sheets"]
            all_periods |= s["periods"]
        primary_group = sorted(all_groups)[0] if all_groups else ""
        merged.append({
            "name": canonical_name,
            "occurrences": total_occ,
            "variants": [v[0] for v in variants if v[0] != canonical_name],
            "periods": sorted(all_periods),
            "sample_sheets": sorted(all_sheets)[:5],
            "groups_seen": sorted(all_groups),
            "group": primary_group,
            "type_statii": config.determine_type(canonical_name, primary_group),
            "uuid": "",
            "possible_dupe_of": [],
        })

    # 3) Detect possible fuzzy duplicates (high threshold to avoid false positives)
    DUPE_THRESHOLD = 92
    norms = [(i, normalize(m["name"])) for i, m in enumerate(merged)]
    for i, ni in norms:
        for j, nj in norms:
            if j <= i:
                continue
            if fuzz.ratio(ni, nj) >= DUPE_THRESHOLD:
                merged[i]["possible_dupe_of"].append(merged[j]["name"])
                merged[j]["possible_dupe_of"].append(merged[i]["name"])

    merged.sort(key=lambda x: -x["occurrences"])
    articles = merged

    dst = config.JSON_DIR / "02_unique_articles.json"
    dst.write_text(json.dumps({"total": len(articles), "articles": articles}, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    print(f"Wrote {dst}  (unique articles: {len(articles)})")


if __name__ == "__main__":
    main()
