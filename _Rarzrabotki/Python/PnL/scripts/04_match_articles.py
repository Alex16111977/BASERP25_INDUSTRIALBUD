"""Step 04: fuzzy match PL articles ↔ ДДС → 05_mapping_pl_to_dds.json (1:N)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from rapidfuzz import fuzz
import config
from utils.names_cleaner import normalize


def best_scores(pl_norm: str, pl_raw: str, dds_items):
    """Return list of (score, dds_item) sorted desc."""
    scored = []
    for d in dds_items:
        n = d["name_norm"]
        r1 = fuzz.ratio(pl_norm, n)
        r2 = fuzz.token_set_ratio(pl_norm, n)
        r3 = fuzz.partial_ratio(pl_norm, n)
        score = max(r1, r2, r3)
        scored.append((score, r1, r2, r3, d))
    scored.sort(key=lambda x: -x[0])
    return scored


def main():
    pl = json.loads((config.JSON_DIR / "02_unique_articles.json").read_text(encoding="utf-8"))["articles"]
    dds = json.loads((config.JSON_DIR / "04_dds_articles.json").read_text(encoding="utf-8"))["items"]

    dds_idx = [dict(d, name_norm=normalize(d["name"])) for d in dds if not d["is_group"]]
    auto_t = config.FUZZY_AUTO_THRESHOLD
    cand_t = config.FUZZY_CANDIDATE_THRESHOLD

    result = []
    unmapped = []
    exact_count = fuzzy_count = none_count = 0

    for a in pl:
        pl_norm = normalize(a["name"])
        scored = best_scores(pl_norm, a["name"], dds_idx)

        dds_items = []
        candidates = []
        match_type = "none"

        # exact (normalized)
        exact_hits = [d for (sc, r1, r2, r3, d) in scored if d["name_norm"] == pl_norm]
        if exact_hits:
            match_type = "exact"
            exact_count += 1
            for d in exact_hits[:1]:  # one exact is enough for auto
                dds_items.append({"uuid": d["uuid"], "name": d["name"], "score": 100, "auto": True})
        else:
            top = scored[:10]
            auto_added = False
            for (sc, r1, r2, r3, d) in top:
                if sc >= auto_t and not auto_added:
                    dds_items.append({"uuid": d["uuid"], "name": d["name"], "score": sc, "auto": True})
                    match_type = "fuzzy"
                    auto_added = True
                elif sc >= cand_t:
                    candidates.append({"uuid": d["uuid"], "name": d["name"], "score": sc})
            if auto_added:
                fuzzy_count += 1
            else:
                none_count += 1

        if not dds_items:
            unmapped.append(a["name"])

        result.append({
            "pl_name": a["name"],
            "pl_group": a.get("group", ""),
            "dds_items": dds_items,
            "candidates": candidates[:5],
            "match_type": match_type,
        })

    # Применяем ручные override из config.MANUAL_DDS_OVERRIDES (добавлено 2026-04-17).
    overrides = getattr(config, "MANUAL_DDS_OVERRIDES", {})
    manual_applied = 0
    existing_by_name = {r["pl_name"]: r for r in result}
    for pl_name, (dds_uuid, dds_name) in overrides.items():
        if pl_name in existing_by_name:
            r = existing_by_name[pl_name]
            r["dds_items"] = [{"uuid": dds_uuid, "name": dds_name, "score": 100, "auto": True, "manual": True}]
            r["match_type"] = "manual_override"
            r["candidates"] = []
            if pl_name in unmapped:
                unmapped.remove(pl_name)
            manual_applied += 1

    dst = config.JSON_DIR / "05_mapping_pl_to_dds.json"
    dst.write_text(json.dumps({"total": len(result), "mappings": result}, ensure_ascii=False, indent=2),
                   encoding="utf-8")

    rep = config.REPORTS_DIR / "unmapped_articles.txt"
    config.REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    rep.write_text("\n".join(unmapped), encoding="utf-8")

    print(f"Wrote {dst}")
    print(f"  exact={exact_count}  fuzzy={fuzzy_count}  none={none_count}  "
          f"manual_override={manual_applied}")
    print(f"Unmapped list: {rep}")


if __name__ == "__main__":
    main()
