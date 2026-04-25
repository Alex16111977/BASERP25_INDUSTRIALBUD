"""Step 07: fuzzy match Excel sheet names ↔ СтруктураПредприятия → 07_mapping_sheet_to_struct.json."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from rapidfuzz import fuzz
import config
from utils.names_cleaner import normalize


def main():
    raw = json.loads((config.JSON_DIR / "01_raw_sheets.json").read_text(encoding="utf-8"))
    struct = json.loads((config.JSON_DIR / "06_struct_predpr.json").read_text(encoding="utf-8"))["items"]
    struct_idx = [dict(s, name_norm=normalize(s["name"])) for s in struct]
    struct_by_uuid = {s["uuid"]: s for s in struct}

    sheet_names = []
    seen = set()
    for period in raw:
        for sh in period["sheets"]:
            n = sh["sheet_name"]
            if n not in seen:
                seen.add(n)
                sheet_names.append(n)

    result = []
    unmapped = []
    exact_count = fuzzy_count = none_count = 0

    for sn in sheet_names:
        sn_norm = normalize(sn)
        scored = []
        for s in struct_idx:
            r1 = fuzz.ratio(sn_norm, s["name_norm"])
            r2 = fuzz.token_set_ratio(sn_norm, s["name_norm"])
            r3 = fuzz.partial_ratio(sn_norm, s["name_norm"])
            score = max(r1, r2, r3)
            scored.append((score, s))
        scored.sort(key=lambda x: -x[0])

        exact_hits = [s for (sc, s) in scored if s["name_norm"] == sn_norm]
        struct_uuid = ""
        struct_name = ""
        match_type = "none"
        confidence = 0
        candidates = []

        if exact_hits:
            struct_uuid = exact_hits[0]["uuid"]
            struct_name = exact_hits[0]["name"]
            match_type = "exact"
            confidence = 100
            exact_count += 1
        else:
            top_score, top_item = scored[0]
            if top_score >= config.FUZZY_AUTO_THRESHOLD:
                struct_uuid = top_item["uuid"]
                struct_name = top_item["name"]
                match_type = "fuzzy"
                confidence = top_score
                fuzzy_count += 1
            else:
                match_type = "none"
                none_count += 1

            for (sc, s) in scored[:5]:
                if sc >= config.FUZZY_CANDIDATE_THRESHOLD:
                    candidates.append({"uuid": s["uuid"], "name": s["name"], "score": sc})

        if not struct_uuid:
            unmapped.append(sn)

        direction_uuid = ""
        direction_name = ""
        if struct_uuid:
            s = struct_by_uuid.get(struct_uuid)
            if s:
                direction_uuid = s.get("direction_uuid") or ""
                direction_name = s.get("direction_name") or ""

        result.append({
            "sheet_name": sn,
            "struct_uuid": struct_uuid,
            "struct_name": struct_name,
            "match_type": match_type,
            "confidence": confidence,
            "direction_uuid": direction_uuid,
            "direction_name": direction_name,
            "include_children": False,
            "candidates": candidates,
        })

    # Применяем ручные override из config.MANUAL_SHEET_TO_STRUCT_OVERRIDES
    # (добавлено 2026-04-17 для Техника→Логистика и ПРОООН→МД ПРООН).
    overrides = getattr(config, "MANUAL_SHEET_TO_STRUCT_OVERRIDES", {})
    override_applied = 0
    override_added = 0
    existing_names = {r["sheet_name"]: r for r in result}
    for sheet_name, ov in overrides.items():
        ov_dir_uuid = ov.get("direction_uuid") or ""
        ov_dir_name = ov.get("direction_name") or ""
        ov_include = bool(ov.get("include_children", False))

        # Если в override не указан direction_uuid — подтянуть из 06 по struct_uuid
        if not ov_dir_uuid:
            s = struct_by_uuid.get(ov["struct_uuid"])
            if s:
                ov_dir_uuid = s.get("direction_uuid") or ""
                ov_dir_name = s.get("direction_name") or ""

        if sheet_name in existing_names:
            r = existing_names[sheet_name]
            r["struct_uuid"] = ov["struct_uuid"]
            r["struct_name"] = ov["struct_name"]
            r["match_type"] = "manual_override"
            r["confidence"] = 100
            r["direction_uuid"] = ov_dir_uuid
            r["direction_name"] = ov_dir_name
            r["include_children"] = ov_include
            r["candidates"] = []
            r["notes"] = ov.get("reason", "")
            override_applied += 1
            if sheet_name in unmapped:
                unmapped.remove(sheet_name)
        else:
            result.append({
                "sheet_name": sheet_name,
                "struct_uuid": ov["struct_uuid"],
                "struct_name": ov["struct_name"],
                "match_type": "manual_override",
                "confidence": 100,
                "direction_uuid": ov_dir_uuid,
                "direction_name": ov_dir_name,
                "include_children": ov_include,
                "candidates": [],
                "notes": ov.get("reason", ""),
            })
            override_added += 1

    dst = config.JSON_DIR / "07_mapping_sheet_to_struct.json"
    dst.write_text(json.dumps({"total": len(result), "mappings": result}, ensure_ascii=False, indent=2),
                   encoding="utf-8")

    config.REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    rep = config.REPORTS_DIR / "unmapped_sheets.txt"
    rep.write_text("\n".join(unmapped), encoding="utf-8")

    print(f"Wrote {dst}")
    print(f"  exact={exact_count}  fuzzy={fuzzy_count}  none={none_count}  "
          f"manual_applied={override_applied}  manual_added={override_added}")
    print(f"Unmapped list: {rep}")


if __name__ == "__main__":
    main()
