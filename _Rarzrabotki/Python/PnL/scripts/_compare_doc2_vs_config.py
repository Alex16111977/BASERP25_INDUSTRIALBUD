"""Сравнить эталонный документ №000000002 (финансист) с текущим config.py + 07_mapping.
Вывести список расхождений: какие sheet_name надо добавить в MANUAL_SHEET_TO_STRUCT_OVERRIDES
или обновить (другой struct_uuid / direction_uuid / include_children)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config


def main():
    ref = json.loads(
        (config.JSON_DIR / "_doc2_reference_mapping.json").read_text(encoding="utf-8")
    )["rows"]
    mapping = json.loads(
        (config.JSON_DIR / "07_mapping_sheet_to_struct.json").read_text(encoding="utf-8")
    )["mappings"]
    mapping_by_sheet = {m["sheet_name"]: m for m in mapping}
    overrides = config.MANUAL_SHEET_TO_STRUCT_OVERRIDES

    diffs = []
    new_entries = []  # для config.py

    for r in ref:
        sn = r["sheet_name"]
        cur = mapping_by_sheet.get(sn)
        ov = overrides.get(sn)

        ref_struct = (r["struct_uuid"] or "").lower()
        ref_dir = (r["direction_uuid"] or "").lower()
        ref_incch = bool(r["include_children"])

        cur_struct = (cur.get("struct_uuid") if cur else "" or "").lower()
        cur_dir = (cur.get("direction_uuid") if cur else "" or "").lower()
        cur_incch = bool(cur.get("include_children", False)) if cur else False

        changes = []
        if ref_struct != cur_struct:
            changes.append(f"struct: {cur_struct or '-'} -> {ref_struct or '-'}")
        if ref_dir != cur_dir:
            changes.append(f"dir: {cur_dir or '-'} -> {ref_dir or '-'}")
        if ref_incch != cur_incch:
            changes.append(f"incch: {cur_incch} -> {ref_incch}")

        if changes:
            diffs.append({
                "sheet": sn,
                "struct_name_ref": r["struct_name"],
                "direction_name_ref": r["direction_name"],
                "include_children_ref": ref_incch,
                "changes": changes,
                "was_in_config": ov is not None,
            })

            # Предложение для config.py
            if ref_struct or ref_dir or ref_incch:
                new_entries.append({
                    "sheet_name": sn,
                    "struct_uuid": r["struct_uuid"] or "",
                    "struct_name": r["struct_name"] or "",
                    "direction_uuid": r["direction_uuid"] or "",
                    "direction_name": r["direction_name"] or "",
                    "include_children": ref_incch,
                })

    out = config.JSON_DIR / "_doc2_vs_config_diff.json"
    out.write_text(
        json.dumps({"diffs": diffs, "suggested_overrides": new_entries},
                   ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"=== Сравнение документа #2 vs текущий config.py + 07_mapping ===")
    print(f"Всего строк в документе: {len(ref)}")
    print(f"Расхождений: {len(diffs)}")
    print()
    for d in diffs:
        mark = "[в config]" if d["was_in_config"] else "[НОВОЕ]"
        print(f"  {mark} {d['sheet']}")
        print(f"      target: struct={d['struct_name_ref']}  dir={d['direction_name_ref']}  incch={d['include_children_ref']}")
        for c in d["changes"]:
            print(f"      change: {c}")
        print()

    print(f"Full diff written: {out}")


if __name__ == "__main__":
    main()
