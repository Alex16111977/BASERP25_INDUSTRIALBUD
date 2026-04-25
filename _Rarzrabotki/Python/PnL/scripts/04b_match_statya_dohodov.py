"""Step 04b: fuzzy match PL (only income articles) ↔ ПВХ.СтатьиДоходов → 10_mapping_statya_dohodov.json.

Применяется к статьям из групп с ключевыми словами "доход"/"выручка".
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from rapidfuzz import fuzz
import config
from utils.com_connect import connect_erp, uuid_str
from utils.names_cleaner import normalize

INCOME_KEYWORDS = ["доход", "доходы", "доходи", "выручк", "дохід"]


def is_income(a):
    name_l = a["name"].lower()
    grp_l = (a.get("group") or "").lower()
    for k in INCOME_KEYWORDS:
        if k in name_l or k in grp_l:
            return True
    return False


def main():
    arts = json.loads(
        (config.JSON_DIR / "02_unique_articles.json").read_text(encoding="utf-8")
    )["articles"]
    income_arts = [a for a in arts if is_income(a)]
    print(f"PL income articles: {len(income_arts)}")

    # Pull СтатьиДоходов via COM
    conn = connect_erp()
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ СД.Ссылка, СД.Наименование ИЗ ПланВидовХарактеристик.СтатьиДоходов КАК СД "
        "ГДЕ НЕ СД.ПометкаУдаления"
    )
    tz = q.Выполнить().Выгрузить()
    sd_idx = []
    for i in range(tz.Количество()):
        r = tz.Получить(i)
        sd_idx.append(
            {
                "uuid": uuid_str(conn, r.Ссылка),
                "name": str(r.Наименование),
                "name_norm": normalize(str(r.Наименование)),
            }
        )
    print(f"ПВХ.СтатьиДоходов: {len(sd_idx)}")

    result = []
    unmapped = []
    auto_t = config.FUZZY_AUTO_THRESHOLD
    cand_t = config.FUZZY_CANDIDATE_THRESHOLD

    for a in income_arts:
        nm = a["name"]
        pl_norm = normalize(nm)
        scored = []
        for s in sd_idx:
            r1 = fuzz.ratio(pl_norm, s["name_norm"])
            r2 = fuzz.token_set_ratio(pl_norm, s["name_norm"])
            r3 = fuzz.partial_ratio(pl_norm, s["name_norm"])
            scored.append((max(r1, r2, r3), s))
        scored.sort(key=lambda x: -x[0])

        match_uuid = ""
        match_name = ""
        match_type = "none"
        candidates = []

        exact = [s for (sc, s) in scored if s["name_norm"] == pl_norm]
        if exact:
            match_uuid = exact[0]["uuid"]
            match_name = exact[0]["name"]
            match_type = "exact"
        else:
            top_score, top = scored[0]
            if top_score >= auto_t:
                match_uuid = top["uuid"]
                match_name = top["name"]
                match_type = "fuzzy"
            for sc, s in scored[:5]:
                if sc >= cand_t:
                    candidates.append(
                        {"uuid": s["uuid"], "name": s["name"], "score": sc}
                    )

        if not match_uuid:
            unmapped.append(nm)

        result.append(
            {
                "pl_name": nm,
                "statya_dohodov_uuid": match_uuid,
                "statya_dohodov_name": match_name,
                "match_type": match_type,
                "candidates": candidates,
            }
        )

    # Применяем ручные override из config.MANUAL_STATYA_DOHODOV_OVERRIDES (2026-04-17).
    overrides = getattr(config, "MANUAL_STATYA_DOHODOV_OVERRIDES", {})
    manual_applied = 0
    existing_by_name = {r["pl_name"]: r for r in result}
    for pl_name, (sd_uuid, sd_name) in overrides.items():
        if pl_name in existing_by_name:
            r = existing_by_name[pl_name]
            r["statya_dohodov_uuid"] = sd_uuid
            r["statya_dohodov_name"] = sd_name
            r["match_type"] = "manual_override"
            r["candidates"] = []
            if pl_name in unmapped:
                unmapped.remove(pl_name)
            manual_applied += 1
        else:
            result.append({
                "pl_name": pl_name,
                "statya_dohodov_uuid": sd_uuid,
                "statya_dohodov_name": sd_name,
                "match_type": "manual_override",
                "candidates": [],
            })
            manual_applied += 1

    dst = config.JSON_DIR / "10_mapping_statya_dohodov.json"
    dst.write_text(
        json.dumps({"total": len(result), "mappings": result}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    config.REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    (config.REPORTS_DIR / "unmapped_statya_dohodov.txt").write_text(
        "\n".join(unmapped), encoding="utf-8"
    )
    print(f"Wrote {dst}  manual_override={manual_applied}")
    for m in result:
        print(
            f"  {m['match_type']:15} {m['pl_name']!r:55} -> {m['statya_dohodov_name']!r}"
        )


if __name__ == "__main__":
    main()
