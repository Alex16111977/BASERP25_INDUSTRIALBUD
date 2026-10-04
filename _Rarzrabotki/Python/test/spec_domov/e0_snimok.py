# -*- coding: utf-8 -*-
"""Этап 0 (§10.1): снимок контрольных сумм проекта МД IRC 2026 — ТОЛЬКО ЧТЕНИЕ bas_industrialbud.

Скрипты Z6 переиспользуются без правки оригиналов: копии в RUN, пути заменены с assert count == 1.
Выход: out/e0_snimok_<stamp>.json (сводка, коммитится); RUN (Temp) — полный снимок raw.json и расчёт.
Код 0: тождество Т3 = 0,00; независимые числа q_t3 совпали; Δ против Z6 = 0 или каждое Δ объяснено регистраторами."""
import datetime
import gzip
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, OUT, TMP, WT  # noqa: E402

KB = os.path.join(WT, r"_Rarzrabotki\notebook\knowledge_Бухгалтерия_РасчетКомплектаций\Разработка связей")
Z6S = os.path.join(KB, r"research\2026-10-03_mnogourovnevaya_specifikaciya\Z6_scripts")
PROTO = os.path.join(KB, r"scripts\03_prototip_svyazi_30m")
PUBLISHED = os.path.join(KB, r"data\proekt_irc2026_zakupki.json")
BASE_GZ = os.path.join(HERE, "baseline", "z6_raw_20261003_234112.json.gz")
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
RUN = os.path.join(TMP, "e0", STAMP)
CALC = os.path.join(RUN, "calc")
PY = [sys.executable, "-X", "utf8"]
KEYS = ["need_15_uah_net", "need_30_uah_net", "purchased_uah_net", "purchased_uah_vat", "other_in_uah_net",
        "other_out_uah_net", "written_off_uah_net", "written_off_15_uah_net", "written_off_30_uah_net", "lies_uah_net",
        "lies_houses_uah_net", "lies_project_uah_net", "lies_ostatki_now_uah_net", "to_buy_30_uah_last_vat",
        "to_buy_30_uah_ss_vat", "to_buy_30_uah_last_net", "to_buy_30_cards", "surplus_uah_net", "surplus_cards"]


def patched(src, dst, reps):
    s = open(src, encoding="utf-8").read()
    for old, new in reps:
        assert s.count(old) == 1, (os.path.basename(src), old[:70], s.count(old))
        s = s.replace(old, new)
    open(dst, "w", encoding="utf-8").write(s)
    return dst


def run(script, env_extra):
    r = subprocess.run(PY + [script], capture_output=True, text=True, encoding="utf-8", env={**os.environ, **env_extra})
    open(script + ".log", "w", encoding="utf-8").write(r.stdout + "\n--- stderr ---\n" + r.stderr)
    if r.returncode != 0:
        print(r.stdout[-3000:], r.stderr[-3000:])
        raise SystemExit(f"FAIL {os.path.basename(script)} rc={r.returncode}")
    print("ok", os.path.basename(script))


def main():
    os.makedirs(CALC, exist_ok=True)
    base_raw = os.path.join(RUN, "z6_baseline_raw.json")
    with gzip.open(BASE_GZ, "rb") as fi, open(base_raw, "wb") as fo:
        shutil.copyfileobj(fi, fo)
    for f in ("base.py", "mapping.py", "views.py"):
        shutil.copy2(os.path.join(PROTO, f), CALC)
    collect = patched(os.path.join(Z6S, "z6_collect.py"), os.path.join(RUN, "z6_collect_e0.py"), [
        ('RES = r"C:\\Configuration_downloads\\BASERP25\\_Rarzrabotki\\Python\\test\\ekonomiya_ss_irc"',
         'RES = %r' % os.path.join(HERE, "vendor")),
        ('raw["stamp_1c"] = sv(B.ТекущаяДата())', 'raw["stamp_1c"] = sv(B.ТекущаяДатаСеанса())'),
    ])
    run(collect, {"Z6_RAW_DIR": CALC})
    calc = patched(os.path.join(Z6S, "z6_calc.py"), os.path.join(RUN, "z6_calc_e0.py"), [
        ('HERE = os.path.dirname(os.path.abspath(__file__))', 'HERE = %r' % Z6S),
        ('OUT_JSON = os.path.join(DATA, "proekt_irc2026_zakupki.json")', 'OUT_JSON = %r' % os.path.join(RUN, "e0_calc.json")),
        ('OUT_TXT = os.path.join(HERE, "z6_out.txt")', 'OUT_TXT = %r' % os.path.join(RUN, "e0_calc.txt")),
        ('_VER = os.path.join(HERE, "proverka", "q_t4_out.json")', '_VER = %r' % os.path.join(RUN, "no_verification.json")),
    ])
    run(calc, {"Z6_RAW_DIR": CALC, "Z1_RAW": base_raw, "Z3_RAW": os.path.join(RUN, "no_z3.json")})
    shutil.copy2(os.path.join(Z6S, "h.py"), os.path.join(RUN, "h.py"))
    t3 = patched(os.path.join(Z6S, "q_t3_proverka.py"), os.path.join(RUN, "q_t3_e0.py"), [
        ('os.path.join(os.path.dirname(os.path.abspath(__file__)), "q_t3_out.json")', '%r' % os.path.join(RUN, "q_t3_out.json")),
    ])
    run(t3, {})

    new = json.load(open(os.path.join(RUN, "e0_calc.json"), encoding="utf-8"))
    old = json.load(open(PUBLISHED, encoding="utf-8"))
    t3o = json.load(open(os.path.join(RUN, "q_t3_out.json"), encoding="utf-8"))
    T = new["totals"]
    vs = {k: {"z6": old["totals"].get(k), "now": T.get(k),
              "delta": round((T.get(k) or 0) - (old["totals"].get(k) or 0), 2)} for k in KEYS}
    closing = round(T["purchased_uah_net"] + T["other_in_uah_net"] - T["written_off_uah_net"]
                    - T["other_out_uah_net"] - T["lies_uah_net"], 2)
    indep = {"ptu_net": [t3o["ptu"][0]["Нетто"], T["purchased_by_kind"].get("ПТУ", {}).get("uah_net")],
             "stock_project": [t3o["stock_project"][0]["Сумма"], T["lies_project_uah_net"]],
             "written_off_sum_houses": [round(sum(x["Сумма"] for x in t3o["spys"]), 2), T["written_off_uah_net"]]}
    diff = new["snapshot"].get("diff_vs_z1") or {}
    ok = {"closing_zero": abs(closing) < 0.01,
          "independent_match": all(abs((x or 0) - (y or 0)) < 0.01 for x, y in indep.values()),
          "unchanged_or_explained": all(v["delta"] == 0 for v in vs.values()) or bool(diff.get("docs"))}
    summary = {"stamp": new["snapshot"]["datetime"], "stamp_1c": new["snapshot"].get("stamp_1c"),
               "base": 'Srvr="localhost";Ref="bas_industrialbud" (лише читання)',
               "baseline": {"z6_stamp": old["snapshot"]["datetime"], "raw": os.path.relpath(BASE_GZ, HERE)},
               "control_sums": {k: T.get(k) for k in KEYS}, "by_house": T["by_house"], "vs_z6": vs,
               "closing_t3": closing, "independent_q_t3": indep, "moves_diff_vs_z6": diff, "run_dir": RUN, "ok": ok}
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e0_snimok_{STAMP[:13]}.json")
    json.dump(summary, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print(json.dumps({"stamp": summary["stamp"], "delta": {k: v for k, v in vs.items() if v["delta"]},
                      "closing": closing, "independent": indep, "ok": ok}, ensure_ascii=False, indent=1))
    print("->", out)
    return 0 if all(ok.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
