# -*- coding: utf-8 -*-
"""Этап 0 не должен зависеть от untracked-файлов основного каталога и от Temp (E2 §5.1, §9 п. 6):
vendor — копии текстов запросов f31/f33 (z6_collect читает из них TEXT), baseline — gzip снимка Z6 03.10.2026 23:41:12.
Запуск: python make_baseline.py           # создать
        python make_baseline.py --check   # 0 — всё на месте и снимок тот самый"""
import gzip
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, MAIN  # noqa: E402

SRC_RES = os.path.join(MAIN, r"_Rarzrabotki\Python\test\ekonomiya_ss_irc")
SRC_RAW = r"C:\Users\SUPPOR~1\AppData\Local\Temp\claude\z6_tmp\calc\raw.json"
VENDOR = os.path.join(HERE, "vendor")
BASE = os.path.join(HERE, "baseline", "z6_raw_20261003_234112.json.gz")
FILES = ["f31_buh_moves.py", "f33_buh_receipts.py"]


def check():
    for f in FILES:
        p = os.path.join(VENDOR, f)
        assert os.path.exists(p), f"нет {p}"
        assert 'TEXT = """' in open(p, encoding="utf-8").read(), f"в {f} нет TEXT"
    assert os.path.exists(BASE), f"нет {BASE}"
    raw = json.loads(gzip.open(BASE, "rt", encoding="utf-8").read())
    assert raw["stamp_full"] == "03.10.2026 23:41:12", raw["stamp_full"]
    assert len(raw["moves"]) == 4800, len(raw["moves"])
    print("OK baseline:", raw["stamp_full"], "движений", len(raw["moves"]))
    return 0


def make():
    os.makedirs(VENDOR, exist_ok=True)
    for f in FILES:
        shutil.copyfile(os.path.join(SRC_RES, f), os.path.join(VENDOR, f))
    os.makedirs(os.path.dirname(BASE), exist_ok=True)
    with open(SRC_RAW, "rb") as fi, gzip.open(BASE, "wb", compresslevel=9) as fo:
        shutil.copyfileobj(fi, fo)
    return check()


if __name__ == "__main__":
    sys.exit(check() if "--check" in sys.argv else make())
