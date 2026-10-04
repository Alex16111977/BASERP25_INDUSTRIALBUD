# -*- coding: utf-8 -*-
"""Приёмка этапа 2 (§10.1) одним прогоном. Каждый тест — отдельным процессом (своё COM-соединение, свои транзакции с откатом):
30 м — 170 позиций, 17 зелёных, жёлтые только на ОН с одной позицией (test_e2a_zapolnit); мембрана, плінтус, труба, кутник,
септик 1/2/1/1 и Σ 12 000,00, Σ норм 1 009 864,74 ± 1,00 (test_e2b_ceny); узлы и узел в узле (test_e2b_uzly); 18 из 22 нулевых
с фактом, 4 алиаса (test_e2b_nulevye); 1 сирота — зелёное (test_e2c_perepriv); 15 м копией — жёлтые расхождения с фактом,
k = 1 у поз. 27/44/45, Σ 704 612,01, И4 0,00, «Не постачається» не мешает «Діє» (test_e2c_kopiya); формы (test_e2d_formy);
регресс этапа 1 (test_e1_priemka). Затем: регистрация СС к обмену в конце = в начале прогона, элементов справочника не
прибавилось."""
import datetime
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, OUT, connect, q  # noqa: E402
from test_e1_priemka import START, exchange_snapshot, k10_exchange_untouched  # noqa: E402

TESTS = ["test_e2_zaprosy.py", "test_e2a_edinicy.py", "test_e2a_zapolnit.py", "test_e2b_uzly.py", "test_e2b_nulevye.py",
         "test_e2b_ceny.py", "test_e2c_perepriv.py", "test_e2c_kopiya.py", "test_e2d_formy.py", "test_e1_priemka.py"]


def spec_refs(B):
    rows = q(B, """ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец""")
    return sorted(B.String(r["Ссылка"].УникальныйИдентификатор()) for r in rows)


def run(name):
    p = subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, name)], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=3600)
    tail = (p.stdout + p.stderr).strip().splitlines()[-4:]
    return p.returncode == 0, tail


def main():
    B = connect()
    before = spec_refs(B)
    START["exchange"] = exchange_snapshot(B)
    ok, fails = {}, {}
    for name in TESTS:
        passed, tail = run(name)
        (ok if passed else fails)[name] = tail
        print("OK  " if passed else "FAIL", name, "|", " / ".join(tail))
    try:
        k10_exchange_untouched(B)
        ok["exchange_untouched"] = "регистрация СС к обмену в конце прогона = в начале"
    except AssertionError as e:
        fails["exchange_untouched"] = repr(e)[:500]
    after = spec_refs(B)
    if after != before:
        fails["no_data_written"] = f"элементов справочника было {len(before)}, стало {len(after)} — тест записал данные мимо транзакции"
    else:
        ok["no_data_written"] = f"элементов справочника {len(after)}, как до прогона"
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e2_priemka_{datetime.datetime.now():%Y%m%d_%H%M}.json")
    json.dump({"ok": ok, "fail": fails}, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("->", out, "ПРОВАЛОВ:", len(fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
