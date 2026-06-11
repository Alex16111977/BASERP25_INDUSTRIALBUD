# -*- coding: utf-8 -*-
"""Сравнение независимого пересчёта (JSON от агентов) с файлом сверки.

Использование: положить erp_verify.json и buh_verify.json рядом (в %TEMP%)
и запустить. Сверяет каждую сумму с листом 'Сверка' выходного файла.
"""
import json
import os
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import openpyxl

TMP = os.environ.get("TEMP", ".")
F_ERP = os.path.join(TMP, "erp_verify.json")
F_BUH = os.path.join(TMP, "buh_verify.json")
F_OUT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Реестры\Сверка актов Глобино.xlsx"

with open(F_ERP, encoding="utf-8") as f:
    erp = json.load(f)
with open(F_BUH, encoding="utf-8") as f:
    buh = json.load(f)

erp_map = {a["num"]: a["sum"] for a in erp["acts"]}
buh_map = {a["num"]: (a["sum205"], a["sum2051"]) for a in buh["acts"]}

wb = openpyxl.load_workbook(F_OUT, data_only=True)
ws = wb["Сверка"]
bad = 0
seen_erp, seen_buh = set(), set()
for row in ws.iter_rows(min_row=2, values_only=True):
    num = row[1]
    if not num or num == "ИТОГО":
        continue
    erp_sum, b205, b2051 = row[4], row[6], row[7]

    if num in erp_map:
        seen_erp.add(num)
        if erp_sum is None or abs(erp_map[num] - erp_sum) > 0.005:
            print(f"MISMATCH ЕРП {num}: verify={erp_map[num]} file={erp_sum}")
            bad += 1
    elif erp_sum is not None:
        print(f"MISMATCH ЕРП {num}: в файле есть ({erp_sum}), у проверяющего нет")
        bad += 1

    if num in buh_map:
        seen_buh.add(num)
        v205, v2051 = buh_map[num]
        if abs(v205 - (b205 or 0)) > 0.005 or abs(v2051 - (b2051 or 0)) > 0.005:
            print(f"MISMATCH Бух {num}: verify=({v205}, {v2051}) file=({b205}, {b2051})")
            bad += 1
    elif b205 is not None or b2051 is not None:
        print(f"MISMATCH Бух {num}: в файле есть ({b205}/{b2051}), у проверяющего нет")
        bad += 1

for num in set(erp_map) - seen_erp:
    print(f"MISMATCH: у проверяющего ЕРП есть {num}, в файле сверки нет")
    bad += 1
for num in set(buh_map) - seen_buh:
    print(f"MISMATCH: у проверяющего Бух есть {num}, в файле сверки нет")
    bad += 1

print("-" * 60)
print(f"ЕРП: документов у проверяющего {erp['count']}, итог {erp['total']:,.2f}")
print(f"Бух: номеров у проверяющего {buh['count']}, итоги 205={buh['total205']:,.2f} 2051={buh['total2051']:,.2f}")
print("VERIFY PASS" if bad == 0 else f"VERIFY FAIL: расхождений {bad}")
sys.exit(0 if bad == 0 else 1)
