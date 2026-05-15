# -*- coding: utf-8 -*-
"""Дамп resolved SQL backend table/_Fld для Balance Stage об'єктів.

Джерело: mapping/baserp_storage.json (структура {"objects": {meta: {primary_table, fields}}}).
Результат -> джерело істини для raw_sql у pipelines/fact_balance.json + dim_pap_articles.json.

Також контролює, що refresh_mapping НЕ зрегресував cashflow/PnL (additive diff).
"""
import io
import json
import pathlib
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

MP = json.loads(
    pathlib.Path(__file__).resolve().parents[1]
    .joinpath("mapping", "baserp_storage.json")
    .read_text(encoding="utf-8")
)
OBJ = MP["objects"]
print(f"baserp_storage.json: {MP['object_count']} objects, gen {MP['generated_at']}")

# --- Additive safety: live cashflow/PnL infra must remain present -------------
REGRESS_GUARD = [
    "РегистрСведений.А_ОтчетDDS_Свод",
    "РегистрСведений.А_ОтчетPL_Свод",
    "Перечисление.А_ИсточникDDS",
    "Перечисление.А_ИсточникPL",
    "Справочник.Организации",
    "Справочник.Номенклатура",
]
missing_guard = [m for m in REGRESS_GUARD if m not in OBJ]
print(f"\n[additive guard] cashflow/PnL present: "
      f"{'OK' if not missing_guard else 'REGRESSED ' + str(missing_guard)}")
assert not missing_guard, f"FAIL: refresh_mapping зрегресував {missing_guard}"

# --- Balance Stage objects ---------------------------------------------------
BALANCE_OBJ = [
    "РегистрСведений.А_ОтчетБаланс_Свод",
    "Документ.А_ФинРез_Баланс",
    "ПланВидовХарактеристик.СтатьиАктивовПассивов",
    "Перечисление.А_ИсточникБаланса",
    "Перечисление.ВидыСтатейУправленческогоБаланса",
]
for meta in BALANCE_OBJ:
    o = OBJ.get(meta)
    print(f"\n=== {meta} ===")
    if o is None:
        print("  <ВІДСУТНІЙ — refresh_mapping не побачив>")
        continue
    print(f"  primary_table: {o.get('primary_table')}")
    for logical, col in sorted(o.get("fields", {}).items()):
        print(f"    {logical:<28} -> {col}")

# --- Контроль: усі 17 dim + 4 res + Регистратор резолвлені -------------------
# А_ОтчетБаланс_Свод — підпорядкований реєстратору (має _RecorderRRef, без
# власного Періоду). Period береться з Документ.А_ФинРез_Баланс._Date_Time
# через INNER JOIN d._IDRRef = r._RecorderRRef — точно патерн fact_cashflow.json.
reg = OBJ.get("РегистрСведений.А_ОтчетБаланс_Свод", {})
F = reg.get("fields", {})
NEED = [
    "Организация", "Подразделение", "Статья", "Номенклатура", "Контрагент",
    "Партнер", "Склад", "ОбъектыЭксплуатации", "Договор", "ФизическоеЛицо",
    "ДенежныеСредства", "Source", "Аналитика1", "Аналитика2", "Аналитика3",
    "ОбъектРасчетов", "НематериальныйАктив",
    "СуммаНачальныйОстаток", "СуммаПриход", "СуммаРасход", "СуммаКонечныйОстаток",
    "Регистратор",
]
miss = [n for n in NEED if n not in F]
print(f"\n[А_ОтчетБаланс_Свод] не резолвлено: {miss if miss else 'НЕМАЄ (усі 22 OK)'}")
assert reg, "FAIL: РегистрСведений.А_ОтчетБаланс_Свод відсутній — повторити refresh_mapping"
assert not miss, f"FAIL: refresh_mapping не дав поля {miss}"

# Period-джерело — реєстратор-документ
doc = OBJ.get("Документ.А_ФинРез_Баланс", {}).get("fields", {})
assert doc.get("Дата") == "_Date_Time", f"FAIL: А_ФинРез_Баланс.Дата != _Date_Time ({doc.get('Дата')})"
assert doc.get("Ссылка") == "_IDRRef", f"FAIL: А_ФинРез_Баланс.Ссылка != _IDRRef ({doc.get('Ссылка')})"
print(f"[Period джерело] {OBJ['Документ.А_ФинРез_Баланс']['primary_table']}"
      f".{doc['Дата']} (JOIN ON _IDRRef = r._RecorderRRef) OK")
print("\nPASS Task 2 mapping")
