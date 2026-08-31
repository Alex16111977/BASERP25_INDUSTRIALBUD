# -*- coding: utf-8 -*-
"""Сверка: то ли написано в документации, что реально лежит в коде и на диске."""
import io
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

БАЗА = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki"
МОДУЛЬ = os.path.join(БАЗА, r"Отчеты\А_ОтчетPL_ПолныйЕРП\Ext\ObjectModule.bsl")
ТЕСТ = os.path.join(БАЗА, "Python", "test")
ДОКИ = os.path.join(БАЗА, "notebook", "knowledge_А_ОтчетPL_ПолныйЕРП")

провалов = [0]


def чек(имя, ок, подр=""):
    if not ок:
        провалов[0] += 1
    print("  [%s] %s%s" % ("OK  " if ок else "FAIL", имя, ("  — " + подр) if подр else ""))


м_сырой = io.open(МОДУЛЬ, encoding='utf-8-sig').read()
# Комментарии выкидываем: имена снятых таблиц законно упоминаются в пояснении, ПОЧЕМУ
# их больше нет. Проверять надо ИСПОЛНЯЕМЫЙ текст, а не пояснения к нему.
м = "\n".join(с for с in м_сырой.split("\n")
              if not с.strip().lstrip("|").lstrip().startswith("//"))

print("=" * 88)
print("1. Код: правило источника без исключений")
print("=" * 88)
чек("условие '= &ПоследнийМесяц' встречается 4 раза (по числу ЕРП-веток)",
    м.count("= &ПоследнийМесяц") == 4, "найдено %d" % м.count("= &ПоследнийМесяц"))
for имя in ("втМесяцыДок", "втДДСАморт", "втМесяцыДокАморт",
            'УстановитьПараметр("СтатьяАмортизации"', "ПолучитьСтатьюPLПоКоду",
            "МД.ПодразделениеОтчета ЕСТЬ NULL"):
    чек("в исполняемом коде НЕТ «%s»" % имя, имя not in м)

print("")
print("=" * 88)
print("2. Файлы тестов и запросов — как описано в 08_testy_etalony.md")
print("=" * 88)
должны_быть = ["v7_full.sql", "v8_full.sql",
               "test_otchetpl_v7_strogoe_pravilo_2026-08-20.py",
               "test_otchetpl_v8_ytd_2026-08-24.py",
               "acceptance_otchetpl_novaya_forma_2026-08-20.py",
               "acceptance_otchetpl_excel_parity_v3.py",
               "acceptance_pl_s_nachala_goda_2026-08-24.py",
               "acceptance_rusifikaciya_2026-08-26.py",
               "v9_full.sql",
               "test_otchetpl_v9_rko_analitika_2026-08-27.py",
               "acceptance_rko_analitika_2026-08-27.py",
               "v10_full.sql",
               "test_otchetpl_v10_kommentarii_2026-08-27.py",
               "v11_full.sql",
               "test_otchetpl_v11_amortizaciya_os_2026-08-28.py",
               "acceptance_amortizaciya_os_2026-08-28.py"]
для_истории = ["v2_full.sql", "v3_full.sql", "v4_full.sql", "v5_full.sql", "v6_full.sql"]
удалены = ["test_otchetpl_v5_amortizaciya_2026-08-20.py",
           "test_otchetpl_v6_pko_forma2_2026-08-20.py"]
for ф in должны_быть + для_истории:
    чек("есть %s" % ф, os.path.exists(os.path.join(ТЕСТ, ф)))
for ф in удалены:
    чек("удалён %s" % ф, not os.path.exists(os.path.join(ТЕСТ, ф)))

print("")
print("=" * 88)
print("3. Запрос в модуле == действующий v11_full.sql")
print("=" * 88)
sql = io.open(os.path.join(ТЕСТ, "v11_full.sql"), encoding='utf-8').read().replace("\r\n", "\n").rstrip("\n")
куски = [s for s in sql.split("\n") if s.strip()]
попало = sum(1 for s in куски if s.replace('"', '""') in м_сырой)
чек("все непустые строки v11 присутствуют в модуле", попало == len(куски),
    "совпало %d из %d" % (попало, len(куски)))

print("")
print("=" * 88)
print("4. Документация: правило прописано жёстко и без старых формулировок")
print("=" * 88)
карта = io.open(os.path.join(ДОКИ, "KNOWLEDGE_MAP.md"), encoding='utf-8').read()
чек("KNOWLEDGE_MAP начинается с блока «ИСКЛЮЧЕНИЙ НЕТ»",
    "ПРАВИЛО ИСТОЧНИКА ФАКТА — ЗАКОН. ИСКЛЮЧЕНИЙ НЕТ." in карта[:2000])
чек("в карте есть дословная формулировка заказчика",
    "регистры только за последний месяц периода, всё раньше — только документы А_ОтчетPL" in карта)
чек("инвариант №1 переписан под строгое правило",
    "РЕГИСТРЫ — ТОЛЬКО ЗА ПОСЛЕДНИЙ МЕСЯЦ ПЕРИОДА" in карта)
чек("действующим документом объявлен файл 12",
    "**ВЕРСИЯ 5 — ДЕЙСТВУЮЩАЯ**" in карта)

МАРКЕРЫ = ("СНЯТ", "снят", "ОТМЕН", "отмен", "НОЛЬ", "ноль", "нуль",
           "ЗАПРЕЩЕНО", "запрещ", "ИСТОРИЯ", "история", "исключений нет",
           "БОЛЬШЕ НЕ", "больше нет")
плохо = []
for имя in sorted(os.listdir(ДОКИ)):
    if not имя.endswith(".md"):
        continue
    строки = io.open(os.path.join(ДОКИ, имя), encoding='utf-8').read().split("\n")
    for стр_н, стр in enumerate(строки, 1):
        if "месяц без документа" not in стр and "документа по ветке нет" not in стр:
            continue
        # старая формулировка допустима, только если рядом сказано, что она отменена
        окно = "\n".join(строки[max(0, стр_н - 5):стр_н + 4])
        if not any(w in окно for w in МАРКЕРЫ):
            плохо.append("%s:%d" % (имя, стр_н))
чек("старой формулировки без пометки об отмене нигде нет", not плохо,
    "; ".join(плохо[:6]))

устарело = []
for имя in sorted(os.listdir(ДОКИ)):
    if not имя.endswith(".md"):
        continue
    т = io.open(os.path.join(ДОКИ, имя), encoding='utf-8').read()
    строки2 = т.split("\n")
    маркеры2 = МАРКЕРЫ + ("Удален", "удален", "УДАЛЕН", "снимает", "снят",
                          "БОЛЬШЕ НЕТ", "нельзя", "прежнее")
    for фраза in ("исключение источника для амортизации", "исключение для амортизации"):
        for стр_н, стр in enumerate(строки2, 1):
            if фраза not in стр:
                continue
            окно2 = "\n".join(строки2[max(0, стр_н - 5):стр_н + 4])
            if not any(w in окно2 for w in маркеры2):
                устарело.append("%s:%d" % (имя, стр_н))
чек("нет утверждений о живом исключении для амортизации", not устарело,
    "; ".join(устарело[:6]))

print("")
print("=" * 88)
print("ИТОГ: %s (провалов %d)" % ("ЗЕЛЁНО" if провалов[0] == 0 else "ЕСТЬ ПРОВАЛЫ", провалов[0]))
print("=" * 88)
sys.exit(1 if провалов[0] else 0)
