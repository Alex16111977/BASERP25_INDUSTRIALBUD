# -*- coding: utf-8 -*-
"""
СКРИПТ 29 (Phase 4 шаг 1) — Перепроведение документов с расхождением

ВАЖНО (memory feedback_no_doc_delete_in_tests):
    НЕ удалять документы! Get → ПолучитьОбъект() → Записать(Проведение).

Берём документы из 20_full_discovery.csv (тип СписаниеБезналичных + РКО + ОПК с
ХозОп=ВозвратОплатыКлиенту — целевые для нашей правки). Также перепроводим
все документы из 26_double_count_candidates (это те где есть РСК+РСП ПереносАванса).
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, ARTIFACTS_DIR

erp = connect_erp()

import win32com.client
# Режим записи документа — через COM-перечисление 1С
РежимПроведения = erp.РежимЗаписиДокумента.Проведение

# Загружаем 26_double_count_candidates — это документы которых правка коснётся
docs_set = set()
with open(os.path.join(ARTIFACTS_DIR, "26_double_count_candidates.csv"),
          encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f, delimiter=";"):
        if r["ТипДок"]:
            docs_set.add((r["ТипДок"], r["Документ"]))

print(f"Документов к перепроведению (из 26): {len(docs_set)}")
print("Перепроводим... (может занять несколько минут)")

import re

def parse_name(имя):
    """Извлекаем (Номер, ДатаСтр) из строки представления.
    'Приходный кассовый ордер 000Ц-000023 от 05.12.2025 13:01:04' →
        ('000Ц-000023', '05.12.2025 13:01:04')
    """
    m = re.search(r"\s([\w\-]+)\s+от\s+([\d\.]+(?:\s+[\d:]+)?)", имя)
    if m:
        return m.group(1), m.group(2).strip()
    return None, None


ok = err = 0
errs = []
for i, (тип, имя) in enumerate(docs_set, 1):
    if i % 20 == 0:
        print(f"  ... {i}/{len(docs_set)} (OK={ok}, ERR={err})")
    try:
        ном, дата_стр = parse_name(имя)
        if not ном:
            errs.append((тип, имя, "не парсится Номер")); err += 1; continue

        q = erp.NewObject("Запрос")
        q.УстановитьПараметр("Ном", ном)
        q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 5 Ссылка, Дата ИЗ Документ.{тип} ГДЕ Номер = &Ном'
        res = q.Выполнить().Выгрузить()
        if res.Количество() == 0:
            errs.append((тип, имя, f"не найден Номер={ном}")); err += 1; continue

        # Если несколько — выберем по совпадению даты
        ref = res.Получить(0).Ссылка
        if res.Количество() > 1 and дата_стр:
            for j in range(res.Количество()):
                rec = res.Получить(j)
                if дата_стр in str(rec.Дата):
                    ref = rec.Ссылка; break

        obj = ref.ПолучитьОбъект()
        if obj is None:
            errs.append((тип, имя, "ПолучитьОбъект=None")); err += 1; continue
        obj.Записать(РежимПроведения)
        ok += 1
    except Exception as e:
        msg = str(e)[:150]
        errs.append((тип, имя, msg))
        err += 1

print(f"\nИТОГО: OK={ok}, ERR={err}")
if errs[:10]:
    print("\nПервые 10 ошибок:")
    for тип, имя, msg in errs[:10]:
        print(f"  {тип}/{имя[:50]}: {msg[:100]}")

sys.exit(0 if err == 0 else 1)
