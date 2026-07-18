# -*- coding: utf-8 -*-
"""
А_ПереносДвиженийИзКазны — періодичний запуск через V83.COMConnector.

Створює (якщо немає) або перепроводить документ "А_ПереносДвиженийИзКазны"
за обраний місяць у базі ERP, викликає його експортну процедуру
ЗаполнитьИзКазны() (підключається до BuhKazn, читає БДДС за період,
заповнює ТЧ СтрокиДвижений), а потім записує з проведенням.

Запуск:
  python perenos_iz_kazny.py

Налаштування періоду — у блоці "НАЛАШТУВАННЯ ПЕРІОДУ" нижче:
  MONTH_FROM = None        → поточний календарний місяць
  MONTH_FROM = "2026-04"   → один квітень 2026
  MONTH_FROM = "2026-01"   + MONTH_TO = "2026-03"   → діапазон Q1 2026

Передумови:
  - pywin32 (win32com.client + pythoncom)
  - 1С Server "SQLSERVER", база ERP "BaseERP" (доступ Администратор/24043)
  - У документі А_ПереносДвиженийИзКазны існує експортна процедура
    ЗаполнитьИзКазны() (Documents/А_ПереносДвиженийИзКазны/Ext/ObjectModule.bsl).
  - Реквізит "Месяц" обов'язковий; "Организация" авто-fallback на ТОВ ІНДАСТРІАЛБУД
    (КодПоЕДРПОУ = 40645273).

Вихідний код:
  exit 0 — усі місяці успішно проведені
  exit 1 — хоча б один місяць провалився (FILL FAIL / POST FAIL / FIND/CREATE FAIL)
"""

import datetime
import io
import sys
import time

import pythoncom
import win32com.client


# === НАЛАШТУВАННЯ ПЕРІОДУ (редагуй ці значення) ======================
MONTH_FROM = "2024-01"   # "YYYY-MM" або None (None → поточний місяць)
MONTH_TO   = "2026-05"   # "YYYY-MM" або None (None → один місяць, що дорівнює MONTH_FROM)
# =====================================================================

# === ПІДКЛЮЧЕННЯ ДО 1С ===============================================
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
ORG_EDRPOU = "40645273"  # ТОВ ІНДАСТРІАЛБУД
# =====================================================================


def parse_month(s):
    return datetime.datetime.strptime(s, "%Y-%m")


def last_day_of_month(d):
    next_m = d.replace(day=28) + datetime.timedelta(days=4)
    return next_m - datetime.timedelta(days=next_m.day)


def month_iter(start, end):
    current = start.replace(day=1)
    end = end.replace(day=1)
    while current <= end:
        yield current
        if current.month == 12:
            current = current.replace(year=current.year + 1, month=1)
        else:
            current = current.replace(month=current.month + 1)


def resolve_periods():
    if MONTH_FROM:
        start = parse_month(MONTH_FROM)
        end = parse_month(MONTH_TO) if MONTH_TO else start
        return list(month_iter(start, end))
    today = datetime.datetime.now()
    return [today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)]


def find_or_create_doc(erp, month, org):
    """
    Шукає не-помічений на видалення документ за діапазоном Месяц
    [початок місяця .. кінець місяця 23:59:59].
    Використовуємо МЕЖДУ замість '=', бо COM-передача datetime через
    pywintypes.Time зазнає timezone-shift і точне порівняння '=' може не
    знаходити існуючі документи навіть з тим самим календарним 1-м числом.
    Додатково фільтруємо по організації.
    """
    last = last_day_of_month(month)
    month_start = datetime.datetime(month.year, month.month, 1, 0, 0, 0)
    month_end = datetime.datetime(last.year, last.month, last.day, 23, 59, 59)

    q = erp.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 "
        "  Док.Ссылка КАК Ссылка, "
        "  Док.Проведен КАК Проведен, "
        "  Док.Номер КАК Номер "
        "ИЗ Документ.А_ПереносДвиженийИзКазны КАК Док "
        "ГДЕ Док.Месяц МЕЖДУ &НачалоМесяца И &КонецМесяца "
        "    И Док.Организация = &Организация "
        "    И НЕ Док.ПометкаУдаления "
        "УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
    )
    q.УстановитьПараметр("НачалоМесяца", month_start)
    q.УстановитьПараметр("КонецМесяца", month_end)
    q.УстановитьПараметр("Организация", org)
    res = q.Выполнить()

    if not res.Пустой():
        sel = res.Выбрать()
        sel.Следующий()
        doc = sel.Ссылка.ПолучитьОбъект()
        if sel.Проведен:
            doc.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        return doc, "found", str(sel.Ссылка)

    doc = erp.Документы.А_ПереносДвиженийИзКазны.СоздатьДокумент()
    doc.Дата = month_end
    doc.Месяц = month_start
    doc.Организация = org
    doc.Записать()
    return doc, "created", str(doc.Ссылка)


def process_month(erp, org, month):
    label = month.strftime("%Y-%m")
    print(f"--- {label} ---")
    result = {"month": label, "status": "?", "rows": 0, "secs": 0.0, "ref": ""}

    try:
        doc, kind, ref = find_or_create_doc(erp, month, org)
        result["ref"] = ref
        print(f"  {'Знайдений' if kind == 'found' else 'Створений'}: {ref}")
    except Exception as e:
        print(f"  [ERR find/create] {e}")
        result["status"] = "FIND/CREATE FAIL"
        return result

    t0 = time.time()
    try:
        doc.ЗаполнитьИзКазны()
        t_fill = time.time() - t0
    except Exception as e:
        print(f"  [ERR fill] {e}")
        result["status"] = "FILL FAIL"
        return result

    rows = doc.СтрокиДвижений.Количество()
    print(f"  Заповнено: {rows} рядків за {t_fill:.1f}с")
    result["rows"] = rows

    t0 = time.time()
    try:
        doc.Записать(erp.РежимЗаписиДокумента.Проведение)
        t_post = time.time() - t0
    except Exception as e:
        print(f"  [ERR post] {e}")
        result["status"] = "POST FAIL"
        result["secs"] = t_fill
        return result

    print(f"  Проведено за {t_post:.1f}с")
    result["status"] = "OK"
    result["secs"] = t_fill + t_post
    return result


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    months = resolve_periods()

    print("=" * 70)
    print(f"А_ПереносДвиженийИзКазны — обробка {len(months)} місяця(ів)")
    print(f"Період: {months[0].strftime('%Y-%m')} → {months[-1].strftime('%Y-%m')}")
    print("=" * 70)

    pythoncom.CoInitialize()
    try:
        v8 = win32com.client.Dispatch("V83.COMConnector")
        erp = v8.Connect(CONN_ERP)
        print("[OK] Підключення до BaseERP\n")
    except Exception as e:
        print(f"[ERR] З'єднання з ERP не вдалося: {e}", file=sys.stderr)
        sys.exit(1)

    org = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", ORG_EDRPOU)
    if not erp.ЗначениеЗаполнено(org):
        print(f"[ERR] Не знайдено організацію з КодПоЕДРПОУ={ORG_EDRPOU}", file=sys.stderr)
        sys.exit(1)

    results = [process_month(erp, org, m) for m in months]
    print()

    print("=" * 70)
    print(f"{'Місяць':10s} {'Статус':18s} {'Рядків':>8s} {'Час':>8s}")
    print("-" * 70)
    total_rows = 0
    total_secs = 0.0
    failures = 0
    for r in results:
        print(f"{r['month']:10s} {r['status']:18s} {r['rows']:>8d} {r['secs']:>7.1f}с")
        total_rows += r["rows"]
        total_secs += r["secs"]
        if r["status"] != "OK":
            failures += 1
    print("-" * 70)
    print(f"{'РАЗОМ':10s} {'':18s} {total_rows:>8d} {total_secs:>7.1f}с")
    print("=" * 70)

    sys.exit(0 if failures == 0 else 1)


if __name__ == "__main__":
    main()
