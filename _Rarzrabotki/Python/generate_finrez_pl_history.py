# -*- coding: utf-8 -*-
"""
Генерує і проводить документи `Документ.А_ФинРез_PL` за період 01.01.2024 – 30.11.2025.

Один документ на (Організація, Місяць). За шаблон береться 00000000001 від 31.12.2025 12:00:00:
копіюється тільки Організація, дати/місяць — нові, номер — авто.

Правила:
 - Якщо вже є ПРОВЕДЕНИЙ документ за цей місяць — SKIP.
 - Якщо є непроведений (Проведен=Ложь, ПометкаУдаления=Ложь) — провести його, не плодити дубль.
 - Інакше створити новий документ (Дата = останній день місяця 12:00:00, Месяц = 1-й день).

Запуск:
    python _Rarzrabotki/Python/generate_finrez_pl_history.py
"""

import sys
import calendar
import datetime
import traceback

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

TEMPLATE_UUID = "8452c387-4992-11f1-a2ea-bf0a2242d914"

START_YEAR, START_MONTH = 2024, 1
END_YEAR, END_MONTH = 2025, 11


def iter_periods(start_year, start_month, end_year, end_month):
    y, m = start_year, start_month
    while (y, m) <= (end_year, end_month):
        yield y, m
        m += 1
        if m == 13:
            m = 1
            y += 1


def find_template(erp):
    uid = erp.NewObject("УникальныйИдентификатор", TEMPLATE_UUID)
    ref = erp.Документы.А_ФинРез_PL.ПолучитьСсылку(uid)
    obj = ref.ПолучитьОбъект()
    if obj is None:
        raise RuntimeError(f"Шаблон UUID={TEMPLATE_UUID} не знайдено у БД")
    return ref


def find_doc_for_month(erp, organization, month_start, month_end):
    """Повертає (посилання, проведено) або (None, None) якщо документа немає."""
    q = erp.NewObject("Запрос")
    q.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка, Проведен КАК Проведен "
        "ИЗ Документ.А_ФинРез_PL "
        "ГДЕ НЕ ПометкаУдаления "
        "И Организация = &Организация "
        "И Месяц МЕЖДУ &НачМесяца И &КонМесяца"
    )
    q.SetParameter("Организация", organization)
    q.SetParameter("НачМесяца", month_start)
    q.SetParameter("КонМесяца", month_end)
    res = q.Execute()
    if res.Пустой():
        return None, None
    sel = res.Выбрать()
    sel.Следующий()
    return sel.Ссылка, sel.Проведен


def process_month(erp, organization, year, month, posting_mode):
    last_day = calendar.monthrange(year, month)[1]
    doc_date = datetime.datetime(year, month, last_day, 12, 0, 0)
    month_start = datetime.datetime(year, month, 1, 0, 0, 0)
    month_end = datetime.datetime(year, month, last_day, 23, 59, 59)
    label = f"{year}-{month:02d}"

    existing_ref, posted = find_doc_for_month(erp, organization, month_start, month_end)

    if existing_ref is not None:
        if posted:
            return ("SKIP", label, f"вже проведено: {erp.String(existing_ref)}")
        obj = existing_ref.ПолучитьОбъект()
        try:
            obj.Записать(posting_mode)
            return ("REPOST", label, f"проведено існуючий {erp.String(existing_ref)}")
        except Exception as exc:
            return ("FAIL", label, f"провести існуючий: {exc}")

    obj = erp.Документы.А_ФинРез_PL.СоздатьДокумент()
    obj.Дата = doc_date
    obj.Организация = organization
    obj.Месяц = month_start

    try:
        obj.Записать(posting_mode)
        return ("OK", label, f"створено №{obj.Номер} від {doc_date.strftime('%d.%m.%Y %H:%M:%S')}")
    except Exception as exc:
        return ("FAIL", label, f"створити+провести: {exc}")


def main():
    print("Підключаюсь до BaseERP...")
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN_ERP)
    print("OK")

    template_ref = find_template(erp)
    template_obj = template_ref.ПолучитьОбъект()
    organization = template_obj.Организация
    print(f"Шаблон: {erp.String(template_ref)}")
    print(f"Організація: {erp.String(organization)}")

    posting_mode = erp.РежимЗаписиДокумента.Проведение

    periods = list(iter_periods(START_YEAR, START_MONTH, END_YEAR, END_MONTH))
    print(f"Періодів: {len(periods)} (від {START_YEAR}-{START_MONTH:02d} до {END_YEAR}-{END_MONTH:02d})")
    print("-" * 80)

    counters = {"OK": 0, "REPOST": 0, "SKIP": 0, "FAIL": 0}
    fails = []

    for year, month in periods:
        try:
            status, label, msg = process_month(erp, organization, year, month, posting_mode)
        except Exception as exc:  # noqa: BLE001
            status, label, msg = ("FAIL", f"{year}-{month:02d}", f"виняток: {exc}")
            traceback.print_exc()

        counters[status] += 1
        print(f"  [{status:6s}] {label}: {msg}")
        if status == "FAIL":
            fails.append((label, msg))

    print("-" * 80)
    print(f"OK (створено):      {counters['OK']}")
    print(f"REPOST (провели):   {counters['REPOST']}")
    print(f"SKIP (вже було):    {counters['SKIP']}")
    print(f"FAIL:               {counters['FAIL']}")

    if fails:
        print("\nПомилки:")
        for label, msg in fails:
            print(f"  {label}: {msg}")
        sys.exit(1)


if __name__ == "__main__":
    main()
