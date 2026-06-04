# -*- coding: utf-8 -*-
"""
Phase 7 — Проверка Подразделения в справочниках БанковскиеСчета и Кассы
для виновных счетов/касс.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp

erp = connect_erp()
S = erp.String

print("=" * 110)
print("Phase 7 — Подразделение в справочниках БС / Касс (виновные)")
print("=" * 110)

# === Банковские счета (БС у которых были пустые проводки в РНДС) ===
print("\n[A] Банковские счета:")
БС_номера = [
    "UA973005280000026004000010559",   # ОТП_ТОВ
    "UA633395002600701537072000003",   # ТАС_Спецтехн
]
for n in БС_номера:
    q = erp.NewObject("Запрос")
    q.Текст = f"""
    ВЫБРАТЬ
        Б.Ссылка КАК Ссылка,
        ПРЕДСТАВЛЕНИЕ(Б.Ссылка) КАК Имя,
        Б.Подразделение КАК Подр,
        Б.НомерСчета КАК Ном,
        Б.Владелец КАК Влад
    ИЗ Справочник.БанковскиеСчетаОрганизаций КАК Б
    ГДЕ Б.НомерСчета = "{n}"
    """
    try:
        r = q.Выполнить().Выгрузить()
        print(f"\n  БС: {n}")
        for i in range(r.Количество()):
            rec = r.Получить(i)
            podr = S(rec.Подр) if erp.ЗначениеЗаполнено(rec.Подр) else "(пусто)"
            vlad = S(rec.Влад) if erp.ЗначениеЗаполнено(rec.Влад) else ""
            print(f"    Имя: {rec.Имя}")
            print(f"    Подразделение: {podr}")
            print(f"    Владелец: {vlad}")
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        print(f"    FAIL: {info[2] if info else e}")

# === Кассы ===
print("\n\n[B] Кассы:")
кассы_имена = ["2 Касса Подгорцы", "Касса Подгорцы", "2 Касса Строительство"]
for kn in кассы_имена:
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Имя", kn)
    q.Текст = """
    ВЫБРАТЬ
        К.Ссылка КАК Ссылка,
        К.Наименование КАК Имя,
        К.Подразделение КАК Подр,
        К.ВалютаДенежныхСредств КАК Валюта
    ИЗ Справочник.Кассы КАК К
    ГДЕ К.Наименование = &Имя
    """
    try:
        r = q.Выполнить().Выгрузить()
        print(f"\n  Касса: {kn}")
        for i in range(r.Количество()):
            rec = r.Получить(i)
            podr = S(rec.Подр) if erp.ЗначениеЗаполнено(rec.Подр) else "(пусто)"
            val = S(rec.Валюта) if erp.ЗначениеЗаполнено(rec.Валюта) else ""
            print(f"    Имя: {rec.Имя}")
            print(f"    Подразделение: {podr}")
            print(f"    Валюта: {val}")
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        print(f"    FAIL: {info[2] if info else e}")
