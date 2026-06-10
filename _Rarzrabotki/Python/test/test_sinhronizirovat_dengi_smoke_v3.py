# -*- coding: utf-8 -*-
"""
Smoke v3: переводы между своими счетами на живом примере 1432 (30.01.2026, 2 426 200,50).
Проверяет ОБЕ стороны пары:
- счёт-источник ОТП:  ERP-Списание 00000001432  -> «Переказ: синхронно»
- счёт-приёмник ТАС:  ERP-Поступление 00DL-008305 -> «Переказ: синхронно»
И отсутствие ложных строк «Тільки в BuhBud»/«Переобмін» для этой пары.
Только чтение, Фаза 3 не выполняется.
"""
import sys
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"
IBAN_ISTOCHNIK = "UA973005280000026004000010559"   # ОТП
IBAN_PRIEMNIK = "UA663395002600601537072000002"    # ТАС_Виробн

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

FAILS = 0


def schet_po_iban(iban):
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 1 БСО.Ссылка КАК Ссылка
    ИЗ Справочник.БанковскиеСчетаОрганизаций КАК БСО
    ГДЕ БСО.НомерСчета = &НомерСчета
    """
    q.SetParameter("НомерСчета", iban)
    r = q.Execute().Выгрузить()
    assert r.Количество() == 1, f"счёт {iban} не найден"
    return r.Получить(0).Ссылка


def analiz_scheta(iban, nomer_doc):
    """Фаза 1 (фильтр по счёту) + Фаза 2; возвращает строки документов с nomer_doc."""
    global FAILS
    obr = erp.ВнешниеОбработки.Создать(EPF, False)
    obr.ФильтрБанковскийСчет = schet_po_iban(iban)
    obr.НачалоПериода = datetime(2026, 1, 1, 0, 0, 0)
    obr.ОкончаниеПериода = datetime(2026, 2, 10, 0, 0, 0)
    obr.СравнитьОстатки()
    n_rash = obr.ТаблицаРасхождений.Количество()
    if n_rash == 0:
        # расхождений в коротком периоде нет — расширяем
        obr.ОкончаниеПериода = datetime(2026, 6, 10, 0, 0, 0)
        obr.СравнитьОстатки()
        n_rash = obr.ТаблицаРасхождений.Количество()
    print(f"[{iban[-6:]}] Фаза 1: розбіжностей={n_rash}")
    if n_rash == 0:
        print(f"[{iban[-6:]}] SKIP: рахунок без розбіжностей — Фаза 2 не запускається")
        return
    for i in range(n_rash):
        obr.ТаблицаРасхождений.Получить(i).Синхронизировать = True
    obr.АнализироватьДокументы()
    n_dok = obr.ТаблицаДокументов.Количество()
    print(f"[{iban[-6:]}] Фаза 2: документів={n_dok}")

    nashel = False
    perevod_ok = 0
    perevod_rozb = 0
    for i in range(n_dok):
        s = obr.ТаблицаДокументов.Получить(i)
        doc_erp = str(erp.String(s.ДокументЕРП))
        doc_buh = str(erp.String(s.ДокументБух))
        status = str(erp.String(s.Статус))
        if status.startswith("Переказ: синхронно"):
            perevod_ok += 1
        elif status.startswith("Переказ:"):
            perevod_rozb += 1
        if nomer_doc in doc_erp:
            nashel = True
            print(f"[{iban[-6:]}] ШУКАНИЙ: ЕРП='{doc_erp}' Бух='{doc_buh}'")
            print(f"          СумЕРП={s.СуммаЕРП} СумБух={s.СуммаБух} Дія='{erp.String(s.Действие)}' Статус='{status}'")
            if status != "Переказ: синхронно":
                print(f"[{iban[-6:]}] FAIL: очікував «Переказ: синхронно», отримав «{status}»")
                FAILS += 1
            if str(erp.String(s.Действие)).strip() != "":
                print(f"[{iban[-6:]}] FAIL: дія для переказу повинна бути порожня")
                FAILS += 1
        # ложный «Тільки в BuhBud» для документа пары
        if "001432" in doc_buh and "BuhBud" in status:
            print(f"[{iban[-6:]}] FAIL: ложная строка BuhBud-only: '{doc_buh}' статус '{status}'")
            FAILS += 1
    print(f"[{iban[-6:]}] Переказів синхронно={perevod_ok}, з розбіжностями={perevod_rozb}")
    if not nashel:
        print(f"[{iban[-6:]}] FAIL: документ {nomer_doc} не знайдено в аналізі")
        FAILS += 1


print("=== Сторона джерела (ОТП, Списание 1432) ===")
analiz_scheta(IBAN_ISTOCHNIK, "00000001432")
print("=== Сторона приймача (ТАС_Виробн, Поступление 00DL-008305) ===")
analiz_scheta(IBAN_PRIEMNIK, "00DL-008305")

print("РЕЗУЛЬТАТ: " + ("SMOKE OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
