# -*- coding: utf-8 -*-
"""
Smoke: переводы между своими счетами на живом примере 1432 (30.01.2026, 2 426 200,50).
v6 API: анализ напрямую по счёту через АнализироватьДокументыПоСчетам(массив, НачП, КонП).
Проверяет ОБЕ стороны пары. Только чтение.
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
    global FAILS
    obr = erp.ВнешниеОбработки.Создать(EPF, False)
    mas = erp.NewObject("Массив")
    mas.Добавить(schet_po_iban(iban))
    rez = obr.АнализироватьДокументыПоСчетам(
        mas, datetime(2026, 1, 25, 0, 0, 0), datetime(2026, 2, 5, 23, 59, 59))
    n_dok = obr.ТаблицаДокументов.Количество()
    print(f"[{iban[-6:]}] {erp.String(rez)} (рядків={n_dok})")

    nashel = False
    for i in range(n_dok):
        s = obr.ТаблицаДокументов.Получить(i)
        doc_erp = str(erp.String(s.ДокументЕРП))
        status = str(erp.String(s.Статус))
        if nomer_doc in doc_erp:
            nashel = True
            print(f"[{iban[-6:]}] ШУКАНИЙ: ЕРП='{doc_erp}' Бух='{erp.String(s.ДокументБух)}'"
                  f" СумЕРП={s.СуммаЕРП} СумБух={s.СуммаБух} Статус='{status}'")
            if status != "Переказ: синхронно" or str(erp.String(s.Действие)).strip() != "":
                print(f"[{iban[-6:]}] FAIL: очікував «Переказ: синхронно» без дії")
                FAILS += 1
    if not nashel:
        print(f"[{iban[-6:]}] FAIL: документ {nomer_doc} не знайдено")
        FAILS += 1


print("=== Сторона джерела (ОТП, Списание 1432) ===")
analiz_scheta(IBAN_ISTOCHNIK, "00000001432")
print("=== Сторона приймача (ТАС_Виробн, Поступление 00DL-008305) ===")
analiz_scheta(IBAN_PRIEMNIK, "00DL-008305")

print("РЕЗУЛЬТАТ: " + ("SMOKE OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
