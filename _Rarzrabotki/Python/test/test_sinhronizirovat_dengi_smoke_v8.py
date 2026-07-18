# -*- coding: utf-8 -*-
"""
Smoke v8: вариант B перевода (подчинённое Списание <- Поступление, Бух = «Надходження»).
Эталон: депозит UA693395002610201537072000001, март 2026.
- 000Ц-000187 -> «Переказ: синхронно», Бух=«Надходження ... 00DL-9066», СумаБух=-44 602 752,10
- остальные 000Ц-строки марта (000192,000197,000199,000201,000203,000205,000211,000212,000219) — синхронно
- нет ложных «Переказ: вхідна сторона в ЕРП не зв'язана» для UUID, покрытых парой
Только чтение, Фаза 3 не выполняется.
"""
import sys
from collections import Counter
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"
IBAN_DEPOZIT = "UA693395002610201537072000001"
ETALON_NOMERA = ["000187", "000192", "000197", "000199", "000201",
                 "000203", "000205", "000211", "000212", "000219"]

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 БСО.Ссылка КАК Ссылка
ИЗ Справочник.БанковскиеСчетаОрганизаций КАК БСО
ГДЕ БСО.НомерСчета = &НомерСчета
"""
q.SetParameter("НомерСчета", IBAN_DEPOZIT)
r = q.Execute().Выгрузить()
assert r.Количество() == 1, "депозитный счёт не найден в ERP"

obr = erp.ВнешниеОбработки.Создать(EPF, False)
obr.ФильтрБанковскийСчет = r.Получить(0).Ссылка
obr.НачалоПериода = datetime(2026, 3, 1, 0, 0, 0)
obr.ОкончаниеПериода = datetime(2026, 3, 31, 0, 0, 0)

obr.СравнитьОстатки()
n_rash = obr.ТаблицаРасхождений.Количество()
print(f"Фаза 1 (березень): розбіжностей={n_rash}")
if n_rash == 0:
    print("Розбіжностей за березень немає — розширюю період до 10.06")
    obr.ОкончаниеПериода = datetime(2026, 6, 10, 0, 0, 0)
    obr.СравнитьОстатки()
    n_rash = obr.ТаблицаРасхождений.Количество()
    print(f"Фаза 1 (розширено): розбіжностей={n_rash}")
assert n_rash >= 1, "немає розбіжностей — Фаза 2 не запускається"

for i in range(n_rash):
    obr.ТаблицаРасхождений.Получить(i).Синхронизировать = True
obr.АнализироватьДокументы()
n_dok = obr.ТаблицаДокументов.Количество()
print(f"Фаза 2: документів={n_dok}")

FAILS = 0
statusy = Counter()
etalon_status = {}
ne_znaydeno_c = 0
vhidna_for_pair = 0

for i in range(n_dok):
    s = obr.ТаблицаДокументов.Получить(i)
    status = str(erp.String(s.Статус))
    doc_erp = str(erp.String(s.ДокументЕРП))
    doc_buh = str(erp.String(s.ДокументБух))
    key = status.split("(")[0].strip()
    statusy[key] += 1

    for nom in ETALON_NOMERA:
        if f"000Ц-{nom}" in doc_erp:
            etalon_status[nom] = (status, doc_buh, float(s.СуммаЕРП), float(s.СуммаБух),
                                  str(erp.String(s.Действие)).strip())

    if "000Ц-" in doc_erp and status == "Переказ: документ не знайдено в Бух":
        ne_znaydeno_c += 1
        if ne_znaydeno_c <= 5:
            print(f"  ЛИШИЛОСЬ не знайдено: {doc_erp}")
    # ложные строки обратной проверки для Надходжень, покрытых парой
    if status.startswith("Переказ: вхідна сторона в ЕРП не зв'язана") and "Надходження" in doc_buh:
        vhidna_for_pair += 1
        if vhidna_for_pair <= 5:
            print(f"  ЛОЖНА вхідна сторона: Бух='{doc_buh}'")

print("\nРозподіл статусів:")
for k, v in sorted(statusy.items(), key=lambda x: -x[1]):
    print(f"  {v:5d}  {k}")

# --- Acceptance ---
print("\nЕталонні 000Ц-рядки:")
for nom in ETALON_NOMERA:
    if nom not in etalon_status:
        print(f"  FAIL: 000Ц-{nom} не знайдено в аналізі")
        FAILS += 1
        continue
    status, doc_buh, sum_erp, sum_buh, deystvie = etalon_status[nom]
    ok = status == "Переказ: синхронно" and deystvie == ""
    print(f"  000Ц-{nom}: '{status}' Бух='{doc_buh}' СумЕРП={sum_erp:.2f} СумБух={sum_buh:.2f}"
          + ("" if ok else "  <-- FAIL"))
    if status != "Переказ: синхронно":
        FAILS += 1
    if deystvie != "":
        print(f"    FAIL: дія для переказу повинна бути порожня, а ='{deystvie}'")
        FAILS += 1

if "000187" in etalon_status:
    status, doc_buh, sum_erp, sum_buh, _ = etalon_status["000187"]
    if "00DL-9066" not in doc_buh:
        print(f"  FAIL 000187: очікував Бух 00DL-9066, отримав '{doc_buh}'")
        FAILS += 1
    if abs(sum_buh - (-44602752.10)) > 0.005:
        print(f"  FAIL 000187: СумаБух={sum_buh}, очікував -44602752.10")
        FAILS += 1
    if abs(sum_erp - (-44602752.10)) > 0.005:
        print(f"  FAIL 000187: СумаЕРП={sum_erp}, очікував -44602752.10")
        FAILS += 1

if ne_znaydeno_c > 0:
    print(f"FAIL: {ne_znaydeno_c} рядків 000Ц-* досі «документ не знайдено в Бух»")
    FAILS += 1
if vhidna_for_pair > 0:
    print(f"FAIL: {vhidna_for_pair} ложних «вхідна сторона в ЕРП не зв'язана» для Надходжень")
    FAILS += 1

print("\nРЕЗУЛЬТАТ: " + ("SMOKE OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
