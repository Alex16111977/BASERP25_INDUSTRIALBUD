# -*- coding: utf-8 -*-
"""Нарахування відсотків за місяць документом А_НачислениеПроцентовПоКредитам (план 2026-10-16, Task 7).

    python nachislenie_procentov.py 2026-09                 # сухий прогін: документ у пам'яті, рядки і параметри ПТУ
    python nachislenie_procentov.py 2026-09 --apply         # записати документ місяця «За виплатою» (знімок «до»)
    python nachislenie_procentov.py 2026-09 --apply --sozdat  # + створити/оновити «Приобретения» відсотків
    python nachislenie_procentov.py 2026-09 --check         # лише перевірки
    python nachislenie_procentov.py 2026-06,2026-07,2026-08,2026-09 --zalishky  # перевірка залишків договорів %% = 0

Документ місяця шукається за організацією і місяцем (непомічений); немає — створюється.
Перевірки (код виходу 1 при провалі):
  - кожен рядок із сумою має проведене непомічене ПТУ з датою КонецМесяца(M), сумою = СуммаКНачислению,
    статтею «Послуги банків» і рядком послуги «Відсотки за користування позикою»;
  - ПрочиеРасходы за статтею «Послуги банків» у місяці M по підрозділах = знімок «до» + суми ПТУ;
  - --zalishky: залишки РасчетыСПоставщикамиПоСрокам по договорах %%, що мали виплати, = 0.
"""
import datetime as dt
import json
import os
import sys

import win32com.client

sys.stdout.reconfigure(encoding='utf-8')
erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
ORG = erp.Справочники.Организации.НайтиПоНаименованию('ТОВ "ІНДАСТРІАЛБУД"', True)
М = erp.А_НачислениеПроцентовПоКредитамСервер
APPLY, SOZDAT, CHECK, ZAL = ("--apply" in sys.argv, "--sozdat" in sys.argv, "--check" in sys.argv, "--zalishky" in sys.argv)
MONTHS = [tuple(int(x) for x in m.split("-")) for m in sys.argv[1].split(",")]
SNAP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_snapshots_procenty")
fails = 0


def d1c(y, m, d=1, h=0, mi=0, s=0):
    """Дата для 1С через COM (+3 год — пам'ять com-date-shift-3-hours)."""
    return dt.datetime(y, m, d, h, mi, s) + dt.timedelta(hours=3)


def konec(y, m):
    nxt = dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)
    return d1c(nxt.year, nxt.month, nxt.day, 23, 59, 59)


def q(text, **p):
    z = erp.NewObject("Запрос")
    z.Text = text
    for k, v in p.items():
        z.SetParameter(k, v)
    return z.Execute().Выгрузить()


def fail(msg):
    global fails
    fails += 1
    print("FAIL", msg)


def rashody(y, m):
    t = q("""ВЫБРАТЬ Об.Подразделение КАК Подразделение, СУММА(Об.СуммаОборот) КАК Сумма
ИЗ РегистрНакопления.ПрочиеРасходы.Обороты(&Н, &К, , СтатьяРасходов = &Ст И Организация = &Орг) КАК Об
СГРУППИРОВАТЬ ПО Об.Подразделение""", Н=d1c(y, m), К=konec(y, m), Ст=М.СтатьяРасходовПроцентов(), Орг=ORG)
    return {S(t.Получить(i).Подразделение): round(float(t.Получить(i).Сумма), 2) for i in range(t.Количество())}


def dokument(y, m):
    t = q("""ВЫБРАТЬ Д.Ссылка ИЗ Документ.А_НачислениеПроцентовПоКредитам КАК Д
ГДЕ Д.Организация = &Орг И Д.МесяцНачисления = &М И НЕ Д.ПометкаУдаления""", Орг=ORG, М=d1c(y, m))
    if t.Количество() > 1:
        sys.exit(f"{y}-{m:02}: кілька документів нарахування — розберіться вручну")
    return t.Получить(0).Получить(0) if t.Количество() else None


def proverki(y, m):
    ref = dokument(y, m)
    if ref is None:
        fail(f"{y}-{m:02}: документа немає")
        return
    d = ref.ПолучитьОбъект()
    snap_path = os.path.join(SNAP_DIR, f"rashody_{y}-{m:02}.json")
    do = json.load(open(snap_path, encoding="utf-8")) if os.path.exists(snap_path) else None
    ochik = dict(do or {})
    last = dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)
    ochik_data = f"{last.day:02}.{m:02}.{y} 23:59:59"
    for i in range(d.Начисления.Количество()):
        r = d.Начисления.Получить(i)
        suma = round(float(r.СуммаКНачислению), 2)
        if suma == 0:
            continue
        p = r.ДокументНачисления
        if not erp.ЗначениеЗаполнено(p):
            fail(f"{S(r.Договор)}: немає ПТУ")
            continue
        tov = p.Товары.Получить(0) if p.Товары.Количество() else None
        good = (p.Проведен and not p.ПометкаУдаления and S(p.Дата) == ochik_data
                and round(float(p.СуммаДокумента), 2) == suma and tov is not None
                and S(tov.СтатьяРасходов) == "Послуги банків" and S(tov.Номенклатура) == "Відсотки за користування позикою")
        print(("OK  " if good else "FAIL"), S(p), "|", S(r.Договор), "|", suma)
        if not good:
            fail(f"ПТУ {S(p)}")
        ochik[S(r.Подразделение)] = round(ochik.get(S(r.Подразделение), 0) + suma, 2)
    if do is None:
        print("(знімка ПрочиеРасходы «до» немає — перевірку витрат пропущено)")
        return
    pislya = rashody(y, m)
    for podr in sorted(set(ochik) | set(pislya)):
        if abs(ochik.get(podr, 0) - pislya.get(podr, 0)) > 0.005:
            fail(f"ПрочиеРасходы «Послуги банків» {podr}: очікувано {ochik.get(podr, 0)}, факт {pislya.get(podr, 0)}")
    print(f"ПрочиеРасходы «Послуги банків» {y}-{m:02}: {len(pislya)} підрозділів, звірено зі знімком «до» + ПТУ")


for y, m in MONTHS:
    if CHECK or ZAL:
        break
    ref = dokument(y, m)
    d = ref.ПолучитьОбъект() if ref is not None else erp.Документы.А_НачислениеПроцентовПоКредитам.СоздатьДокумент()
    if ref is None:
        d.Организация = ORG
        d.МесяцНачисления = d1c(y, m)
        d.Дата = konec(y, m)
        d.Ответственный = erp.Справочники.Пользователи.НайтиПоНаименованию("Администратор", True)
        d.Комментарий = "Відсотки за виплатами наступного місяця (16.10.2026)"
    d.СпособЗаполнения = erp.Перечисления.А_СпособыРасчетаПроцентов.ПоВыплате
    preds = d.ЗаполнитьНачисления()
    print(f"== {y}-{m:02}: документ {S(ref) if ref is not None else '(новий)'}; рядків {d.Начисления.Количество()}, "
          f"Σ до нарахування {d.Начисления.Итог('СуммаКНачислению')}; непереоформлених виплат {preds.Количество()}")
    for i in range(d.Начисления.Количество()):
        r = d.Начисления.Получить(i)
        print(f"   {S(r.Договор)} | {S(r.Подразделение)} | виплачено {S(r.СуммаВыплат)} | формула {S(r.СуммаПоФормуле)} "
              f"| до нарахування {S(r.СуммаКНачислению)}")
    if not APPLY:
        r = d.Начисления.Получить(0)
        ш = М.ПараметрыПриобретения(d, r).Шапка
        print(f"   приклад ПТУ: {S(ш.Дата)} | {S(ш.Контрагент)} | {S(ш.Договор)} | {S(ш.Подразделение)} | {S(ш.СуммаДокумента)}")
        continue
    os.makedirs(SNAP_DIR, exist_ok=True)
    snap_path = os.path.join(SNAP_DIR, f"rashody_{y}-{m:02}.json")
    if not os.path.exists(snap_path):
        json.dump(rashody(y, m), open(snap_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("   знімок ПрочиеРасходы «до»:", snap_path)
    d.Записать(erp.РежимЗаписиДокумента.Запись)
    print("   документ записано:", S(d.Ссылка))
    if SOZDAT:
        d.СоздатьНачисленияНаСервере()
        print("   «Приобретения» створено/оновлено")

if APPLY or CHECK:
    for y, m in MONTHS:
        proverki(y, m)

if ZAL:
    t = q("""ВЫБРАТЬ О.ОбъектРасчетов.Договор КАК Договор, СУММА(О.ДолгРеглОстаток) КАК Долг, СУММА(О.ПредоплатаРеглОстаток) КАК Пред
ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам.Остатки(,
  ВЫРАЗИТЬ(ОбъектРасчетов.Договор КАК Справочник.ДоговорыКонтрагентов).А_ВидКредитногоДоговора = ЗНАЧЕНИЕ(Перечисление.А_ВидыКредитныхДоговоров.Проценты)) КАК О
СГРУППИРОВАТЬ ПО О.ОбъектРасчетов.Договор""")
    for i in range(t.Количество()):
        r = t.Получить(i)
        good = abs(float(r.Долг)) < 0.005 and abs(float(r.Пред)) < 0.005
        print(("OK  " if good else "FAIL"), S(r.Договор), "| борг", S(r.Долг), "| аванс", S(r.Пред))
        if not good:
            fail(f"залишок {S(r.Договор)}")
    print("договорів %% із залишком:", t.Количество())

print("ПРОВАЛОВ:", fails)
sys.exit(1 if fails else 0)
