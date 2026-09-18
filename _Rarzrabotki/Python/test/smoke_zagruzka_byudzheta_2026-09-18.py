# -*- coding: utf-8 -*-
"""Smoke собранной .epf «А_ЗагрузкаБюджетаПроизводства» + прогон движка на живой базе.

Excel читается ЗДЕСЬ ровно так, как это делает клиентская форма обработки, и отдаётся
в движок — проверяется тот самый код, который поедет в бою.

  (без ключей)   разбор книги и раскладка листов, записи в базу НЕТ
  --preview      плюс предпросмотр: что движок соберёт по каждому листу и месяцу (без записи)
  --load         загрузка: создаёт/перезаполняет документы (сценарий создаётся при отсутствии)
  --no-post      не проводить документы
  --sheets A,B   только эти листы
"""
import sys, os, json, collections
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

EPF = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\production-budget-load-e9afc9\_Rarzrabotki\Обработки\А_ЗагрузкаБюджетаПроизводства.epf"
BOOK = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Отчеты высланные\ДДС Производство Вересень - Грудень 2026 Прогноз.xlsx"
SCENARIO = "Бюджет экономиста ДДС"
ROOT_CODE = "0Ц-000004"
HERE = os.path.dirname(os.path.abspath(__file__))

LOAD = "--load" in sys.argv
PREVIEW = "--preview" in sys.argv
POST = "--no-post" not in sys.argv
ONLY = None
for i, a in enumerate(sys.argv):
    if a == "--sheets" and i + 1 < len(sys.argv):
        ONLY = [s.strip() for s in sys.argv[i + 1].split(",")]

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def fmt(x):
    try:
        return f"{float(x):,.2f}".replace(",", " ")
    except Exception:
        return str(x)


print("=== Загрузка .epf ===")
обр = erp.ВнешниеОбработки.Создать(EPF, False)
мд = обр.Метаданные()
безтипа = []
for i in range(мд.Реквизиты.Количество()):
    р = мд.Реквизиты.Получить(i)
    типы = [S(р.Тип.Типы().Получить(j)) for j in range(р.Тип.Типы().Количество())]
    print(f"  реквизит {р.Имя}: {типы}")
    if not типы:
        безтипа.append(р.Имя)
assert not безтипа, f"реквизиты БЕЗ ТИПА (Произвольный): {безтипа}"
тч = мд.ТабличныеЧасти.Получить(0)
колонки = [тч.Реквизиты.Получить(i).Имя for i in range(тч.Реквизиты.Количество())]
print(f"  ТЧ {тч.Имя}: {колонки}")
for i in range(тч.Реквизиты.Количество()):
    р = тч.Реквизиты.Получить(i)
    assert р.Тип.Типы().Количество() > 0, f"колонка ТЧ {р.Имя} без типа"

обр.ПутьКФайлу = BOOK
обр.ПроводитьДокументы = POST
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ П.Ссылка КАК Ссылка ИЗ Справочник.СтруктураПредприятия КАК П ГДЕ П.Код = &Код"
q.SetParameter("Код", ROOT_CODE)
sel = q.Execute().Выбрать()
sel.Следующий()
обр.КореньНаправления = sel.Ссылка
обр.Сценарий = обр.СценарийПоНаименованию(SCENARIO, LOAD)
print(f"  сценарий: «{S(обр.Сценарий)}»; корень: {S(обр.КореньНаправления)}")
if LOAD:
    assert erp.ЗначениеЗаполнено(обр.Сценарий), "сценарий не получен"

print("\n=== Чтение книги через Excel COM ===")
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(BOOK, False, True)


def значения_ряда(ws, row, last_col):
    d = erp.NewObject("Соответствие")
    for c in range(1, last_col + 1):
        v = ws.Cells(row, c).Value
        if v is None:
            continue
        if isinstance(v, str) or hasattr(v, "year"):
            d.Вставить(c, v)
    return d


def число(v):
    if v is None or isinstance(v, bool) or hasattr(v, "year"):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(t)
    except Exception:
        return 0.0


def имя_строки(ws, row):
    for c in (1, 2, 3):
        v = ws.Cells(row, c).Value
        if v is not None and str(v).strip():
            return str(v).strip()
    return ""


def разобрать(ws):
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    last_col = used.Column + used.Columns.Count - 1
    res = {"Блоки": [], "Строки": erp.NewObject("Массив"), "СтрокСДанными": 0, "Замечание": ""}
    if last_row < 6 or last_col < 3:
        return res
    разбор = обр.БлокиМесяцев(значения_ряда(ws, 3, last_col), значения_ряда(ws, 4, last_col))
    блоки = [(разбор.Блоки.Получить(i).Колонка, разбор.Блоки.Получить(i).Месяц)
             for i in range(разбор.Блоки.Количество())]
    res["Блоки"] = блоки
    if разбор.Отброшено.Количество() > 0:
        res["Замечание"] = "отброшено блоков: " + str(разбор.Отброшено.Количество())
    if not блоки:
        return res
    for row in range(6, last_row + 1):
        значения = erp.NewObject("Массив")
        есть = False
        for col, месяц in блоки:
            bn = число(ws.Cells(row, col).Value)
            nal = число(ws.Cells(row, col + 1).Value)
            if bn == 0 and nal == 0:
                continue
            st = erp.NewObject("Структура", "Месяц, Безнал, Нал")
            st.Месяц = месяц
            st.Безнал = bn
            st.Нал = nal
            значения.Добавить(st)
            есть = True
        имя = имя_строки(ws, row)
        if not имя and not есть:
            continue
        стр = erp.NewObject("Структура", "Ряд, Имя, Уровень, Значения")
        стр.Ряд = row
        стр.Имя = имя
        стр.Уровень = int(ws.Rows(row).OutlineLevel)
        стр.Значения = значения
        res["Строки"].Добавить(стр)
        if есть:
            res["СтрокСДанными"] += 1
    return res


данные = erp.NewObject("Массив")
разборы = {}
for i in range(1, wb.Sheets.Count + 1):
    ws = wb.Sheets(i)
    name = str(ws.Name)
    if ONLY and name not in ONLY:
        continue
    r = разобрать(ws)
    разборы[name] = r
    d = erp.NewObject("Структура", "ИмяЛиста, Скрыт, Месяцев, СтрокСДанными, Замечание")
    d.ИмяЛиста = name
    d.Скрыт = (ws.Visible != -1)
    d.Месяцев = len(r["Блоки"])
    d.СтрокСДанными = r["СтрокСДанными"]
    d.Замечание = r["Замечание"]
    данные.Добавить(d)
    print(f"  «{name}»: блоки {[str(m)[:10] for _, m in r['Блоки']]}, строк с данными {r['СтрокСДанными']}")

обр.ЗаполнитьТаблицуЛистов(данные)
print(f"\n=== ТЧ Листы: {обр.Листы.Количество()} строк ===")
плановые = []
for i in range(обр.Листы.Количество()):
    s = обр.Листы.Получить(i)
    mark = "ГРУЗИМ" if s.Загружать else "  -   "
    print(f"  {mark} | {s.ИмяЛиста:<26} | {s.Роль:<8} | {S(s.Подразделение):<26} | мес {s.Месяцев} | строк {s.СтрокСДанными} | {s.Примечание[:72]}")
    if s.Загружать:
        плановые.append(s.ИмяЛиста)
print("  отмечено к загрузке:", len(плановые))

if PREVIEW:
    print("\n=== ПРЕДПРОСМОТР (без записи) ===")
    сводка = {}
    for i in range(обр.Листы.Количество()):
        s = обр.Листы.Получить(i)
        if not s.Загружать or not erp.ЗначениеЗаполнено(s.Подразделение):
            continue
        итог = обр.НовыйИтогЛиста(s.ИмяЛиста)
        тз = обр.СобратьТаблицуБюджета(разборы[s.ИмяЛиста]["Строки"], итог)
        помесячно = collections.defaultdict(lambda: [0.0, 0.0, 0])
        for k in range(тз.Количество()):
            r = тз.Получить(k)
            m = str(r.Месяц)[:10]
            d = помесячно[m]
            if S(r.ВидОплаты) == "Безналичные":
                d[0] += float(r.Сумма)
            else:
                d[1] += float(r.Сумма)
            d[2] += 1
        сводка[s.ИмяЛиста] = {
            "подразделение": S(s.Подразделение),
            "помесячно": {m: {"безнал": round(v[0], 2), "нал": round(v[1], 2), "строк": v[2]}
                          for m, v in помесячно.items()},
            "не_найдены": [итог.НеНайденныеСтатьи.Получить(k) for k in range(итог.НеНайденныеСтатьи.Количество())],
            "нестрогие": [итог.НестрогиеСтатьи.Получить(k) for k in range(итог.НестрогиеСтатьи.Количество())],
            "отсечено": [итог.Отсечено.Получить(k) for k in range(итог.Отсечено.Количество())],
            "хвост": итог.ХвостСоСтроки}
        print(f"  «{s.ИмяЛиста}» -> {S(s.Подразделение)}, хвост со строки {итог.ХвостСоСтроки}")
        for m in sorted(помесячно):
            v = помесячно[m]
            print(f"      {m}: безнал {fmt(v[0]):>16}  нал {fmt(v[1]):>14}  строк ТЧ {v[2]:>3}")
        for x in сводка[s.ИмяЛиста]["не_найдены"]:
            print("      НЕ НАЙДЕНА СТАТЬЯ:", x)
        for x in сводка[s.ИмяЛиста]["нестрогие"]:
            print("      нестрого:", x)
        for x in сводка[s.ИмяЛиста]["отсечено"]:
            print("      отсечено:", x)
    json.dump(сводка, open(os.path.join(HERE, "preview_result.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1, default=str)
    print("\n  Предпросмотр сохранён: preview_result.json")

if LOAD:
    print(f"\n=== ЗАГРУЗКА (проводить: {POST}) ===")
    итоги = {}
    всего_док = всего_строк = 0
    всего_сумма = 0.0
    for i in range(обр.Листы.Количество()):
        s = обр.Листы.Получить(i)
        if not s.Загружать or not erp.ЗначениеЗаполнено(s.Подразделение):
            continue
        итог = обр.ЗагрузитьЛист(s.ИмяЛиста, s.Подразделение, разборы[s.ИмяЛиста]["Строки"],
                                 os.path.basename(BOOK))
        if итог.Ошибка:
            print(f"  ОШИБКА «{s.ИмяЛиста}»: {итог.Ошибка}")
            continue
        docs = []
        for j in range(итог.Документы.Количество()):
            d = итог.Документы.Получить(j)
            docs.append({"месяц": str(d.Месяц)[:10], "документ": d.Документ, "строк": d.Строк,
                         "сумма": float(d.Сумма), "создан": bool(d.Создан),
                         "проведен": bool(d.Проведен), "ошибка": d.Ошибка})
            всего_док += 1
            всего_строк += d.Строк
            всего_сумма += float(d.Сумма)
            метка = "создан " if d.Создан else "обновлён"
            хвост = f" ОШИБКА: {d.Ошибка}" if d.Ошибка else (", проведён" if d.Проведен else ", НЕ проведён")
            print(f"  {метка} {s.ИмяЛиста:<24} {str(d.Месяц)[:10]}: строк {d.Строк:>3}, сумма {fmt(d.Сумма):>16}{хвост}")
        итоги[s.ИмяЛиста] = {
            "подразделение": S(s.Подразделение),
            "документы": docs,
            "не_найдены": [итог.НеНайденныеСтатьи.Получить(k) for k in range(итог.НеНайденныеСтатьи.Количество())],
            "нестрогие": [итог.НестрогиеСтатьи.Получить(k) for k in range(итог.НестрогиеСтатьи.Количество())],
            "отсечено": [итог.Отсечено.Получить(k) for k in range(итог.Отсечено.Количество())],
            "хвост": итог.ХвостСоСтроки}
    print(f"\n  ИТОГО документов {всего_док}, строк {всего_строк}, сумма {fmt(всего_сумма)}")
    for k, v in итоги.items():
        if v["не_найдены"] or v["отсечено"] or v["нестрогие"]:
            print(f"\n  «{k}»:")
            for x in v["не_найдены"]:
                print("     НЕ НАЙДЕНА СТАТЬЯ:", x)
            for x in v["нестрогие"]:
                print("     нестрого:", x)
            for x in v["отсечено"]:
                print("     отсечено:", x)
    json.dump(итоги, open(os.path.join(HERE, "load_result.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1, default=str)
    print("\n  Результат сохранён: load_result.json")

wb.Close(False)
xl.Quit()
print("\nГОТОВО")
