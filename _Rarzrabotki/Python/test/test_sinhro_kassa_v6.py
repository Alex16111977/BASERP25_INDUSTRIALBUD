import sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client
import pywintypes

v8 = win32com.client.Dispatch("V83.COMConnector")
kazna = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')

def kdate(y, m, d, hh=0, mm=0, ss=0):
    # Дата как настоящий тип 1С: pywintypes.Time (передаётся в SetParameter как тип Дата,
    # без строкового параметра → без пустого результата/tz-сдвига).
    # tzinfo не задаём (naive local) — 1С трактует как локальную дату сервера.
    import datetime as _dt
    return pywintypes.Time(_dt.datetime(y, m, d, hh, mm, ss))

print("=" * 70)
print("ЗАДАЧА V6 — Казна (BuhKazn): план/узел обмена + реверс по кассе")
print("=" * 70)

# ============================================================
# 1) ПЕРЕЧИСЛИТЬ ПЛАНЫ ОБМЕНА, найти тот что указывает на ЕРП
# ============================================================
print("\n--- 1) Планы обмена в базе Казна ---")
planNames = []
for p in kazna.Метаданные.ПланыОбмена:
    nm = kazna.String(p.Имя)
    planNames.append(nm)
    print("  ПланОбмена:", nm)

# Кандидаты по маске (ЕРП / ERP / BASERP / Казн)
candidates = []
for nm in planNames:
    low = nm.lower()
    if ("erp" in low) or ("ерп" in low) or ("baserp" in low) or ("казн" in low) or ("treasury" in low):
        candidates.append(nm)
print("\n  Кандидаты (маска ERP/ЕРП/BASERP/Казн):", candidates)

# ============================================================
#    Для каждого кандидата — узел ГДЕ НЕ ЭтотУзел
# ============================================================
print("\n--- Узлы кандидатов (ГДЕ НЕ ЭтотУзел) ---")
node_found = {}
for nm in candidates:
    q = kazna.NewObject("Запрос")
    q.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1\n"
        "    Узел.Ссылка КАК Ссылка,\n"
        "    Узел.Наименование КАК Наименование,\n"
        "    Узел.Код КАК Код\n"
        "ИЗ\n"
        "    ПланОбмена." + nm + " КАК Узел\n"
        "ГДЕ\n"
        "    НЕ Узел.ЭтотУзел"
    )
    try:
        r = q.Execute().Выгрузить()
        if r.Количество() > 0:
            s = r.Получить(0)
            uid = kazna.String(s.Ссылка.УникальныйИдентификатор())
            print(f"  [{nm}] узел: Наим='{kazna.String(s.Наименование)}' Код='{kazna.String(s.Код)}' UID={uid}")
            node_found[nm] = {
                "Наименование": kazna.String(s.Наименование),
                "Код": kazna.String(s.Код),
                "UID": uid,
            }
        else:
            print(f"  [{nm}] узлов кроме ЭтотУзел НЕТ (план пустой/локальный)")
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        msg = info[2] if info else str(e)
        print(f"  [{nm}] FAIL: {msg}")

# ============================================================
# 2) РЕВЕРС ПО КАССЕ — РН.ДенежныеСредства (Наличные), периодичность Регистратор
# ============================================================
print("\n" + "=" * 70)
print("2) Реверс по кассе — обороты по Регистратору (декабрь 2025)")
print("=" * 70)

def run_reverse(d1, d2, label):
    # Вариант A: .Обороты() с периодичностью Регистратор (имена ресурсов СуммаПриход/СуммаРасход)
    qa = kazna.NewObject("Запрос")
    qa.Text = (
        "ВЫБРАТЬ\n"
        "    Т.Регистратор КАК Регистратор,\n"
        "    ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК РегистраторПредст,\n"
        "    Т.БанковскийСчетКасса КАК Касса,\n"
        "    Т.СуммаПриход КАК Приход,\n"
        "    Т.СуммаРасход КАК Расход\n"
        "ИЗ\n"
        "    РегистрНакопления.ДенежныеСредства.Обороты(&Д1, &Д2, Регистратор,\n"
        "        ВидДенежныхСредств = ЗНАЧЕНИЕ(Перечисление.ВидыДенежныхСредств.Наличные)) КАК Т"
    )
    qa.SetParameter("Д1", d1)
    qa.SetParameter("Д2", d2)
    try:
        r = qa.Execute().Выгрузить()
        return ("A(.Обороты Регистратор)", r, qa.Text, None)
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        msgA = info[2] if info else str(e)

    # Вариант B: запись-уровневый РН.ДенежныеСредства с группировкой по Регистратору
    qb = kazna.NewObject("Запрос")
    qb.Text = (
        "ВЫБРАТЬ\n"
        "    Т.Регистратор КАК Регистратор,\n"
        "    ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК РегистраторПредст,\n"
        "    Т.БанковскийСчетКасса КАК Касса,\n"
        "    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)\n"
        "                ТОГДА Т.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Приход,\n"
        "    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)\n"
        "                ТОГДА Т.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Расход\n"
        "ИЗ\n"
        "    РегистрНакопления.ДенежныеСредства КАК Т\n"
        "ГДЕ\n"
        "    Т.Период >= &Д1 И Т.Период <= &Д2\n"
        "    И Т.Активность\n"
        "    И Т.ВидДенежныхСредств = ЗНАЧЕНИЕ(Перечисление.ВидыДенежныхСредств.Наличные)\n"
        "СГРУППИРОВАТЬ ПО\n"
        "    Т.Регистратор, Т.БанковскийСчетКасса"
    )
    qb.SetParameter("Д1", d1)
    qb.SetParameter("Д2", d2)
    try:
        r = qb.Execute().Выгрузить()
        return ("B(запись-уровень GROUP BY)", r, qb.Text, msgA)
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        msgB = info[2] if info else str(e)
        return ("FAIL", None, qb.Text, f"A: {msgA} | B: {msgB}")

# границы декабря 2025 через серверный 1С (тип Дата)
d1 = kdate(2025, 12, 1, 0, 0, 0)
d2 = kdate(2025, 12, 31, 23, 59, 59)
print(f"\nПериод: {d1} .. {d2}")

mode, res, qtext, errnote = run_reverse(d1, d2, "дек2025")
print(f"\nРежим запроса: {mode}")
if errnote:
    print("  (примечание/ошибки альтернатив):", errnote)

if res is None:
    print("  ОБА ВАРИАНТА УПАЛИ. Пробую январь 2026...")
    d1 = kdate(2026, 1, 1, 0, 0, 0)
    d2 = kdate(2026, 1, 31, 23, 59, 59)
    mode, res, qtext, errnote = run_reverse(d1, d2, "янв2026")
    print(f"  Режим (янв2026): {mode}; примечание: {errnote}")

if res is not None and res.Количество() == 0:
    print("  Декабрь пуст — пробую январь 2026...")
    d1 = kdate(2026, 1, 1, 0, 0, 0)
    d2 = kdate(2026, 1, 31, 23, 59, 59)
    mode, res, qtext, errnote = run_reverse(d1, d2, "янв2026")
    print(f"  Режим (янв2026): {mode}; кол-во={res.Количество() if res else 'FAIL'}")

print(f"\nРабочий текст запроса реверса ({mode}):")
print(qtext)

# ============================================================
#    Вывод регистраторов: тип + UUID
# ============================================================
print("\n--- Регистраторы Казны (тип / UUID), до 6 шт. ---")
samples = []
if res is not None:
    print(f"Всего строк: {res.Количество()}")
    n = min(6, res.Количество())
    for i in range(n):
        s = res.Получить(i)
        reg = s.Регистратор
        # тип через Метаданные().Имя в контексте Казны
        tip = kazna.String(reg.Метаданные().Имя)
        # UUID двумя способами
        uid_xml = kazna.XMLСтрока(reg.УникальныйИдентификатор())
        uid_str = kazna.String(reg.УникальныйИдентификатор())
        kassa_naim = kazna.String(s.Касса) if s.Касса is not None else "<пусто>"
        print(f"  [{i}] Тип={tip}")
        print(f"       Предст={kazna.String(s.РегистраторПредст)}")
        print(f"       Касса={kassa_naim}")
        print(f"       Приход={s.Приход}  Расход={s.Расход}")
        print(f"       UUID(XMLСтрока)={uid_xml}")
        print(f"       UUID(String)   ={uid_str}")
        samples.append({"тип": tip, "uid_xml": uid_xml, "uid_str": uid_str})

# --- Дополнительно: типы регистраторов и наличие РКО (Расход) во всём наборе ---
print("\n--- Все типы регистраторов в реверсе + есть ли РКО (Расход) ---")
if res is not None:
    types_all = {}
    rko_example = None
    for i in range(res.Количество()):
        s = res.Получить(i)
        t = kazna.String(s.Регистратор.Метаданные().Имя)
        types_all[t] = types_all.get(t, 0) + 1
        if rko_example is None and float(s.Расход) != 0:
            rko_example = {
                "тип": t,
                "предст": kazna.String(s.РегистраторПредст),
                "касса": kazna.String(s.Касса) if s.Касса is not None else "<пусто>",
                "приход": float(s.Приход), "расход": float(s.Расход),
                "uid_xml": kazna.XMLСтрока(s.Регистратор.УникальныйИдентификатор()),
                "uid_str": kazna.String(s.Регистратор.УникальныйИдентификатор()),
            }
    print("  Типы (тип -> кол-во строк):", types_all)
    print("  Пример строки с Расход<>0 (РКО):", rko_example)

print("\n=== ИТОГ ===")
print("Кандидаты планов обмена:", candidates)
print("Найденные узлы:", node_found)
print("Типы регистраторов в реверсе:", sorted({s['тип'] for s in samples}))
print("ГОТОВО")
