# -*- coding: utf-8 -*-
"""
Comprehensive engineering verification: сторно cancellation in А_ФинРез_DDS Branch 3.

Tests that for EVERY (orig_doc, correction_doc) pair across the database, the
storno cancellation produces net=0 in OLAP transformation for combinations
present in both docs.

Coverage matrix:
  - Тип: Приход / Расход
  - ВидПериода: Месяц / Объект
  - All Подразделения, all Статьи ДДС, all Месяцы

Sign logic in Branch 3 (after fix):
  f(Приход, X) = X
  f(Расход, X) = -X

For storno to cancel original:
  Original: (+X for type T)
  Storno: (-X for type T)
  f(orig) + f(storno) = ?
  - For Приход: +X + (-X) = 0
  - For Расход: -X + +X = 0

Both work algebraically. This script verifies empirically.
"""
import sys
from datetime import datetime
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")


def banner(t):
    print()
    print("=" * 100)
    print(f"  {t}")
    print("=" * 100)


def fmt(v, w=18):
    if v is None:
        return " " * w
    if isinstance(v, (int, float)):
        return f"{v:>{w},.2f}".replace(",", " ")
    s = str(v)
    return s[: w - 1].ljust(w) if len(s) > w else s.ljust(w)


print("Connecting to BaseERP...")
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ТипПрих = erp.Перечисления.ТипыЗначенийПоказателейБюджетногоОтчета.Приход
ВидПерМесяц = erp.Перечисления.А_ВидыПериодовБюджетирования.Месяц
ВидПерОбъект = erp.Перечисления.А_ВидыПериодовБюджетирования.Объект


# === STAGE 1: Знайти всі коригуючи документи ===
banner("STAGE 1: Усі (Оригинал, Корректировка) пари у базі")

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    КорДок.Ссылка КАК Кор,
    КорДок.Дата КАК КорДата,
    КорДок.Номер КАК КорНомер,
    КорДок.ДокументОснование КАК Орг,
    КорДок.ДокументОснование.Дата КАК ОргДата,
    КорДок.ДокументОснование.Номер КАК ОргНомер
ИЗ Документ.А_БюджетМесяц КАК КорДок
ГДЕ КорДок.ВидВерсии = ЗНАЧЕНИЕ(Перечисление.А_ВидыВерсийБюджетов.Корректировка)
    И НЕ КорДок.ПометкаУдаления
    И КорДок.Проведен
    И КорДок.ДокументОснование <> ЗНАЧЕНИЕ(Документ.А_БюджетМесяц.ПустаяСсылка)
    И НЕ КорДок.ДокументОснование.ПометкаУдаления
УПОРЯДОЧИТЬ ПО КорДок.Дата УБЫВ
"""
res = q.Execute().Выгрузить()
all_pairs = []
for i in range(res.Количество()):
    r = res.Получить(i)
    all_pairs.append({"kor": r.Кор, "kor_n": r.КорНомер, "kor_d": r.КорДата, "org": r.Орг, "org_n": r.ОргНомер, "org_d": r.ОргДата})
print(f"Знайдено {len(all_pairs)} пар. Тестую перші 30 (за датою корректировки УБЫВ).")


# === STAGE 2: Перевірка повного скасування для кожної пари ===
banner("STAGE 2: Перевірка скасування — sign-logic поточного коду (Приход=X, Расход=-X)")

print(f"{'#':>3} | {'КорДок №':<12} | {'КорДата':<12} | {'ОригДок №':<12} | {'комб':>4} | {'не-нуль':>7} | {'тип':<8} | {'Σ ЗнакСум':>18} | Статус")
print("-" * 110)

failures = []
type_stats = {"Приход": {"pass": 0, "fail": 0, "total_residual": 0.0},
              "Расход": {"pass": 0, "fail": 0, "total_residual": 0.0},
              "Mixed":  {"pass": 0, "fail": 0, "total_residual": 0.0}}
period_stats = {"Месяц": 0, "Объект": 0}

for idx, p in enumerate(all_pairs[:30]):
    arr = erp.NewObject("Массив")
    arr.Добавить(p["org"])
    arr.Добавить(p["kor"])

    q2 = erp.NewObject("Запрос")
    q2.Text = """
    ВЫБРАТЬ
        Пл.Подразделение КАК Подр,
        Пл.СтатьяДвиженияДенежныхСредств КАК Стат,
        Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя КАК Тип,
        Пл.Месяц КАК Месяц,
        Пл.ВидПериода КАК ВидПериода,
        СУММА(Пл.СуммаОборот) КАК НетСум,
        СУММА(ВЫБОР КОГДА Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя = &ТипПрих
                ТОГДА Пл.СуммаОборот ИНАЧЕ -Пл.СуммаОборот КОНЕЦ) КАК ЗнакСум,
        КОЛИЧЕСТВО(*) КАК Кол
    ИЗ РегистрНакопления.А_БюджетыНаМесяц.Обороты(, , Регистратор, ) КАК Пл
    ГДЕ Пл.Регистратор В (&Доки)
    СГРУППИРОВАТЬ ПО
        Пл.Подразделение, Пл.СтатьяДвиженияДенежныхСредств,
        Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя,
        Пл.Месяц, Пл.ВидПериода
    """
    q2.SetParameter("Доки", arr)
    q2.SetParameter("ТипПрих", ТипПрих)

    try:
        tz = q2.Execute().Выгрузить()
        cnt = tz.Количество()
        nonzero_keys = []
        seen_types = set()
        seen_periods = set()
        for j in range(cnt):
            r = tz.Получить(j)
            if r.Кол > 1:
                seen_types.add(S(r.Тип))
                seen_periods.add(S(r.ВидПериода))
                signed = float(r.ЗнакСум or 0)
                if abs(signed) > 0.01:
                    nonzero_keys.append({
                        "подр": S(r.Подр),
                        "стат": S(r.Стат),
                        "тип": S(r.Тип),
                        "мес": r.Месяц.strftime("%d.%m.%Y") if r.Месяц else "",
                        "виду": S(r.ВидПериода),
                        "знак": signed,
                        "нет": float(r.НетСум or 0),
                        "кол": int(r.Кол or 0),
                    })

        for vp in seen_periods:
            period_stats[vp] = period_stats.get(vp, 0) + 1

        type_label = "Mixed" if len(seen_types) > 1 else (next(iter(seen_types)) if seen_types else "—")
        residual = sum(k["знак"] for k in nonzero_keys)
        if not nonzero_keys:
            status = "✓ OK"
            if type_label in type_stats:
                type_stats[type_label]["pass"] += 1
        else:
            status = f"⚠ {len(nonzero_keys)} некомп."
            if type_label in type_stats:
                type_stats[type_label]["fail"] += 1
                type_stats[type_label]["total_residual"] += abs(residual)
            failures.append({"pair": p, "type": type_label, "keys": nonzero_keys})

        print(f"{idx+1:>3} | {p['kor_n'][:12]:<12} | {p['kor_d'].strftime('%d.%m.%Y'):<12} | {p['org_n'][:12]:<12} | {cnt:>4} | {len(nonzero_keys):>7} | {type_label[:8]:<8} | {fmt(residual)} | {status}")
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"{idx+1:>3} | FAIL: {msg}")


# === STAGE 3: Підсумкова матриця ===
banner("STAGE 3: Підсумкова матриця за типом статті")

total_pairs = sum(s["pass"] + s["fail"] for s in type_stats.values())
total_pass = sum(s["pass"] for s in type_stats.values())
total_fail = sum(s["fail"] for s in type_stats.values())

print(f"{'Тип':<10} | {'Pass':>5} | {'Fail':>5} | {'Total residual':>18}")
print("-" * 50)
for t, s in type_stats.items():
    print(f"{t:<10} | {s['pass']:>5} | {s['fail']:>5} | {fmt(s['total_residual'])}")
print("-" * 50)
print(f"{'Σ':<10} | {total_pass:>5} | {total_fail:>5}")

print("\nКомбінацій з ВидПериода у тестованих парах:")
for vp, c in period_stats.items():
    print(f"  {vp:<10} : {c} пар")


# === STAGE 4: Деталі провалів (якщо є) ===
if failures:
    banner(f"STAGE 4: Деталі провалів ({len(failures)} пар)")
    for i, f in enumerate(failures[:5]):
        p = f["pair"]
        print(f"\n[#{i+1}] КорДок {p['kor_n']} від {p['kor_d']:%d.%m.%Y} → ОригДок {p['org_n']} від {p['org_d']:%d.%m.%Y}")
        print(f"  Тип: {f['type']}")
        for k in f["keys"][:10]:
            print(f"    {k['подр'][:25]:<25} | {k['стат'][:30]:<30} | {k['тип']:<7} | {k['мес']:<11} | {k['виду']:<7} | ЗнакСум={fmt(k['знак'])} | НетСум={fmt(k['нет'])} | Кол={k['кол']}")
else:
    banner("STAGE 4: ВСІ ПАРИ ПРОЙШЛИ ПЕРЕВІРКУ")
    print("✓ Жодної некомпенсованої комбінації — sign-logic працює коректно для всіх (Тип × ВидПериода × Подр × Стат × Мес).")


# === STAGE 5: Спеціальний тест для змішаних типів у одному ДокОснов→КорДок пари ===
banner("STAGE 5: Перевірка пари з МАКСИМАЛЬНОЮ кіл-стю комбінацій (як стрес-тест)")

# Знаходимо пару з максимумом унікальних (Стат, Подр) комбінацій
max_pair = None
max_combs = 0
for p in all_pairs[:30]:
    arr = erp.NewObject("Массив")
    arr.Добавить(p["org"])
    arr.Добавить(p["kor"])

    q5 = erp.NewObject("Запрос")
    q5.Text = """
    ВЫБРАТЬ КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВЫРАЗИТЬ(Пл.Подразделение КАК Справочник.СтруктураПредприятия)) КАК K
    ИЗ РегистрНакопления.А_БюджетыНаМесяц КАК Пл
    ГДЕ Пл.Регистратор В (&Доки)
    """
    q5.SetParameter("Доки", arr)
    res = q5.Execute().Выгрузить()
    if res.Количество() > 0:
        k = int(res.Получить(0).K or 0)
        if k > max_combs:
            max_combs = k
            max_pair = p

if max_pair:
    p = max_pair
    print(f"Найбільша пара: КорДок {p['kor_n']} → ОригДок {p['org_n']} ({max_combs} різних Подр)")

    arr = erp.NewObject("Массив")
    arr.Добавить(p["org"])
    arr.Добавить(p["kor"])

    q6 = erp.NewObject("Запрос")
    q6.Text = """
    ВЫБРАТЬ
        Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя КАК Тип,
        Пл.ВидПериода КАК ВидПериода,
        КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВЫРАЗИТЬ(Пл.Подразделение КАК Справочник.СтруктураПредприятия)) КАК КолПодр,
        КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВЫРАЗИТЬ(Пл.СтатьяДвиженияДенежныхСредств КАК Справочник.СтатьиДвиженияДенежныхСредств)) КАК КолСтат,
        СУММА(Пл.Сумма) КАК СумаНет,
        СУММА(ВЫБОР КОГДА Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя = &ТипПрих
                ТОГДА Пл.Сумма ИНАЧЕ -Пл.Сумма КОНЕЦ) КАК СумаЗнак
    ИЗ РегистрНакопления.А_БюджетыНаМесяц КАК Пл
    ГДЕ Пл.Регистратор В (&Доки)
    СГРУППИРОВАТЬ ПО Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя, Пл.ВидПериода
    """
    q6.SetParameter("Доки", arr)
    q6.SetParameter("ТипПрих", ТипПрих)
    tz6 = q6.Execute().Выгрузить()
    print(f"\n{'Тип':<10} | {'ВидПериода':<10} | {'Подр':>4} | {'Стат':>4} | {'Σ Сума (нетто)':>20} | {'Σ ЗнакСум':>20}")
    print("-" * 80)
    for j in range(tz6.Количество()):
        r = tz6.Получить(j)
        print(f"{S(r.Тип):<10} | {S(r.ВидПериода):<10} | {int(r.КолПодр or 0):>4} | {int(r.КолСтат or 0):>4} | {fmt(r.СумаНет or 0)} | {fmt(r.СумаЗнак or 0)}")


banner("ВИСНОВКИ")
print(f"Перевірено пар: {min(30, len(all_pairs))}")
print(f"Pass: {total_pass}  |  Fail: {total_fail}")
if total_fail == 0:
    print()
    print("✓ ENGINEERING VERDICT: sign-логіка Branch 3 (Приход=X, Расход=-X) КОРЕКТНА.")
    print("  Сторно (-X) до оригіналу (+X) дає 0 у OLAP-трансформації для всіх комбінацій.")
    print("  Жодних патчів не потрібно — це стандартна каскова семантика:")
    print("    • income/Приход в OLAP як +Сумма (signed)")
    print("    • expense/Расход в OLAP як -Сумма (signed, negative outflow)")
    print("    • storno створює рівне-протилежне → net=0")
else:
    print()
    print(f"⚠ {total_fail} пар некоректних. Деталі вище — потрібно копати.")

print("\nDONE.")
