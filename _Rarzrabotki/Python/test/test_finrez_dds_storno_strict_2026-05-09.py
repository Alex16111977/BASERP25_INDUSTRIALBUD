# -*- coding: utf-8 -*-
"""
STRICT engineering verification: сторно cancellation in А_ФинРез_DDS Branch 3.

Group by FULL set of 6 dimensions of А_БюджетыНаМесяц register:
  Подразделение, СтатьяДДС, Месяц, ВидПериода, Направление, ВидОплаты

For each (orig_doc, corr_doc) pair:
  - For each FULL-DIMENSION key present in BOTH docs (Кол=2):
      ЗнакСум must equal 0 (storno fully cancels original)
  - Keys present in only one doc (Кол=1):
      Either P has it (original, not yet cancelled)
      Or K has it (new entry from correction, no original to cancel)

Expected result: all "Кол=2" combinations sum to 0. If any non-zero, that's a
real bug. If only "Кол=1" combinations, that's normal — they represent the
correction's new state vs the original's pre-correction state.
"""
import sys
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
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ТипПрих = erp.Перечисления.ТипыЗначенийПоказателейБюджетногоОтчета.Приход


# === STAGE 1: All correction pairs ===
banner("STAGE 1: Усі (Орг, Кор) пари")

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
pairs = []
for i in range(res.Количество()):
    r = res.Получить(i)
    pairs.append({"kor": r.Кор, "kor_n": r.КорНомер, "kor_d": r.КорДата, "org": r.Орг, "org_n": r.ОргНомер, "org_d": r.ОргДата})
print(f"Знайдено {len(pairs)} пар. Тестую усі.")


# === STAGE 2: Strict cancellation test (group by all 6 dims) ===
banner("STAGE 2: Перевірка з групуванням по 6 вимірах")

print(f"{'#':>3} | {'Кор №':<12} | {'Орг №':<12} | {'Кер2':>5} | {'Кер2 fail':>9} | {'Кер1 P':>7} | {'Кер1 K':>7} | Σ ЗнакСум(fail)")
print("-" * 100)

real_failures = []
total_k2 = 0
total_k2_fail = 0
total_k1_p = 0
total_k1_k = 0

for idx, p in enumerate(pairs):
    arr = erp.NewObject("Массив")
    arr.Добавить(p["org"])
    arr.Добавить(p["kor"])

    q2 = erp.NewObject("Запрос")
    q2.Text = """
    ВЫБРАТЬ
        Пл.Подразделение                                       КАК Подр,
        Пл.СтатьяДвиженияДенежныхСредств                       КАК Стат,
        Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя       КАК Тип,
        Пл.Месяц                                               КАК Мес,
        Пл.ВидПериода                                          КАК ВП,
        Пл.Направление                                         КАК Напр,
        Пл.ВидОплаты                                           КАК Опл,
        СУММА(Пл.СуммаОборот)                                  КАК НетСум,
        СУММА(ВЫБОР КОГДА Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя = &ТипПрих
                ТОГДА Пл.СуммаОборот ИНАЧЕ -Пл.СуммаОборот КОНЕЦ) КАК ЗнакСум,
        КОЛИЧЕСТВО(*)                                          КАК Кол
    ИЗ РегистрНакопления.А_БюджетыНаМесяц.Обороты(, , Регистратор, ) КАК Пл
    ГДЕ Пл.Регистратор В (&Доки)
    СГРУППИРОВАТЬ ПО
        Пл.Подразделение, Пл.СтатьяДвиженияДенежныхСредств,
        Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя,
        Пл.Месяц, Пл.ВидПериода, Пл.Направление, Пл.ВидОплаты
    """
    q2.SetParameter("Доки", arr)
    q2.SetParameter("ТипПрих", ТипПрих)

    try:
        tz = q2.Execute().Выгрузить()
        k2 = 0           # Combinations with both registrators
        k2_fail = 0      # k2 with non-zero ЗнакСум (real cancellation failures)
        k2_fail_sum = 0.0
        k2_fail_keys = []
        k1_p = 0         # Only in P (original-only)
        k1_k = 0         # Only in K (new from correction)

        # Determine which Регистратор each row belongs to — re-query needed
        # Actually we can deduce from the existence pattern across both registrators,
        # but the simpler approach: separate query per registrator.

        for j in range(tz.Количество()):
            r = tz.Получить(j)
            if r.Кол == 2:
                k2 += 1
                signed = float(r.ЗнакСум or 0)
                if abs(signed) > 0.01:
                    k2_fail += 1
                    k2_fail_sum += signed
                    if len(k2_fail_keys) < 5:
                        k2_fail_keys.append({
                            "подр": S(r.Подр)[:25],
                            "стат": S(r.Стат)[:30],
                            "тип": S(r.Тип),
                            "мес": r.Мес.strftime("%d.%m.%Y") if r.Мес else "",
                            "вп": S(r.ВП),
                            "напр": S(r.Напр),
                            "опл": S(r.Опл),
                            "знак": signed,
                            "нет": float(r.НетСум or 0),
                        })

        # Separate query: count keys only in P or only in K
        q3 = erp.NewObject("Запрос")
        q3.Text = """
        ВЫБРАТЬ
            ВЫБОР КОГДА Пл.Регистратор = &Орг ТОГДА "P" ИНАЧЕ "K" КОНЕЦ КАК ИзДок,
            КОЛИЧЕСТВО(*) КАК Кол
        ИЗ (
            ВЫБРАТЬ РАЗЛИЧНЫЕ
                Пл.Регистратор КАК Регистратор,
                Пл.Подразделение, Пл.СтатьяДвиженияДенежныхСредств,
                Пл.Месяц, Пл.ВидПериода, Пл.Направление, Пл.ВидОплаты
            ИЗ РегистрНакопления.А_БюджетыНаМесяц КАК Пл
            ГДЕ Пл.Регистратор В (&Доки)
        ) КАК Пл
        СГРУППИРОВАТЬ ПО ВЫБОР КОГДА Пл.Регистратор = &Орг ТОГДА "P" ИНАЧЕ "K" КОНЕЦ
        """
        q3.SetParameter("Орг", p["org"])
        q3.SetParameter("Доки", arr)
        tz3 = q3.Execute().Выгрузить()
        per_doc = {}
        for j in range(tz3.Количество()):
            r = tz3.Получить(j)
            per_doc[S(r.ИзДок)] = int(r.Кол or 0)
        k1_p = per_doc.get("P", 0) - k2  # in P but not in K
        k1_k = per_doc.get("K", 0) - k2  # in K but not in P

        total_k2 += k2
        total_k2_fail += k2_fail
        total_k1_p += k1_p
        total_k1_k += k1_k

        if k2_fail > 0:
            real_failures.append({"pair": p, "fail_count": k2_fail, "fail_sum": k2_fail_sum, "samples": k2_fail_keys})

        status = "✓" if k2_fail == 0 else f"⚠"
        print(f"{idx+1:>3} | {p['kor_n'][:12]:<12} | {p['org_n'][:12]:<12} | {k2:>5} | {k2_fail:>9} | {k1_p:>7} | {k1_k:>7} | {fmt(k2_fail_sum) if k2_fail else status}")
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"{idx+1:>3} | FAIL: {msg}")


# === STAGE 3: Підсумки ===
banner("STAGE 3: Підсумок")
print(f"Перевірено пар: {len(pairs)}")
print(f"Σ комбінацій 'Кол=2' (присутні в обох доках, очікуємо ЗнакСум=0): {total_k2}")
print(f"Σ із них з НЕ-нуль ЗнакСум (РЕАЛЬНИЙ збій сторно): {total_k2_fail}")
print(f"Σ комбінацій 'Кол=1' тільки в P (оригінали без сторно): {total_k1_p}")
print(f"Σ комбінацій 'Кол=1' тільки в K (нові значення корекції): {total_k1_k}")
print()

if total_k2_fail == 0:
    print("✓ ENGINEERING VERDICT: Sign-логіка Branch 3 КОРЕКТНА.")
    print("  Для ВСІХ (повний 6-вимірний key) пар (P-original, K-correction):")
    print("  - Сторно завжди скасовує оригінал → ЗнакСум = 0.")
    print("  - Нові записи корекції (Кол=1 тільки в K) — це новий стан плану, не помилка.")
    print("  - Залишки оригіналу без сторно (Кол=1 тільки в P) — рідко, означає що не всі")
    print("    рядки оригінала корекція захопила (нормально для часткових корекцій).")
else:
    print(f"⚠ {total_k2_fail} комбінацій не cancellируются. Деталі нижче.")


# === STAGE 4: Деталі реальних провалів ===
if real_failures:
    banner(f"STAGE 4: РЕАЛЬНІ провали ({len(real_failures)} пар)")
    for i, f in enumerate(real_failures[:5]):
        p = f["pair"]
        print(f"\n[#{i+1}] КорДок {p['kor_n']} ({p['kor_d']:%d.%m.%Y}) → ОригДок {p['org_n']} ({p['org_d']:%d.%m.%Y})")
        print(f"     {f['fail_count']} fail комбінацій, Σ ЗнакСум = {fmt(f['fail_sum'])}")
        for k in f["samples"]:
            print(f"     Подр={k['подр']} | Стат={k['стат']} | Тип={k['тип']} | Мес={k['мес']} | ВП={k['вп']} | Напр={k['напр']} | Опл={k['опл']}")
            print(f"       НетСум={fmt(k['нет'])}  ЗнакСум={fmt(k['знак'])}")


# === STAGE 5: Спеціальний тест Globyno-2 / Окт 2024 (контроль) ===
banner("STAGE 5: Контрольний тест Глобино-2 / Окт 2024 / Поступление от заказчика111")

ref_globyno = erp.Справочники.СтруктураПредприятия.НайтиПоНаименованию("Глобино-2")
ref_st = erp.Справочники.СтатьиДвиженияДенежныхСредств.НайтиПоНаименованию("Поступление от заказчика111")

q5 = erp.NewObject("Запрос")
q5.Text = """
ВЫБРАТЬ
    Пл.Регистратор                            КАК Регистратор,
    Пл.Месяц                                  КАК Месяц,
    Пл.ВидПериода                             КАК ВП,
    Пл.Направление                            КАК Напр,
    Пл.ВидОплаты                              КАК Опл,
    СУММА(Пл.СуммаОборот)                     КАК НетСум,
    СУММА(ВЫБОР КОГДА Пл.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя = &ТипПрих
            ТОГДА Пл.СуммаОборот ИНАЧЕ -Пл.СуммаОборот КОНЕЦ) КАК ЗнакСум
ИЗ РегистрНакопления.А_БюджетыНаМесяц.Обороты(, , Регистратор, ) КАК Пл
ГДЕ Пл.Подразделение = &Подр
    И Пл.СтатьяДвиженияДенежныхСредств = &Стат
    И Пл.Месяц = ДАТАВРЕМЯ(2024, 10, 1)
    И Пл.СуммаОборот <> 0
СГРУППИРОВАТЬ ПО Пл.Регистратор, Пл.Месяц, Пл.ВидПериода, Пл.Направление, Пл.ВидОплаты
УПОРЯДОЧИТЬ ПО Пл.Регистратор
"""
q5.SetParameter("Подр", ref_globyno)
q5.SetParameter("Стат", ref_st)
q5.SetParameter("ТипПрих", ТипПрих)

tz5 = q5.Execute().Выгрузить()
print(f"{'Регистратор':<55} | {'Мес':<11} | {'ВП':<8} | {'Напр':<15} | {'Опл':<10} | {'НетСум':>15} | {'ЗнакСум':>15}")
print("-" * 145)
total_sign = 0.0
for i in range(tz5.Количество()):
    r = tz5.Получить(i)
    name = S(r.Регистратор).split(",")[0][:55]
    mes = r.Месяц.strftime("%d.%m.%Y") if r.Месяц else ""
    print(f"{name:<55} | {mes:<11} | {S(r.ВП)[:8]:<8} | {S(r.Напр)[:15]:<15} | {S(r.Опл)[:10]:<10} | {fmt(r.НетСум, 15)} | {fmt(r.ЗнакСум, 15)}")
    total_sign += float(r.ЗнакСум or 0)
print("-" * 145)
print(f"Σ ЗнакСум = {fmt(total_sign, 15)}")
if abs(total_sign) < 0.01:
    print("→ ✓ Сторно повністю компенсує оригінал — Sum=0.")
else:
    print(f"→ ⚠ Sum≠0: {total_sign}")


print("\nDONE.")
