# -*- coding: utf-8 -*-
"""
Диагностика цепочки распространения сторно для Бюджет 000002597:
  1С регистр А_БюджетыНаМесяц (raw)
    ↓ ОбработкаПроведения А_ФинРез_DDS
  1С регистр А_ОтчетDDS_Свод
    ↓ Python ETL Ai_Olap
  SQL OlapBASERP.Fact_Cashflow
    ↓ Power BI Refresh
  PL.pbix матрица

Каждый шаг проверяется отдельно. Где разрыв — там пользователю нужно
выполнить обновление.

Контрольная сторно-запись:
  Регистратор: Бюджет 000002597 от 16.05.2025
  Подразделение: Глобино-2
  СтатьяДДС: Поступление от заказчика111
  Месяц: 01.10.2024
  ВидПериода: Объект
  Сумма: -410 968 634 (storno)
"""
import sys
import win32com.client
import pyodbc

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


print("Connecting to BaseERP via COM...")
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ref_globyno = erp.Справочники.СтруктураПредприятия.НайтиПоНаименованию("Глобино-2")
ref_st = erp.Справочники.СтатьиДвиженияДенежныхСредств.НайтиПоНаименованию("Поступление от заказчика111")

q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_БюджетМесяц ГДЕ Номер = &Н УПОРЯДОЧИТЬ ПО Дата УБЫВ"
q.SetParameter("Н", "000002597")
ref_b2597 = q.Execute().Выгрузить().Получить(0).Ссылка


# === STEP 1: А_БюджетыНаМесяц (raw register) — сторно повинна бути ===
banner("STEP 1: РегНакоп.А_БюджетыНаМесяц — Бюджет 000002597 / Глобино-2 / Поступл.заказчика111 / Месяц=Окт2024")

q1 = erp.NewObject("Запрос")
q1.Text = """
ВЫБРАТЬ
    Р.НомерСтроки КАК НомерСтроки,
    Р.Период      КАК Период,
    Р.Месяц       КАК Месяц,
    Р.ВидПериода  КАК ВидПериода,
    Р.Сумма       КАК Сумма
ИЗ РегистрНакопления.А_БюджетыНаМесяц КАК Р
ГДЕ Р.Регистратор = &Рег
    И Р.Подразделение = &Подр
    И Р.СтатьяДвиженияДенежныхСредств = &Стат
    И Р.Месяц = ДАТАВРЕМЯ(2024, 10, 1)
"""
q1.SetParameter("Рег", ref_b2597)
q1.SetParameter("Подр", ref_globyno)
q1.SetParameter("Стат", ref_st)
tz1 = q1.Execute().Выгрузить()
step1_ok = False
for i in range(tz1.Количество()):
    r = tz1.Получить(i)
    print(f"  Стр {int(r.НомерСтроки):>3} | Период={S(r.Период)} | Месяц={r.Месяц.strftime('%d.%m.%Y')} | ВидПер={S(r.ВидПериода)} | Сумма={fmt(r.Сумма)}")
    if float(r.Сумма) < 0:
        step1_ok = True
print(f"\nРезультат STEP 1: {'✓ Сторно ЗНАЙДЕНО у регістрі' if step1_ok else '✗ Сторно ВІДСУТНЯ у регістрі (бажано перевірити А_БюджетМесяц провеsенням)'}")


# === STEP 2: А_ОтчетDDS_Свод — після BSL fix і перепроведення OLAP-документа ===
banner("STEP 2: РегСв.А_ОтчетDDS_Свод — рухи від Бюджет 000002597 для Глобино-2 / Поступл.заказчика111")

q2 = erp.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ
    Р.Регистратор             КАК ОЛАПДок,
    Р.Регистратор.Месяц       КАК ОЛАПМесяц,
    Р.Документ_Источник       КАК ИсточникДок,
    Р.Источник                КАК Источник,
    Р.СуммаПлан               КАК СуммаПлан,
    Р.СуммаПланОбъект         КАК СуммаПланОбъект,
    Р.СуммаКазна              КАК СуммаКазна,
    Р.Сумма                   КАК Сумма
ИЗ РегистрСведений.А_ОтчетDDS_Свод КАК Р
ГДЕ Р.Документ_Источник = &Бюджет
    И Р.Подразделение = &Подр
    И Р.СтатьяДвиженияДенежныхСредств = &Стат
"""
try:
    q2.SetParameter("Бюджет", ref_b2597)
    q2.SetParameter("Подр", ref_globyno)
    q2.SetParameter("Стат", ref_st)
    tz2 = q2.Execute().Выгрузить()
    step2_storno = 0.0
    step2_normal = 0.0
    for i in range(tz2.Количество()):
        r = tz2.Получить(i)
        plan_obj = float(r.СуммаПланОбъект or 0)
        if plan_obj < 0:
            step2_storno += plan_obj
        else:
            step2_normal += plan_obj
        print(f"  ОЛАП={S(r.ОЛАПДок).split(',')[0][:50]} | Мес={r.ОЛАПМесяц.strftime('%d.%m.%Y') if r.ОЛАПМесяц else ''} | Источник={S(r.Источник)} | СуммаПланОбъект={fmt(r.СуммаПланОбъект or 0)}")
    if tz2.Количество() == 0:
        print("  (немає рухів)")
        step2_ok = False
    else:
        step2_ok = step2_storno < -0.01  # storno -410M expected
        print(f"\nΣ позитив = {fmt(step2_normal)}, Σ сторно = {fmt(step2_storno)}")
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
    print(f"FAIL: {msg}")
    step2_ok = False

print(f"\nРезультат STEP 2: {'✓ Сторно є у А_ОтчетDDS_Свод (BSL fix працює, OLAP-доку перепроведено)' if step2_ok else '✗ Сторно ВІДСУТНЯ у А_ОтчетDDS_Свод — потрібно перепровести OLAP-документ за Жовтень 2024'}")


# === STEP 2b: Перевірити які OLAP-документи провelи для Жовтня 2024 ===
banner("STEP 2b: OLAP-документи (А_ФинРез_DDS) які захоплюють Жовтень 2024 (Месяц = 01.10.2024)")

q2b = erp.NewObject("Запрос")
q2b.Text = """
ВЫБРАТЬ
    Д.Ссылка КАК Док,
    Д.Номер  КАК Номер,
    Д.Дата   КАК Дата,
    Д.Месяц  КАК Месяц,
    Д.Проведен КАК Проведен,
    Д.ПометкаУдаления КАК Удалено
ИЗ Документ.А_ФинРез_DDS КАК Д
ГДЕ Д.Месяц = ДАТАВРЕМЯ(2024, 10, 1)
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ
"""
tz2b = q2b.Execute().Выгрузить()
if tz2b.Количество() == 0:
    print("  ✗ Жодного OLAP-документа за Жовтень 2024 НЕ існує. Створіть новий.")
else:
    for i in range(tz2b.Количество()):
        r = tz2b.Получить(i)
        status = "✓ проведен" if r.Проведен else "✗ НЕ проведен"
        if r.Удалено:
            status = "✗ помічений на видалення"
        print(f"  №{r.Номер} от {r.Дата.strftime('%d.%m.%Y %H:%M')} | Месяц={r.Месяц.strftime('%d.%m.%Y')} | {status}")


# === STEP 3: SQL OlapBASERP.Fact_Cashflow ===
banner("STEP 3: SQL OlapBASERP.Fact_Cashflow — Бюджет 000002597 / Глобино-2 / Поступл.заказчика111")

try:
    sql_conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=localhost;"
        "DATABASE=OlapBASERP;"
        "UID=sa;PWD=Brw739182465!;"
        "TrustServerCertificate=yes;"
    )
    cur = sql_conn.cursor()

    # Спочатку знайдемо UUID Бюджет 000002597 у форматі який використовується у Fact_Cashflow
    uuid_b2597 = S(ref_b2597.УникальныйИдентификатор()).replace("-", "")
    uuid_globyno = S(ref_globyno.УникальныйИдентификатор()).replace("-", "")
    uuid_st = S(ref_st.УникальныйИдентификатор()).replace("-", "")

    cur.execute("""
        SELECT TOP 50
            f.Source_Recorder_Type,
            f.Source_Recorder_Id,
            f.Period_Month,
            f.Source,
            f.Sum_Plan,
            f.Sum_Plan_Object,
            f.Sum_Kazna,
            f.Sum_Fact
        FROM Fact_Cashflow f
        WHERE UPPER(f.Source_Recorder_Id) = UPPER(?)
          AND UPPER(f.Department_Id) = UPPER(?)
          AND UPPER(f.DDS_Article_Id) = UPPER(?)
        ORDER BY f.Period_Month
    """, uuid_b2597, uuid_globyno, uuid_st)

    rows = cur.fetchall()
    step3_storno = 0.0
    step3_normal = 0.0
    if not rows:
        print("  (немає записів у Fact_Cashflow)")
        step3_ok = False
    else:
        print(f"  {'Период_Месяц':<12} | {'Source':<15} | {'Sum_Plan':>15} | {'Sum_Plan_Object':>15} | {'Sum_Kazna':>15} | {'Sum_Fact':>15}")
        print("-" * 100)
        for r in rows:
            pm = r.Period_Month.strftime("%d.%m.%Y") if r.Period_Month else ""
            so = float(r.Sum_Plan_Object or 0)
            if so < 0:
                step3_storno += so
            else:
                step3_normal += so
            print(f"  {pm:<12} | {(r.Source or ''):<15} | {fmt(r.Sum_Plan or 0, 15)} | {fmt(r.Sum_Plan_Object or 0, 15)} | {fmt(r.Sum_Kazna or 0, 15)} | {fmt(r.Sum_Fact or 0, 15)}")
        step3_ok = step3_storno < -0.01

    print(f"\nΣ позитив = {fmt(step3_normal)}, Σ сторно = {fmt(step3_storno)}")
    sql_conn.close()
except Exception as e:
    print(f"FAIL: {e}")
    step3_ok = False

print(f"\nРезультат STEP 3: {'✓ Сторно є у SQL Fact_Cashflow (ETL прогнаний)' if step3_ok else '✗ Сторно ВІДСУТНЯ у SQL — потрібно прогнати ETL `python _Rarzrabotki/Olap/Ai_Olap/main.py`'}")


# === ВИСНОВОК ===
banner("ДІАГНОЗ ЛАНЦЮГА")

print(f"  STEP 1 — 1C регістр А_БюджетыНаМесяц (raw):  {'✓' if step1_ok else '✗'}")
print(f"  STEP 2 — 1C регістр А_ОтчетDDS_Свод (OLAP): {'✓' if step2_ok else '✗'}")
print(f"  STEP 3 — SQL OlapBASERP.Fact_Cashflow (ETL): {'✓' if step3_ok else '✗'}")

print()
if step1_ok and not step2_ok:
    print("→ РОЗРИВ між 1С raw register і OLAP register.")
    print("  Дії: 1) Конфігуратор → Поддержка → Загрузить конфігурацію з файлів")
    print("       2) F7 (Обновить конфигурацию БД)")
    print("       3) Знайти OLAP-документ за Жовтень 2024 → Перепровести (Ctrl+Shift+S)")
elif step2_ok and not step3_ok:
    print("→ РОЗРИВ між 1C OLAP register і SQL OlapBASERP.")
    print("  Дії: cd C:\\Configuration_downloads\\BASERP25\\_Rarzrabotki\\Olap\\Ai_Olap")
    print("       python main.py")
elif step1_ok and step2_ok and step3_ok:
    print("→ ВСЯ цепь 1С → SQL працює. Залишилось:")
    print("  Дії: PowerBI Desktop → Главная → Обновить (refresh) → зберегти PBIX (Ctrl+S)")
else:
    print("→ ВСІ ланки порушені або діагностика виявила несподіваний стан.")

print("\nDONE.")
