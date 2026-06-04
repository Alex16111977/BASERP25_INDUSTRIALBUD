# -*- coding: utf-8 -*-
"""Diagnostic v2 — correct field names."""
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


v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Lookup
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_БюджетМесяц ГДЕ Номер = &Н"
q.SetParameter("Н", "000002597")
ref_b2597 = q.Execute().Выгрузить().Получить(0).Ссылка
uuid_b2597 = S(ref_b2597.УникальныйИдентификатор()).replace("-", "").upper()
print(f"Бюджет 000002597 UUID: {uuid_b2597}")


# === STEP 1: А_БюджетыНаМесяц — RECAP ===
banner("STEP 1: А_БюджетыНаМесяц — Бюджет 000002597 / Глобино-2 / Сума<0 / Месяц=Окт2024")

q1 = erp.NewObject("Запрос")
q1.Text = """
ВЫБРАТЬ
    Р.Подразделение.Наименование             КАК Подр,
    Р.СтатьяДвиженияДенежныхСредств.Наименование КАК Стат,
    Р.Месяц, Р.ВидПериода, Р.Сумма
ИЗ РегистрНакопления.А_БюджетыНаМесяц КАК Р
ГДЕ Р.Регистратор = &Рег
    И Р.Подразделение.Наименование = "Глобино-2"
    И Р.Месяц = ДАТАВРЕМЯ(2024, 10, 1)
    И Р.Сумма < 0
УПОРЯДОЧИТЬ ПО Р.Сумма
"""
q1.SetParameter("Рег", ref_b2597)
res1 = q1.Execute().Выгрузить()
print(f"Знайдено сторно-рядків: {res1.Количество()}")
total1 = 0
for i in range(res1.Количество()):
    r = res1.Получить(i)
    total1 += float(r.Сумма)
    print(f"  Стат={S(r.Стат)[:35]:<35} | ВидПер={S(r.ВидПериода)} | Сума={fmt(r.Сумма)}")
print(f"Σ сторно = {fmt(total1)}")
step1_ok = total1 < -0.01


# === STEP 2: А_ОтчетDDS_Свод — для Бюджет 000002597 ===
banner("STEP 2: А_ОтчетDDS_Свод — записи от Бюджет 000002597 / Глобино-2")

q2 = erp.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ
    Р.Регистратор             КАК ОЛАП,
    Р.Регистратор.Месяц       КАК ОЛАПМесяц,
    Р.Подразделение.Наименование КАК Подр,
    Р.СтатьяДвиженияДенежныхСредств.Наименование КАК Стат,
    Р.Источник                КАК Источник,
    Р.СуммаПлан               КАК СуммаПлан,
    Р.СуммаПланОбъект         КАК СуммаПланОбъект,
    Р.Сумма                   КАК Сумма
ИЗ РегистрСведений.А_ОтчетDDS_Свод КАК Р
ГДЕ Р.ДокументДвижения = &Бюджет
    И Р.Подразделение.Наименование = "Глобино-2"
"""
q2.SetParameter("Бюджет", ref_b2597)
try:
    res2 = q2.Execute().Выгрузить()
    print(f"Знайдено: {res2.Количество()}")
    total2_storno = 0
    for i in range(res2.Количество()):
        r = res2.Получить(i)
        plan_o = float(r.СуммаПланОбъект or 0)
        if plan_o < 0:
            total2_storno += plan_o
        print(f"  ОЛАП={S(r.ОЛАП).split(',')[0][:40]:<40} | Мес={r.ОЛАПМесяц.strftime('%d.%m.%Y') if r.ОЛАПМесяц else ''} | Стат={S(r.Стат)[:30]:<30} | СуммаПлан={fmt(r.СуммаПлан)} | СуммаПланОбъект={fmt(r.СуммаПланОбъект)} | Сумма={fmt(r.Сумма)}")
    step2_ok = total2_storno < -0.01
    print(f"Σ сторно у А_ОтчетDDS_Свод = {fmt(total2_storno)}")
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
    print(f"FAIL: {msg}")
    step2_ok = False


# === STEP 2b: OLAP-документи для Окт 2024 ===
banner("STEP 2b: OLAP-документи з Месяц=Окт2024")

q2b = erp.NewObject("Запрос")
q2b.Text = """
ВЫБРАТЬ
    Д.Номер, Д.Дата, Д.Месяц, Д.Проведен, Д.ПометкаУдаления
ИЗ Документ.А_ФинРез_DDS КАК Д
ГДЕ Д.Месяц = ДАТАВРЕМЯ(2024, 10, 1)
"""
res2b = q2b.Execute().Выгрузить()
print(f"Знайдено OLAP-документів: {res2b.Количество()}")
for i in range(res2b.Количество()):
    r = res2b.Получить(i)
    pr = "✓ проведен" if r.Проведен else "✗ НЕ проведен"
    if r.ПометкаУдаления:
        pr = "✗ помічений на видалення"
    print(f"  №{r.Номер} от {r.Дата.strftime('%d.%m.%Y %H:%M')} | Месяц={r.Месяц.strftime('%d.%m.%Y')} | {pr}")


# === STEP 3: SQL OlapBASERP.Fact_Cashflow ===
banner("STEP 3: SQL Fact_Cashflow — Document_ID = Бюджет 000002597")

try:
    sql_conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=localhost;DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;"
        "TrustServerCertificate=yes;"
    )
    cur = sql_conn.cursor()
    cur.execute("""
        SELECT
            f.Period_Month, f.Source,
            d.Department_Name,
            a.DDS_Article_Name,
            f.Sum_Plan, f.Sum_Plan_Object, f.Sum_Kazna, f.Sum_Fact
        FROM Fact_Cashflow f
        LEFT JOIN Dim_Departments d ON UPPER(d.Department_ID) = UPPER(f.Department_ID)
        LEFT JOIN Dim_DDS_Articles a ON UPPER(a.DDS_Article_ID) = UPPER(f.DDS_Article_ID)
        WHERE UPPER(f.Document_ID) = UPPER(?)
        ORDER BY f.Period_Month, a.DDS_Article_Name
    """, uuid_b2597)
    rows = cur.fetchall()
    print(f"Знайдено: {len(rows)}")
    total3_storno = 0
    for r in rows[:30]:
        pm = r.Period_Month.strftime("%d.%m.%Y") if r.Period_Month else ""
        po = float(r.Sum_Plan_Object or 0)
        if po < 0:
            total3_storno += po
        print(f"  Period_Mes={pm:<11} | Source={(r.Source or '')[:18]:<18} | Подр={(r.Department_Name or '')[:15]:<15} | Стат={(r.DDS_Article_Name or '')[:30]:<30} | Plan={fmt(r.Sum_Plan or 0)} | PlanObj={fmt(r.Sum_Plan_Object or 0)}")
    step3_ok = total3_storno < -0.01
    print(f"Σ сторно у SQL Fact_Cashflow = {fmt(total3_storno)}")
    sql_conn.close()
except Exception as e:
    print(f"FAIL: {e}")
    step3_ok = False


# === ВИСНОВОК ===
banner("ДІАГНОЗ ЛАНЦЮГА")

print(f"  STEP 1 — 1C регістр А_БюджетыНаМесяц (raw):  {'✓' if step1_ok else '✗'}")
print(f"  STEP 2 — 1C регістр А_ОтчетDDS_Свод (OLAP): {'✓' if step2_ok else '✗'}")
print(f"  STEP 3 — SQL OlapBASERP.Fact_Cashflow (ETL): {'✓' if step3_ok else '✗'}")

print()
if step1_ok and not step2_ok:
    print("→ РОЗРИВ між 1С raw register і OLAP register.")
    print("  Дії в 1С:Конфігураторі:")
    print("  1) Конфігурація → Поддержка → Загрузить конфігурацію з файлів")
    print("     (вибрати Documents/А_ФинРез_DDS/Ext/ObjectModule.bsl)")
    print("  2) F7 (Обновить конфигурацию БД) — підтвердити прийняття змін")
    print("  3) Открити OLAP-документ з Месяц=01.10.2024 → Перепровести (Ctrl+Shift+S)")
    print("  4) Якщо такого документа немає — створити новий за Жовтень 2024")
elif step2_ok and not step3_ok:
    print("→ РОЗРИВ між 1C OLAP register і SQL OlapBASERP.")
    print("  Дії: cd C:\\Configuration_downloads\\BASERP25\\_Rarzrabotki\\Olap\\Ai_Olap")
    print("       .venv\\Scripts\\python main.py")
elif step1_ok and step2_ok and step3_ok:
    print("→ ВСЯ цепь 1С → SQL працює. Залишилось:")
    print("  Дії: PowerBI Desktop → Главная → Обновить (refresh) → зберегти PBIX (Ctrl+S)")
else:
    print("→ STEP 1 теж зламаний — критичне розслідування.")

print("\nDONE.")
