# -*- coding: utf-8 -*-
"""READ-ONLY верификация Dim_VidyKontragentov + 2 новых FK-колонок Dim_Contracts
(VidKontragenta_ID -> Dim_VidyKontragentov, NapravlenieUslug_ID -> Dim_Directions).

Гейты:
1. Dim_VidyKontragentov = 5 видов + "(Пусто)"; наименования из 1С.
2. Dim_Contracts: раскладка договоров по видам == 1С (COM, тот же момент); FK orphans 0.
3. Dim_Contracts: NapravlenieUslug_ID count == 1С; FK -> Dim_Directions orphans 0.
4. Кросс-сверка баланса: Fact_Balance 2026-01 Σ Sum_Close по виду "Внутригрупповые"
   через JOIN Dim_Contracts == РС А_ОтчетБаланс_Свод (COM, Договор.А_ВидКонтрагента).
5. Регресс: Fact_Balance 2026-01 Σ Sum_Close (все Source) == 0,00 (полный баланс).
"""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql
import win32com.client

fail = []

# ---------- 1С (COM) ----------
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ Д.А_ВидКонтрагента.Наименование КАК Вид, КОЛИЧЕСТВО(*) КАК Кол
ИЗ Справочник.ДоговорыКонтрагентов КАК Д
ГДЕ НЕ Д.ПометкаУдаления
    И Д.А_ВидКонтрагента <> ЗНАЧЕНИЕ(Справочник.А_ВидыКонтрагентовДляБаланса.ПустаяСсылка)
СГРУППИРОВАТЬ ПО Д.А_ВидКонтрагента.Наименование
"""
r = q.Execute().Выгрузить()
com_vid_counts = {}
for i in range(r.Количество()):
    row = r.Получить(i)
    com_vid_counts[S(row.Вид)] = int(row.Кол)

q2 = erp.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол
ИЗ Справочник.ДоговорыКонтрагентов КАК Д
ГДЕ НЕ Д.ПометкаУдаления
    И Д.А_НаправлениеОказаниеУслуг <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка)
"""
com_napr_count = int(q2.Execute().Выгрузить().Получить(0).Кол)

ВидВнутригрупповые = erp.Справочники.А_ВидыКонтрагентовДляБаланса.Внутригрупповые
q3 = erp.NewObject("Запрос")
q3.Text = """
ВЫБРАТЬ СУММА(Т.СуммаКонечныйОстаток) КАК КО
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Т
ГДЕ Т.Регистратор.Месяц МЕЖДУ ДАТАВРЕМЯ(2026,1,1) И ДАТАВРЕМЯ(2026,1,31,23,59,59)
    И Т.Договор.А_ВидКонтрагента = &Вид
"""
q3.SetParameter("Вид", ВидВнутригрупповые)
r3 = q3.Execute().Выгрузить()
com_vg_ko = float(r3.Получить(0).КО or 0)

print("1С (COM):")
for k, v in sorted(com_vid_counts.items()):
    print(f"  договоров [{k}]: {v}")
print(f"  договоров с НаправлениеОказаниеУслуг: {com_napr_count}")
print(f"  РС свод 2026-01 Σ КО (Внутригрупповые): {com_vg_ko:,.2f}")

# ---------- SQL (OlapBASERP) ----------
with get_olap_sql() as c:
    cur = c.cursor()

    # 1. Dim_VidyKontragentov
    rows = cur.execute(
        "SELECT VidKontragenta_Name, Code FROM Dim_VidyKontragentov ORDER BY Code").fetchall()
    names = [r[0] for r in rows]
    print(f"\nDim_VidyKontragentov ({len(rows)}): {names}")
    if len(rows) != 6:
        fail.append(f"Dim_VidyKontragentov rows={len(rows)} (ожид 6 = 5 видов + (Пусто))")
    for need in ("Внутригрупповые", "Внутренние подразделения", "Собственники",
                 "Внешние", "Кредиторы", "(Пусто)"):
        if need not in names:
            fail.append(f"в Dim_VidyKontragentov нет «{need}»")

    # 2. раскладка по видам + FK orphans
    rows = cur.execute("""
        SELECT v.VidKontragenta_Name, COUNT(*) AS cnt
        FROM Dim_Contracts d
        JOIN Dim_VidyKontragentov v ON v.VidKontragenta_ID = d.VidKontragenta_ID
        WHERE d.Marked_For_Deletion = 0
        GROUP BY v.VidKontragenta_Name""").fetchall()
    sql_vid_counts = {r[0]: r[1] for r in rows}
    print(f"SQL раскладка по видам: {sql_vid_counts}")
    for vid, n_com in com_vid_counts.items():
        n_sql = sql_vid_counts.get(vid, 0)
        if n_sql != n_com:
            fail.append(f"вид «{vid}»: SQL={n_sql} != 1С={n_com}")

    orph_v = cur.execute("""
        SELECT COUNT(*) FROM Dim_Contracts d
        WHERE d.VidKontragenta_ID IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM Dim_VidyKontragentov v
                          WHERE v.VidKontragenta_ID = d.VidKontragenta_ID)""").fetchval()
    print(f"FK orphans VidKontragenta_ID: {orph_v} (ожид 0)")
    if orph_v:
        fail.append(f"{orph_v} Dim_Contracts.VidKontragenta_ID без Dim_VidyKontragentov")

    # 3. НаправлениеУслуг
    n_sql_napr = cur.execute("""
        SELECT COUNT(*) FROM Dim_Contracts
        WHERE NapravlenieUslug_ID IS NOT NULL AND Marked_For_Deletion = 0""").fetchval()
    print(f"SQL NapravlenieUslug_ID NOT NULL: {n_sql_napr} (1С: {com_napr_count})")
    if n_sql_napr != com_napr_count:
        fail.append(f"NapravlenieUslug: SQL={n_sql_napr} != 1С={com_napr_count}")

    orph_n = cur.execute("""
        SELECT COUNT(*) FROM Dim_Contracts d
        WHERE d.NapravlenieUslug_ID IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM Dim_Directions r
                          WHERE r.Direction_ID = d.NapravlenieUslug_ID)""").fetchval()
    print(f"FK orphans NapravlenieUslug_ID -> Dim_Directions: {orph_n} (ожид 0)")
    if orph_n:
        fail.append(f"{orph_n} Dim_Contracts.NapravlenieUslug_ID без Dim_Directions")

    # имя направления (денорм-колонка для PBIX, v2: вместо отдельного измерения)
    n_name = cur.execute("""
        SELECT COUNT(*) FROM Dim_Contracts
        WHERE NapravlenieUslug_Name IS NOT NULL AND Marked_For_Deletion = 0""").fetchval()
    print(f"SQL NapravlenieUslug_Name NOT NULL: {n_name} (1С: {com_napr_count})")
    if n_name != com_napr_count:
        fail.append(f"NapravlenieUslug_Name: SQL={n_name} != 1С={com_napr_count}")

    # 4. кросс-сверка баланса по виду «Внутригрупповые», 2026-01
    sql_vg_ko = cur.execute("""
        SELECT COALESCE(SUM(f.Sum_Close), 0)
        FROM Fact_Balance f
        JOIN Dim_Contracts d ON d.Contract_ID = f.Contract_ID
        JOIN Dim_VidyKontragentov v ON v.VidKontragenta_ID = d.VidKontragenta_ID
        WHERE f.Period_Month = '2026-01-01'
          AND v.VidKontragenta_Name = N'Внутригрупповые'""").fetchval()
    sql_vg_ko = float(sql_vg_ko or 0)
    print(f"Fact_Balance 2026-01 Σ Sum_Close (Внутригрупповые): {sql_vg_ko:,.2f}")
    if abs(sql_vg_ko - com_vg_ko) > 0.01:
        fail.append(f"кросс-сверка Внутригрупповые: SQL={sql_vg_ko:,.2f} != РС={com_vg_ko:,.2f}")

    # 5. регресс полного баланса 2026-01
    total = float(cur.execute("""
        SELECT COALESCE(SUM(Sum_Close), 0) FROM Fact_Balance
        WHERE Period_Month = '2026-01-01'""").fetchval() or 0)
    aktiv = float(cur.execute("""
        SELECT COALESCE(SUM(Sum_Close), 0) FROM Fact_Balance
        WHERE Period_Month = '2026-01-01' AND TipPokazatelya = N'Актив'""").fetchval() or 0)
    print(f"Fact_Balance 2026-01: Σ Sum_Close (все)={total:,.2f} (ожид 0,00); Актив={aktiv:,.2f}")
    if abs(total) > 0.01:
        fail.append(f"полный баланс 2026-01 нарушен: Σ Sum_Close={total:,.2f}")

print("\n" + "=" * 64)
if fail:
    for f in fail:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: Dim_VidyKontragentov + VidKontragenta_ID/NapravlenieUslug_ID в Dim_Contracts; "
      "раскладка == 1С; FK 0 orphans; кросс-сверка Внутригрупповые до копейки; "
      "полный баланс 2026-01 не регрессировал")
