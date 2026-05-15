# -*- coding: utf-8 -*-
"""Acceptance Fact_Balance: Σ по PAP_Article == ПАП BaseERP (січень/ТОВ) до
копійки; Актив=Пасив (Σ Sum_Close ≈ 0). Свёртка в 1С (Етапи 1-4) — ETL лише
копіює регістр. Дзеркало test_balans_s4_verify.py (еталон ПАП ОстаткиИОбороты)."""
import pyodbc
import pythoncom
import win32com.client

OLAP = ("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
        "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;")
ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
TOL = 0.01      # копійка
TOL_OT = 1.0    # «Оплата труда» — округлення розкладу по фізлицю


def test_balance_sum_matches_pap_and_zero():
    cx = pyodbc.connect(OLAP)
    rows = cx.execute("""SELECT d.PAP_Article_Name, SUM(f.Sum_Close)
        FROM Fact_Balance f JOIN Dim_PAP_Articles d
          ON f.PAP_Article_ID = d.PAP_Article_ID
        WHERE f.Period_Month = '2026-01-01'
        GROUP BY d.PAP_Article_Name""").fetchall()
    olap = {r[0]: float(r[1] or 0) for r in rows}
    total = sum(olap.values())
    cx.close()
    assert abs(total) < 1.0, f"Актив!=Пасив: Σ Sum_Close={total:,.2f}"

    pythoncom.CoInitialize()
    conn = win32com.client.Dispatch("V83.COMConnector").Connect(ERP)
    q0 = conn.NewObject("Запрос")
    q0.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Справочник.Организации '
                'ГДЕ КодПоЕДРПОУ = "40645273"')
    s = q0.Выполнить().Выбрать()
    s.Следующий()
    org = s.С
    qb = conn.NewObject("Запрос")
    qb.Текст = ("ВЫБРАТЬ НАЧАЛОПЕРИОДА(ДАТАВРЕМЯ(2026,1,1),МЕСЯЦ) КАК НМ, "
                "КОНЕЦПЕРИОДА(ДАТАВРЕМЯ(2026,1,31,23,59,59),ДЕНЬ) КАК КД")
    rb = qb.Выполнить().Выбрать()
    rb.Следующий()
    qi = conn.NewObject("Запрос")
    qi.Текст = ('ВЫБРАТЬ Ссылка КАК С ИЗ '
                'ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование В (&Сп)')
    a = conn.NewObject("Массив")
    for n in ["Собственные средства", "Доходы текущего периода",
              "Расходы текущего периода"]:
        a.Добавить(n)
    qi.УстановитьПараметр("Сп", a)
    искл = qi.Выполнить().Выгрузить().ВыгрузитьКолонку("С")
    qp = conn.NewObject("Запрос")
    qp.Текст = (
        "ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Б.Статья) КАК Н, "
        "СУММА(Б.СуммаКонечныйОстаток) КАК sK "
        "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.ОстаткиИОбороты(&Д1,&Д2,Авто,,"
        "Организация=&Орг И НЕ Статья В ИЕРАРХИИ(&Искл)) КАК Б "
        "СГРУППИРОВАТЬ ПО ПРЕДСТАВЛЕНИЕ(Б.Статья) "
        "ИМЕЮЩИЕ СУММА(Б.СуммаКонечныйОстаток) <> 0")
    qp.УстановитьПараметр("Д1", rb.НМ)
    qp.УстановитьПараметр("Д2", rb.КД)
    qp.УстановитьПараметр("Орг", org)
    qp.УстановитьПараметр("Искл", искл)
    t = qp.Выполнить().Выгрузить()
    pap = {str(t.Получить(i).Н): float(t.Получить(i).sK or 0)
           for i in range(t.Количество())}

    bad = []
    for k in set(pap) | set(olap):
        d = abs(olap.get(k, 0) - pap.get(k, 0))
        lim = TOL_OT if "плата труда" in k.lower() else TOL
        if d > lim:
            bad.append((k, pap.get(k, 0), olap.get(k, 0), d))
    assert not bad, ("Fact_Balance != ПАП по статтях: "
                     + "; ".join(f"{k[:32]} ПАП={p:,.2f} OLAP={o:,.2f} Δ={dl:,.4f}"
                                 for k, p, o, dl in bad[:8]))
