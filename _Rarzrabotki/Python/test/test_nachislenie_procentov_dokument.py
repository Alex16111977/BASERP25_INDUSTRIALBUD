# -*- coding: utf-8 -*-
"""Нарахування відсотків: заповнення документа і захист від дублів (план 2026-10-16, Task 4).

Документи створюються В ПАМ'ЯТІ й не записуються. Код виходу 0 = OK.
"""
import datetime as dt
import sys

import win32com.client

sys.stdout.reconfigure(encoding='utf-8')
erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
ok = True
ORG = erp.Справочники.Организации.НайтиПоНаименованию('ТОВ "ІНДАСТРІАЛБУД"', True)
СП = erp.Перечисления.А_СпособыРасчетаПроцентов
PREF = "%% за использование от кредитора -- "


def check(c, m):
    global ok
    print(("OK  " if c else "FAIL") + " " + m)
    ok = ok and bool(c)


def doc(y, m, sposob):
    d = erp.Документы.А_НачислениеПроцентовПоКредитам.СоздатьДокумент()
    d.Организация = ORG
    d.МесяцНачисления = dt.datetime(y, m, 1) + dt.timedelta(hours=3)  # COM −3 год
    d.СпособЗаполнения = sposob
    return d


def rows(d):
    return {S(d.Начисления.Получить(i).Договор).replace(PREF, ""): d.Начисления.Получить(i) for i in range(d.Начисления.Количество())}


# вересень, за виплатою
d = doc(2026, 9, СП.ПоВыплате)
d.ЗаполнитьНачисления()
r = rows(d)
total = round(sum(float(x.СуммаКНачислению) for x in r.values()), 2)
kr, kam = r.get("КРИВОЙ РОГ ЦЕМЕНТ-2"), r.get("Каменец Цемент")
check(len(r) == 5 and total == 102780 and kr is not None and float(kr.СуммаВыплат) == 7550
      and abs(float(kr.СуммаПоФормуле) - 7549.67) <= 0.0101 and float(kam.СуммаВыплат) == 9070
      and abs(float(kam.СуммаПоФормуле) - 9070.36) <= 0.0101,
      f"test_zapolnit_sentyabr_po_vyplate: рядків {len(r)}, Σ {total}, КР {kr and (S(kr.СуммаВыплат), S(kr.СуммаПоФормуле))}")

# червень: Подгорцы без тіла, ПРООН — підрозділ з договору
d6 = doc(2026, 6, СП.ПоВыплате)
d6.ЗаполнитьНачисления()
r6 = rows(d6)
pod = r6.get('КГ "Подгорцы"')
check(pod is not None and float(pod.СуммаПоФормуле) == 0 and float(pod.СуммаВыплат) == 197500
      and float(pod.СуммаКНачислению) == 197500 and S(pod.Подразделение) == 'КГ "Подгорцы"',
      f"test_zapolnit_iyun_podgorcy: {pod and (S(pod.СуммаПоФормуле), S(pod.СуммаВыплат), S(pod.Подразделение))}")
proon = r6.get("МД ПРООН Черкаси  ДСНС")
check(proon is not None and S(proon.Подразделение) == "МД ПРООН Черкаси  ДСНС" and float(proon.СуммаКНачислению) == 295,
      f"test_zapolnit_iyun_podrazdelenie_iz_dogovora: {proon and S(proon.Подразделение)}")
check(round(sum(float(x.СуммаКНачислению) for x in r6.values()), 2) == 314820, "червень Σ = 314 820")

# ручна сума зберігається
kr = rows(d)["КРИВОЙ РОГ ЦЕМЕНТ-2"]
kod = int(kr.КодСтроки)
kr.СпособРасчета = СП.Вручную
kr.СуммаКНачислению = 1000
d.ЗаполнитьНачисления()
kr2 = rows(d)["КРИВОЙ РОГ ЦЕМЕНТ-2"]
check(float(kr2.СуммаКНачислению) == 1000 and S(kr2.СпособРасчета) == S(СП.Вручную) and int(kr2.КодСтроки) == kod,
      f"test_ruchnaya_summa_sohranyaetsya: {S(kr2.СуммаКНачислению)} / {S(kr2.СпособРасчета)} / код {S(kr2.КодСтроки)} (був {kod})")

# липень за формулою
d7 = doc(2026, 7, СП.ПоФормуле)
d7.ЗаполнитьНачисления()
kr7 = rows(d7).get("КРИВОЙ РОГ ЦЕМЕНТ-2")
intervaly = [(S(d7.Расшифровка.Получить(i).ДатаНачала)[:10], S(d7.Расшифровка.Получить(i).ДатаОкончания)[:10])
             for i in range(d7.Расшифровка.Количество())
             if S(d7.Расшифровка.Получить(i).Договор).endswith("КРИВОЙ РОГ ЦЕМЕНТ-2")
             and S(d7.Расшифровка.Получить(i).ВидСтроки) == "Інтервал"]
check(kr7 is not None and abs(float(kr7.СуммаКНачислению) - 4222.93) <= 0.0101
      and intervaly == [("22.07.2026", "27.07.2026"), ("28.07.2026", "29.07.2026"), ("30.07.2026", "31.07.2026")],
      f"test_po_formule_iyul_kr: {kr7 and S(kr7.СуммаКНачислению)}, інтервали {intervaly}")

# дублі в одній ТЧ
check(not d.ЕстьДублиДоговоров(), "test_dubl_v_odnoy_tch: без копії → Ложь")
nov = d.Начисления.Добавить()
nov.Договор = rows(d)["КРИВОЙ РОГ ЦЕМЕНТ-2"].Договор
check(bool(d.ЕстьДублиДоговоров()), "test_dubl_v_odnoy_tch: з копією → Истина")

sys.exit(0 if ok else 1)
