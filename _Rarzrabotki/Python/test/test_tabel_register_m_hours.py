# -*- coding: utf-8 -*-
"""Приёмка ресурса ЧасыОфициальные регистра ДанныеТабелирования (kazna).

Перепроводит РОВНО ОДИН согласованный документ: Табель 000003260 от 01.08.2026 (ІНДЕПТ СТІЛ,
скриншот пользователя). Ничего не удаляет. Остальную историю перепроводит пользователь.

Проверки:
  1. До перепроведения снимаем Обороты() по Состоянию "Р" документа (ЧасыФакт).
  2. Перепроводим документ через COM.
  3. После: у Чешко С.А. и Нерезенко Є.В. ровно 3 строки Состояние "М" (11, 12, 13.08.2026,
     период 23:59:59), ЧасыОфициальные = 8, ЧасыФакт = 0.
  4. "Р"-обороты документа равны ДО; строк Состояние "?" нет.
Печать PASS/FAIL, exit code 1 при провале.
"""
import sys
import datetime
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'
DOC_NUMBER = "000003260"
EXPECT_M = {"Чешко Сергій Анатолійович": [11, 12, 13], "Нерезенко Євгеній Віктоорович": [11, 12, 13]}

fails = 0


def check(cond, msg):
    global fails
    print(("PASS " if cond else "FAIL ") + msg)
    if not cond:
        fails += 1


def rows(t):
    cols = [t.Колонки.Получить(i).Имя for i in range(t.Колонки.Количество())]
    return [{c: getattr(t.Получить(i), c) for c in cols} for i in range(t.Количество())]


def run(base, txt, params=None):
    q = base.NewObject("Запрос")
    q.Text = txt
    for k, v in (params or {}).items():
        q.SetParameter(k, v)
    return rows(q.Execute().Выгрузить())


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kz = v8.Connect(CONN)
    S = kz.String

    docs = run(kz, """ВЫБРАТЬ Д.Ссылка КАК Ссылка, Д.Проведен КАК Проведен, Д.Организация.Наименование КАК Орг
        ИЗ Документ.ТабельУчетаРабочегоВремени КАК Д ГДЕ Д.Номер = &Номер И НЕ Д.ПометкаУдаления""", {"Номер": DOC_NUMBER})
    check(len(docs) == 1, f"документ {DOC_NUMBER} найден один: {len(docs)}")
    if len(docs) != 1:
        return 1
    ref = docs[0]["Ссылка"]
    print("   документ:", S(ref), "| орг:", S(docs[0]["Орг"]), "| проведён:", docs[0]["Проведен"])

    q_r = """ВЫБРАТЬ Т.Сотрудник.Наименование КАК ФИО, СУММА(Т.ЧасыФакт) КАК Факт
        ИЗ РегистрНакопления.ДанныеТабелирования КАК Т
        ГДЕ Т.Регистратор = &Док И Т.Состояние = "Р" И Т.Активность
        СГРУППИРОВАТЬ ПО Т.Сотрудник.Наименование"""
    before = {S(r["ФИО"]): float(r["Факт"]) for r in run(kz, q_r, {"Док": ref})}
    print(f"   ДО: строк Р-оборотов {len(before)}, Σ ЧасыФакт = {sum(before.values())}")

    # перепроведение
    obj = ref.ПолучитьОбъект()
    obj.Записать(kz.РежимЗаписиДокумента.Проведение)
    print("   документ перепроведён")

    after = {S(r["ФИО"]): float(r["Факт"]) for r in run(kz, q_r, {"Док": ref})}
    nz_before = {k: v for k, v in before.items() if v != 0}
    nz_after = {k: v for k, v in after.items() if v != 0}
    check(nz_before == nz_after, f"Р-часы документа (ненулевые) не изменились: Σ {sum(after.values())}, строк {len(nz_after)}")

    m_rows = run(kz, """ВЫБРАТЬ Т.Сотрудник.Наименование КАК ФИО, Т.Период КАК Период, Т.Состояние КАК Состояние,
            Т.ЧасыФакт КАК ЧасыФакт, Т.ЧасыОфициальные КАК ЧасыОфициальные
        ИЗ РегистрНакопления.ДанныеТабелирования КАК Т
        ГДЕ Т.Регистратор = &Док И Т.Состояние В ("М", "?") И Т.Активность
        УПОРЯДОЧИТЬ ПО ФИО, Период""", {"Док": ref})
    q_rows = [r for r in m_rows if S(r["Состояние"]) == "?"]
    check(len(q_rows) == 0, f"строк Состояние '?' нет: {len(q_rows)}")
    for fio, days in EXPECT_M.items():
        mine = [r for r in m_rows if S(r["ФИО"]) == fio and S(r["Состояние"]) == "М"]
        got_days = sorted(int(S(r["Период"])[:2]) for r in mine)
        check(got_days == days, f"{fio}: дни М {got_days} == {days}")
        check(all(float(r["ЧасыОфициальные"]) == 8 for r in mine) and len(mine) > 0, f"{fio}: ЧасыОфициальные = 8 во всех строках М")
        check(all(float(r["ЧасыФакт"]) == 0 for r in mine), f"{fio}: ЧасыФакт = 0 во всех строках М")
        for r in mine:
            print(f"      {S(r['Период'])} | М | факт {float(r['ЧасыФакт'])} | офиц {float(r['ЧасыОфициальные'])}")
        check(all(S(r["Период"]).endswith("23:59:59") for r in mine), f"{fio}: период строк М на 23:59:59")

    # ожидание из самой ТЧ документа: все ячейки вида "NМ" и "М"
    exp_hours, exp_cells, exp_people = 0.0, 0, set()
    tch = obj.ОтработанноеВремя
    for i in range(tch.Количество()):
        row = tch.Получить(i)
        for d in range(1, 32):
            v = str(getattr(row, "ПервыйЧасов%d" % d) or "").strip().upper().replace("M", "М")
            if v.endswith("М"):
                exp_cells += 1
                exp_people.add(S(row.СотрудникСпр))
                num = v[:-1].replace(",", ".")
                exp_hours += float(num) if num else 0.0
    tot = run(kz, """ВЫБРАТЬ СУММА(Т.ЧасыОфициальные) КАК Офиц, КОЛИЧЕСТВО(*) КАК Строк, КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Т.Сотрудник) КАК Сотр
        ИЗ РегистрНакопления.ДанныеТабелирования КАК Т ГДЕ Т.Регистратор = &Док И Т.Состояние = "М" И Т.Активность""", {"Док": ref})
    got_h, got_rows, got_people = float(tot[0]["Офиц"]), int(tot[0]["Строк"]), int(tot[0]["Сотр"])
    print(f"   ТЧ: ячеек с М {exp_cells}, людей {len(exp_people)} {sorted(exp_people)}, Σ часов {exp_hours}")
    print(f"   регистр: строк М {got_rows}, людей {got_people}, Σ ЧасыОфициальные {got_h}")
    check(got_rows == exp_cells and got_people == len(exp_people) and got_h == exp_hours,
          "строки М в регистре = ячейки М в ТЧ документа (кол-во, люди, Σ часов)")

    print(f"\nИтого провалов: {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
