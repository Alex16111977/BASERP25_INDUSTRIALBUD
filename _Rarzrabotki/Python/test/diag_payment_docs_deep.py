# -*- coding: utf-8 -*-
# Глубокий разбор платёжных документов с Партнер=Неизв: все заполненные поля шапки + строки,
# чтобы понять, можно ли вывести партнёра (через Получатель/ПодотчЛицо/Основание/банк-счёт и т.п.).
import sys
sys.path.insert(0, r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test')
import _common_neizv as c

erp = c.connect()
P = c.neizv(erp)


def dump_doc(obj, тч_имена):
    md = obj.Метаданные()
    print("  --- шапка (заполненные ссылочные/важные поля) ---")
    for r in md.Реквизиты:
        имя = r.Имя
        try:
            v = obj[имя]
        except Exception:
            continue
        if not erp.ЗначениеЗаполнено(v):
            continue
        # печатаем только не-примитивные (ссылки/перечисления) и строки покороче
        try:
            tn = erp.XMLСтрока(erp.ТипЗнч(v))
        except Exception:
            tn = "?"
        if "Ref" in str(tn) or "Enum" in str(tn) or имя in ("НазначениеПлатежа",):
            print(f"     {имя} = {v}")
    for тч in тч_имена:
        try:
            строки = obj[тч]
        except Exception:
            continue
        print(f"  --- ТЧ {тч} ---")
        i = 0
        for row in строки:
            i += 1
            if i > 4:
                print("     ...")
                break
            partner = ""
            try:
                partner = "Неизв" if c.same(erp, row.Партнер, P) else str(row.Партнер)
            except Exception:
                partner = "(нет Партнер)"
            орс = ""
            try:
                орс = str(row.ОбъектРасчетов) if erp.ЗначениеЗаполнено(row.ОбъектРасчетов) else "-"
            except Exception:
                орс = "(нет ОР)"
            стр = ""
            try:
                стр = str(row.СтатьяРасходов) if erp.ЗначениеЗаполнено(row.СтатьяРасходов) else "-"
            except Exception:
                стр = "?"
            print(f"     стр{i}: Партнер={partner} | ОбъектРасчетов={орс} | СтатьяРасходов={стр}")


for dt, тч in [("СписаниеБезналичныхДенежныхСредств", ["РасшифровкаПлатежа"]),
               ("ЗаявкаНаРасходованиеДенежныхСредств", ["РасшифровкаПлатежа"]),
               ("ПоступлениеБезналичныхДенежныхСредств", ["РасшифровкаПлатежа"])]:
    # документы с Неизв в шапке ИЛИ в ТЧ
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ РАЗЛИЧНЫЕ Ссылка КАК С ИЗ Документ." + dt + " ГДЕ Партнер = &P "
              "ОБЪЕДИНИТЬ ВЫБРАТЬ РАЗЛИЧНЫЕ Ссылка ИЗ Документ." + dt + ".РасшифровкаПлатежа ГДЕ Партнер = &P")
    q.УстановитьПараметр("P", P)
    s = q.Execute().Выбрать()
    docs = []
    while s.Следующий():
        docs.append(s.С)
    print(f"\n========== {dt}: {len(docs)} док. (показываю первые 3) ==========")
    for d in docs[:3]:
        o = d.ПолучитьОбъект()
        print(f"\n  ДОКУМЕНТ: {d}  | Помечен={o.ПометкаУдаления} Проведен={o.Проведен}")
        dump_doc(o, тч)
