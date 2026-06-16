# -*- coding: utf-8 -*-
# Можно ли автоматически вывести партнёра для остатка?
# Платёжные ТЧ: есть ли ОбъектРасчетов (и у него Контрагент/Партнер)? РеестрДокументов: контрагент регистратора?
import sys
sys.path.insert(0, r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test')
import _common_neizv as c

erp = c.connect()
P = c.neizv(erp)


def real(p):
    return erp.ЗначениеЗаполнено(p) and not c.same(erp, p, P)


# --- Платёжные документы: строки РасшифровкаПлатежа с Партнер=Неизв ---
for dt in ("СписаниеБезналичныхДенежныхСредств", "ЗаявкаНаРасходованиеДенежныхСредств",
           "ПоступлениеБезналичныхДенежныхСредств"):
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ РАЗЛИЧНЫЕ Ссылка КАК С ИЗ Документ." + dt +
              ".РасшифровкаПлатежа ГДЕ Партнер = &P")
    q.УстановитьПараметр("P", P)
    s = q.Execute().Выбрать()
    docs = []
    while s.Следующий():
        docs.append(s.С)
    print(f"\n=== {dt}: документов со строками Неизв = {len(docs)} ===")
    cnt_or = cnt_orkontr = cnt_hdrkontr = cnt_none = 0
    for d in docs[:50]:
        o = d.ПолучитьОбъект()
        hdr_k = None
        try:
            hdr_k = o.Контрагент
        except Exception:
            pass
        for row in o.РасшифровкаПлатежа:
            if not c.same(erp, row.Партнер, P):
                continue
            орс = None
            try:
                орс = row.ОбъектРасчетов
            except Exception:
                pass
            if орс is not None and erp.ЗначениеЗаполнено(орс):
                cnt_or += 1
                if erp.ЗначениеЗаполнено(орс.Контрагент):
                    cnt_orkontr += 1
            elif hdr_k is not None and erp.ЗначениеЗаполнено(hdr_k):
                cnt_hdrkontr += 1
            else:
                cnt_none += 1
    print(f"  строк: есть ОбъектРасчетов={cnt_or} (из них с Контрагентом={cnt_orkontr}), "
          f"только шапка-Контрагент={cnt_hdrkontr}, НИЧЕГО={cnt_none}")

# --- РеестрДокументов: контрагент регистратора ---
print("\n=== РеестрДокументов: derivability по регистратору ===")
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Ссылка КАК Рег, ТипСсылки.Имя КАК Тип ИЗ РегистрСведений.РеестрДокументов ГДЕ Партнер = &P"
q.УстановитьПараметр("P", P)
s = q.Execute().Выбрать()
by_type = {}
с_контр = 0
всего = 0
while s.Следующий():
    всего += 1
    рег = s.Рег
    t = s.Тип
    by_type[t] = by_type.get(t, 0) + 1
    try:
        k = рег.Контрагент
        if erp.ЗначениеЗаполнено(k) and erp.XMLСтрока(erp.ТипЗнч(k)) == "jcfg:CatalogRef.Контрагенты":
            с_контр += 1
    except Exception:
        pass
print(f"  всего записей {всего}; у регистратора есть Контрагент(Контрагенты)={с_контр}")
print("  по типам регистратора:", by_type)
