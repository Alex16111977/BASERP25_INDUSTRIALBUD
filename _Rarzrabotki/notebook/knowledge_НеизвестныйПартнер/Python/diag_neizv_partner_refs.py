# -*- coding: utf-8 -*-
# Диагностика: где реально ссылается "Неизвестный партнер" (платформенный НайтиПоСсылкам)
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

partner = erp.Справочники.Партнеры.НайтиПоНаименованию("Неизвестный партнер", True)
if not erp.ЗначениеЗаполнено(partner):
    try:
        partner = erp.Справочники.Партнеры.НеизвестныйПартнер
    except Exception:
        pass

print("Наименование:", partner.Наименование)
print("UUID:", erp.String(partner.Ссылка.УникальныйИдентификатор()))

# Дубли по маске (через параметр)
q = erp.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ Ссылка, Наименование, ПометкаУдаления ИЗ Справочник.Партнеры "
          "ГДЕ Наименование ПОДОБНО &Маска")
q.УстановитьПараметр("Маска", "%еизвестн%")
sel = q.Execute().Выбрать()
print("\n--- Партнеры по маске 'еизвестн' ---")
while sel.Следующий():
    print(" *", sel.Наименование, "| del=", sel.ПометкаУдаления,
          "|", erp.String(sel.Ссылка.УникальныйИдентификатор()))

# НайтиПоСсылкам — глобальная картина
arr = erp.NewObject("Массив")
arr.Добавить(partner)
try:
    res = erp.НайтиПоСсылкам(arr)
    print("\n--- НайтиПоСсылкам: всего строк =", res.Количество(), "---")
    cols = [c.Имя for c in res.Колонки]
    print("Колонки:", cols)

    agg = {}
    n = res.Количество()
    for i in range(n):
        row = res.Получить(i)
        try:
            fullname = row.Метаданные.ПолноеИмя()
        except Exception:
            fullname = "?"
        agg[fullname] = agg.get(fullname, 0) + 1
    print("\n--- Сводка по объектам метаданных (где встречается ссылка) ---")
    for k in sorted(agg, key=lambda x: -agg[x]):
        print(f"  {agg[k]:6d}  {k}")
    print("\nИТОГО объектов-источников:", n, "| уникальных метаданных:", len(agg))
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print("НайтиПоСсылкам FAIL:", e.excepinfo[2])
    else:
        print("НайтиПоСсылкам FAIL:", e)
