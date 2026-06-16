# -*- coding: utf-8 -*-
# Что реально осталось со ссылками на Неизвестного ПОСЛЕ прогона формы "Замінити всі":
# разбивка по метаданным + помечен ли на удаление (для справочников).
import sys
sys.path.insert(0, r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test')
import _common_neizv as c

erp = c.connect()
P = c.neizv(erp)
arr = erp.NewObject("Массив"); arr.Добавить(P)
res = erp.НайтиПоСсылкам(arr)

agg = {}      # метаданное -> [всего, помечено, не_помечено]
for i in range(res.Количество()):
    row = res.Получить(i)
    try:
        fn = row.Метаданные.ПолноеИмя()
    except Exception:
        fn = "?"
    данные = row.Данные
    помечен = None
    try:
        помечен = bool(данные.ПометкаУдаления)
    except Exception:
        помечен = None  # не справочник/нет реквизита
    a = agg.setdefault(fn, [0, 0, 0])
    a[0] += 1
    if помечен is True:
        a[1] += 1
    elif помечен is False:
        a[2] += 1

print("Всего ссылок на Неизвестного:", res.Количество())
print(f"{'Метаданное':52} {'всего':>6} {'помечено':>9} {'НЕпомечено':>11}")
for k in sorted(agg, key=lambda x: -agg[x][0]):
    a = agg[k]
    print(f"{k:52} {a[0]:6} {a[1]:9} {a[2]:11}")
