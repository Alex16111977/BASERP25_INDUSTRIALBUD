# -*- coding: utf-8 -*-
# Фаза 6 smoke: баланс сум по рядку (ВНормеСумма+ПонадНормуСумма=СуммаОстатка),
# одиниця на всіх рядках, баланс кількості не зламано.
import win32com.client, sys, os, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета.epf"
assert os.path.exists(EPF), "немає .epf — спершу збери"
обр = buh.ВнешниеОбработки.Создать(EPF, False)
обр.Период = datetime.datetime(2026, 6, 25)
обр.Спецификация = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000005").Ссылка
выб = buh.Справочники.Склады.Выбрать()
while выб.Следующий():
    н = выб.Наименование
    if "МХП" in н and "ОР" in н and "виробнич" in н:
        обр.СкладыОстатков.Добавить().Склад = выб.Ссылка
выбО = buh.Справочники.Организации.Выбрать()
while выбО.Следующий():
    if "ІНДЕПТ" in выбО.Наименование:
        обр.Организация = выбО.Ссылка
обр.РассчитатьАнализ()
тч = обр.ТабличнаяЧастьОстатков
n = тч.Количество()
bad_bal = 0; no_unit = 0; sum_ost = 0.0; vn = 0.0; pn = 0.0; vns = 0.0; pns = 0.0
for i in range(n):
    r = тч.Получить(i)
    if abs(float(r.ВНормеСумма) + float(r.ПонадНормуСумма) - float(r.СуммаОстатка)) > 0.01:
        bad_bal += 1
    if r.Единица is None or buh.String(r.Единица) == "":
        no_unit += 1
    sum_ost += float(r.СуммаОстатка); vn += float(r.ВНорме); pn += float(r.ПонадНорму)
    vns += float(r.ВНормеСумма); pns += float(r.ПонадНормуСумма)
qty_ost = sum(float(тч.Получить(i).Остаток) for i in range(n))
print(f"рядків={n} Σсума={sum_ost:.2f} (ВНорме={vns:.2f}+Понад={pns:.2f}) "
      f"баланс_кільк={abs(vn+pn-qty_ost):.3f} порушень_сум={bad_bal} без_одиниці={no_unit}")
assert n > 0, "0 рядків"
assert bad_bal == 0, f"{bad_bal} рядків з порушенням балансу сум"
assert no_unit == 0, f"{no_unit} рядків без одиниці"
assert abs(vn + pn - qty_ost) < 0.001, "баланс кількості зламано"
print("PASS")
