# -*- coding: utf-8 -*-
# Фаза 7 smoke: пріоритет плану в розподілі — баланс цілий + НормаНоменкл один раз на номенклатуру.
import win32com.client, sys, os, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
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
vn = pn = ost = 0.0
по_ном = {}
план_рядків = 0
for i in range(n):
    r = тч.Получить(i)
    vn += float(r.ВНорме); pn += float(r.ПонадНорму); ost += float(r.Остаток)
    nm = buh.String(r.Номенклатура)
    по_ном.setdefault(nm, 0.0)
    по_ном[nm] += float(r.НормаНоменкл)
    if float(r.НормаНоменкл) > 0:
        план_рядків += 1
планових_номенкл = sum(1 for v in по_ном.values() if v > 0)
print(f"рядків={n} баланс={abs(vn+pn-ost):.3f} планових_номенкл={планових_номенкл} рядків_з_планом={план_рядків}")
assert n > 0, "0 рядків"
assert abs(vn + pn - ost) < 0.001, "баланс кількості зламано"
assert план_рядків == планових_номенкл, "НормаНоменкл задвоюється або відсутня"
print("PASS")
