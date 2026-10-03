import sys, os; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from com_conn import *
c = connect('buh')
md = c.Метаданные
def dump(mdobj):
    print('===', mdobj.Имя, '|', c.String(mdobj.Синоним))
    for a in mdobj.Реквизиты:
        print('  A:', a.Имя, '|', c.String(a.Синоним), '|', c.String(a.Тип))
    for ts in mdobj.ТабличныеЧасти:
        print('  TS:', ts.Имя, '|', c.String(ts.Синоним))
        for a in ts.Реквизиты:
            print('     -', a.Имя, '|', c.String(a.Синоним), '|', c.String(a.Тип))
dump(md.Справочники.СтруктураСебестоимости)
dump(md.Справочники.ОбщиеНазванияНоменклатуры)
dump(md.Документы.КомплектацияНоменклатуры)
n = md.Справочники.Номенклатура
for a in n.Реквизиты:
    nm = a.Имя
    if 'Общ' in nm or 'А_' in nm or 'Комплект' in nm or 'Загальн' in nm:
        print('  Ном.A:', nm, '|', c.String(a.Синоним), '|', c.String(a.Тип))
for ts in n.ТабличныеЧасти:
    print('  Ном.TS:', ts.Имя, '|', c.String(ts.Синоним))
# any catalog/register about kits
for coll, nm in [(md.РегистрыСведений, 'РС'), (md.Справочники, 'Спр')]:
    for o in coll:
        if 'омплект' in o.Имя or 'Соответств' in o.Имя and 'Номенкл' in o.Имя or 'ОбщиеНазв' in o.Имя:
            print(nm, o.Имя, '|', c.String(o.Синоним))
