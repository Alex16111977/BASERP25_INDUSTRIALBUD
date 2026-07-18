# -*- coding: utf-8 -*-
# Smoke (bas-verify): заполнение ТЧ НачисленияБухЗасчетФондов через метод документа.
# In-memory объект (НЕ записывается) -> 0 мутаций реальных данных.
# Эталон дек.2025: новая ТЧ = 4 строки, Σ=6 349,20, СчетКт=661, только вид 00021.
# Регрессия: НачисленияБух наполнена и НЕ содержит 00021; Налоги/Удержания наполнены.
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected BaseERP")

doc = erp.Документы.А_ОтражениеЗПпоКазне.СоздатьДокумент()
doc.Дата = datetime.datetime(2025, 12, 31, 12, 0, 0)   # месяц = декабрь 2025
print("Created in-memory doc, Дата =", doc.Дата)

print("Calling ЗагрузитьНалогиИУдержанияИзЗарплатыБух() ...")
doc.ЗагрузитьНалогиИУдержанияИзЗарплатыБух()
print("Call returned OK\n")

def tc_sum(tc):
    s = 0.0
    for i in range(tc.Количество()):
        s += float(tc.Получить(i).Сумма)
    return s

fond = doc.НачисленияБухЗасчетФондов
buh  = doc.НачисленияБух
nal  = doc.НалогиБухгалтерия
ud   = doc.УдержанияБухгалтерия

print(f"=== НачисленияБухЗасчетФондов: строк={fond.Количество()}, Σ={tc_sum(fond):,.2f} (ожид. 4 / 6 349,20)")
codes_fond = set()
acc_ok = True
sub_ok = True
for i in range(fond.Количество()):
    r = fond.Получить(i)
    codes_fond.add(str(r.КодВидРасчета))
    acc = r.СчетКт.Код if erp.ЗначениеЗаполнено(r.СчетКт) else "<пусто>"
    if acc != "661":
        acc_ok = False
    # СубконтоКт1 == ФизЛицо
    if erp.ЗначениеЗаполнено(r.СубконтоКт1) and erp.ЗначениеЗаполнено(r.ФизЛицо):
        if erp.XMLСтрока(r.СубконтоКт1) != erp.XMLСтрока(r.ФизЛицо):
            sub_ok = False
    print(f"  {r.ФИО} | код={r.КодВидРасчета} вид='{r.ВидРасчета}' | {float(r.Сумма):,.2f} | СчетКт={acc} | орг={r.Организация}")
print(f"  Коды видов в ТЧ: {codes_fond} (ожид. {{'00021'}})")
print(f"  Все СчетКт==661: {acc_ok} | СубконтоКт1==ФизЛицо: {sub_ok}")

# Регрессия
codes_buh = set()
for i in range(buh.Количество()):
    codes_buh.add(str(buh.Получить(i).КодВидРасчета))
print(f"\n=== Регрессия:")
print(f"  НачисленияБух: строк={buh.Количество()}, Σ={tc_sum(buh):,.2f}; содержит 00021? {'00021' in codes_buh} (ожид. False)")
print(f"  НалогиБухгалтерия: строк={nal.Количество()}")
print(f"  УдержанияБухгалтерия: строк={ud.Количество()}")

# Итог
ok = (fond.Количество() == 4 and abs(tc_sum(fond) - 6349.20) < 0.01
      and codes_fond == {'00021'} and acc_ok and sub_ok
      and '00021' not in codes_buh and buh.Количество() > 0)
print(f"\n[VERDICT] {'GREEN — все проверки пройдены' if ok else 'RED — см. выше'}")

erp = None
print("\nDone (doc НЕ записан).")
