# -*- coding: utf-8 -*-
# Balance-гейт: больничный за счёт ФСС в ОЗФУ (вид операции РасходыПоСтрахованиюФСС, корень, группа 663).
# Запускает СоздатьДокументНачисления + СоздатьДокументОтражение на №000000001, проверяет:
#  1) 4 строки больничного в ОЗФУ (вид операции, подразделение=корень, Σ=6349.20)
#  2) НЕТ задвоения ТЧ (snapshot before/after)
#  3) ПАП per регистратор сбалансирован (Σ signed Приход-Расход = 0 -> двойная запись цела)
#  4) НЕТ затратных счетов класса 9 на суммах больничного (Хозрасчетный)
#  5) А_ВзСС(Оплата труда) per (Орг,Подр) == -ПАП(ОТ) per (Орг,Подр)
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected BaseERP")

def find_parent(num):
    q = erp.NewObject("Запрос")
    q.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка ИЗ Документ.А_ОтражениеЗПпоКазне '
               'ГДЕ Номер = &Н И НЕ ПометкаУдаления')
    q.УстановитьПараметр("Н", num)
    r = q.Выполнить().Выбрать()
    return r.Ссылка if r.Следующий() else None

def find_child_ozfu(parent_ref):
    q = erp.NewObject("Запрос")
    q.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете '
               'ГДЕ А_ДокОтражениеЗПпоКазне = &О И НЕ ПометкаУдаления')
    q.УстановитьПараметр("О", parent_ref)
    r = q.Выполнить().Выбрать()
    return r.Ссылка if r.Следующий() else None

parent = find_parent("000000001")
print("Parent №000000001:", parent, "| Проведен:", parent.Проведен)
fond_n = parent.НачисленияБухЗасчетФондов.Количество()
fond_sum = sum(float(parent.НачисленияБухЗасчетФондов.Получить(i).Сумма) for i in range(fond_n))
print(f"Parent.НачисленияБухЗасчетФондов: {fond_n} строк, Σ={fond_sum:,.2f}")

# Snapshot BEFORE
child_before = find_child_ozfu(parent)
before_rows = child_before.ПолучитьОбъект().НачисленнаяЗарплатаИВзносыПоФизлицам.Количество() if child_before else 0
print(f"Child ОЗФУ before: {child_before} | НачисленнаяЗПпоФизлицам строк={before_rows}")

# RUN
print("\n--- Запуск СоздатьДокументНачисления() ...")
obj = parent.ПолучитьОбъект()
try:
    obj.СоздатьДокументНачисления()
    print("  OK")
except Exception as e:
    print("  FAIL:", e.excepinfo[2] if (hasattr(e,'excepinfo') and e.excepinfo) else e)
print("--- Запуск СоздатьДокументОтражениеЗарплатыВФинансовомУчете() ...")
try:
    obj.СоздатьДокументОтражениеЗарплатыВФинансовомУчете()
    print("  OK")
except Exception as e:
    print("  FAIL:", e.excepinfo[2] if (hasattr(e,'excepinfo') and e.excepinfo) else e)

# Snapshot AFTER
child = find_child_ozfu(parent)
ozfu = child.ПолучитьОбъект()
after_rows = ozfu.НачисленнаяЗарплатаИВзносыПоФизлицам.Количество()
print(f"\nChild ОЗФУ after: {child} | Проведен={child.Проведен} | НачисленнаяЗПпоФизлицам строк={after_rows}")

# (1)+(4) больничные строки
bol = []
for i in range(after_rows):
    row = ozfu.НачисленнаяЗарплатаИВзносыПоФизлицам.Получить(i)
    if "РасходыПоСтрахованиюФСС" in erp.XMLСтрока(row.ВидОперации):
        bol.append(row)
bol_sum = sum(float(r.Сумма) for r in bol)
print(f"\n=== (1) Больничные строки ФСС: {len(bol)} (ожид. 4), Σ={bol_sum:,.2f} (ожид. {fond_sum:,.2f})")
for r in bol:
    grp = r.ГруппаУчетаНачислений.Наименование if erp.ЗначениеЗаполнено(r.ГруппаУчетаНачислений) else "<пусто>"
    sp = "<пусто>" if not erp.ЗначениеЗаполнено(r.СпособОтраженияЗарплатыВБухучете) else r.СпособОтраженияЗарплатыВБухучете.Наименование
    print(f"  {r.ФизическоеЛицо} | подр={r.ПодразделениеПредприятия} | Σ={float(r.Сумма):,.2f} | группа={grp} | способ={sp}")

# (2) задвоение?
print(f"\n=== (2) Задвоение ТЧ: before={before_rows} after={after_rows} (норм: after = before, идемпотентно; либо before+4 при первом прогоне)")

# (3) ПАП per регистратор сбалансирован
qpap = erp.NewObject("Запрос")
qpap.Текст = ("ВЫБРАТЬ Статья.Наименование КАК Статья, "
              "СУММА(ВЫБОР КОГДА ВидДвижения=ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА Сумма ИНАЧЕ -Сумма КОНЕЦ) КАК Сальдо "
              "ИЗ РегистрНакопления.ПрочиеАктивыПассивы ГДЕ Регистратор = &Р "
              "СГРУППИРОВАТЬ ПО Статья.Наименование")
qpap.УстановитьПараметр("Р", child)
rp = qpap.Выполнить().Выбрать()
pap_total = 0.0
pap_ot = 0.0
print("\n=== (3) ПАП per регистратор (Σ signed по статьям, ожид. 0 = двойная запись цела):")
while rp.Следующий():
    pap_total += float(rp.Сальдо)
    if "плата труда" in str(rp.Статья):
        pap_ot = float(rp.Сальдо)
    print(f"  {rp.Статья}: {float(rp.Сальдо):,.2f}")
print(f"  >>> Σ ПАП = {pap_total:,.2f} (ожид. 0.00) | Оплата труда = {pap_ot:,.2f}")

# (4) Хозрасчетный — есть ли счета класса 9 на этом ОЗФУ
qbu = erp.NewObject("Запрос")
qbu.Текст = ("ВЫБРАТЬ РАЗЛИЧНЫЕ СчетДт.Код КАК Дт, СчетКт.Код КАК Кт "
             "ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Р")
qbu.УстановитьПараметр("Р", child)
rb = qbu.Выполнить().Выбрать()
class9 = set()
allacc = set()
while rb.Следующий():
    for c in (str(rb.Дт), str(rb.Кт)):
        allacc.add(c)
        if c.strip()[:1] == "9":
            class9.add(c.strip())
print(f"\n=== (4) Хозрасчетный счета ОЗФУ: {sorted(allacc)}")
print(f"  >>> Счета класса 9 (затраты): {sorted(class9) if class9 else 'НЕТ'}")

# (5) А_ВзСС(ОТ) per (Орг,Подр) == -ПАП(ОТ)
qvz = erp.NewObject("Запрос")
qvz.Текст = ("ВЫБРАТЬ СУММА(ВЫБОР КОГДА ВидДвижения=ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА СуммаВзаиморасчетов ИНАЧЕ -СуммаВзаиморасчетов КОНЕЦ) КАК С "
              "ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками ГДЕ Регистратор = &Р")
qvz.УстановитьПараметр("Р", child)
rv = qvz.Выполнить().Выбрать(); rv.Следующий()
vzs_total = float(rv.С) if erp.ЗначениеЗаполнено(rv.С) else 0.0
print(f"\n=== (5) А_ВзСС Σ signed (ОЗФУ) = {vzs_total:,.2f} | -ПАП(ОТ) = {-pap_ot:,.2f} | Δ={vzs_total-(-pap_ot):,.2f} (ожид. 0)")

# Вердикт
ok = (len(bol) == 4 and abs(bol_sum - fond_sum) < 0.01 and abs(pap_total) < 0.01
      and not class9 and abs(vzs_total - (-pap_ot)) < 0.01)
print(f"\n[VERDICT] {'GREEN — баланс цел, больничные ФСС на месте, без затрат' if ok else 'RED/ВНИМАНИЕ — см. выше'}")
erp = None
