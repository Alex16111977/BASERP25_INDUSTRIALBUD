# -*- coding: utf-8 -*-
# §7 SMOKE (in-memory, БЕЗ записи): на эталоне №000000003 вызвать ЗаполнитьДляОтражения
# (теперь с новым проходом _ОтрЗП_ВыровнятьОргБухФорма1ПоНачисленияБух) и проверить НЗ per-org.
# Документ НЕ записывается → боевые данные не меняются.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected ERP")

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ Ссылка КАК С ИЗ Документ.А_ОтражениеЗПпоКазне ГДЕ Номер = "000000003" И ГОД(Дата)=2026 И МЕСЯЦ(Дата)=2'
sel = q.Выполнить().Выбрать(); sel.Следующий()
obj = sel.С.ПолучитьОбъект()
print("Документ:", obj.Номер, "Проведен(persisted):", erp.XMLСтрока(sel.С))

# GROSS-веса из НачисленияБух (до fill — это исходная ТЧ документа)
buh = {}
tc = obj.НачисленияБух
for i in range(tc.Количество()):
    r = tc.Получить(i)
    fl = r.ФизЛицо.Наименование; org = r.Организация.Наименование
    buh.setdefault(fl, {})[org] = round(buh.get(fl, {}).get(org, 0) + float(r.Сумма), 2)

# === ВЫЗОВ FILL В ПАМЯТИ (с новым проходом) ===
try:
    obj.ЗаполнитьДляОтражения_ДокументОтражениеЗарплатыВФинансовомУчете()
    print("Fill OK (в памяти, без записи)\n")
except Exception as e:
    msg = e.excepinfo[2] if (hasattr(e,'excepinfo') and e.excepinfo) else str(e)
    print("FILL FAILED:", msg); sys.exit(1)

# Собрать НЗ Форма1 per (ФЛ, ОргБух): Сумма (net → Приход) и Сумма+Взносы
ф1 = erp.XMLСтрока(erp.Перечисления.А_ФормыPL.Форма1)
nz = {}
tc = obj.НачисленнаяЗарплатаИВзносыПоФизлицам
for i in range(tc.Количество()):
    r = tc.Получить(i)
    if erp.XMLСтрока(r.ФормаPL) != ф1: continue
    fl = r.ФизическоеЛицо.Наименование
    org = r.ОрганизацияБухгалтерия.Наименование if erp.ЗначениеЗаполнено(r.ОрганизацияБухгалтерия) else "(пусто)"
    d = nz.setdefault(fl, {}).setdefault(org, [0.0, 0.0])
    d[0] += float(r.Сумма); d[1] += float(r.Сумма)+float(r.ВзносыВсего)

fixed = ["Аулова Маргарита Володимирівна","Воронцов Олександр Володимирович","Шинкаренко Олег Павлович",
         "Шкурат Ігор Миколайович","Шостак Жанета Германівна"]
balanced = ["Арановська Наталя Анатоліївна","Манжула Максим Анатолійович","Пономаренко Віталій Вікторович"]
def show(group, names):
    print(f"=== {group} ===")
    for f in names:
        gtot = round(sum(buh.get(f,{}).values()),2)
        net  = round(sum(v[0] for v in nz.get(f,{}).values()),2)
        full = round(sum(v[1] for v in nz.get(f,{}).values()),2)
        print(f"{f}  (GROSS={gtot:,.2f}  NET-Ф1={net:,.2f})")
        for o in sorted(set(buh.get(f,{}))|set(nz.get(f,{}))):
            g = buh.get(f,{}).get(o,0); netо = nz.get(f,{}).get(o,[0,0])[0]
            exp = round(net*g/gtot,2) if gtot else 0
            flag = "OK" if abs(netо-exp)<=0.10 else "XX"
            print(f"     {o[:26]:27} GROSS={g:11,.2f}  NET-Ф1(после)={netо:11,.2f}  ожид(GROSSдоля)={exp:11,.2f}  {flag}")
        print()
show("5 ИСПРАВЛЯЕМЫХ (ДЖИ ТРИ) — ожидаем per-org = GROSS-доля", fixed)
show("БАЛАНСНЫЕ (ІНДЕПТ) — ожидаем без коллапса", balanced)
print("⚠️ Документ НЕ записан. Боевые данные не изменены. Перепроведение — пользователь в UI.")
erp=None
