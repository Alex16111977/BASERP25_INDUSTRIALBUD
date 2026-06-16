# -*- coding: utf-8 -*-
"""
Override Подразделения для ПАП-строки "Денежные средства в пути" в проекции упр.баланса:
DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl, процедура
ДобавитьТекстЗапросаДенежныеСредстваВПути, каскад МАКСИМУМ(ВЫБОР ...).
Вставляем приоритетную ветку ДокПКО.А_ПодразделениеОтправитель СРАЗУ ПОСЛЕ ветки РКО
(перед веткой ДокПКО.Подразделение). РКО-плечо не затрагивается (для РКО ДокПКО=NULL).
Гард: ровно 1 вставка; ветки ещё нет.
"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)

PATH = r"C:\Configuration_downloads\BASERP25\DataProcessors\ДвиженияАктивовПассивов\Ext\ManagerModule.bsl"
with open(PATH, "r", encoding="utf-8") as f:
    text = f.read()

ps = text.find("Процедура ДобавитьТекстЗапросаДенежныеСредстваВПути")
if ps < 0:
    print("[FAIL] процедура ДобавитьТекстЗапросаДенежныеСредстваВПути не найдена"); sys.exit(1)
pe = text.find("КонецПроцедуры", ps)
proc = text[ps:pe]

if "А_ПодразделениеОтправитель" in proc:
    print("[STOP] А_ПодразделениеОтправитель уже есть в процедуре"); sys.exit(2)

# Якорь: строка "ТОГДА ДокРКО.Подразделение" + следующая "КОГДА ДокПКО.Подразделение <> ..."
pat = re.compile(
    r"(^[ \t]*\|[ \t]*ТОГДА ДокРКО\.Подразделение[ \t]*\n)"
    r"([ \t]*\|([ \t]*)КОГДА ДокПКО\.Подразделение <> ЗНАЧЕНИЕ\(Справочник\.СтруктураПредприятия\.ПустаяСсылка\))",
    re.M)
m = pat.search(proc)
if not m:
    print("[FAIL] якорь (ТОГДА ДокРКО.Подразделение / КОГДА ДокПКО.Подразделение) не найден"); sys.exit(3)

# Префикс перед "КОГДА" во второй строке (например "\t|\t\t")
kogda_line = m.group(2)
prefix = kogda_line[:kogda_line.index("КОГДА")]      # "\t|\t\t"
inner = prefix + "\t"                                  # доп. отступ для строки "И ..."

new_branch = (
    f"{prefix}КОГДА ДокПКО.А_ПодразделениеОтправитель <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)\n"
    f"{inner}И ДокПКО.А_ПодразделениеОтправитель ЕСТЬ НЕ NULL\n"
    f"{prefix}ТОГДА ДокПКО.А_ПодразделениеОтправитель\n"
)

def repl(mm):
    return mm.group(1) + new_branch + mm.group(2)

proc_new, n = pat.subn(repl, proc, count=1)
if n != 1:
    print(f"[FAIL] ожидалась 1 вставка, получено {n}"); sys.exit(4)

text_new = text[:ps] + proc_new + text[pe:]
with open(PATH, "w", encoding="utf-8") as f:
    f.write(text_new)
print(f"[OK] вставлена приоритетная ветка А_ПодразделениеОтправитель ({n})")

print("\n=== Контроль (каскад Подразделение) ===")
seg = proc_new[proc_new.find("МАКСИМУМ(ВЫБОР"):proc_new.find("КОНЕЦ) КАК Подразделение")+25]
for ln in seg.split("\n"):
    if "ДокРКО" in ln or "ДокПКО" in ln or "А_ПодразделениеОтправитель" in ln:
        print("  " + ln)
