# -*- coding: utf-8 -*-
"""Правка ЗаполнитьИзПередачиНачислений: читать ФормаPL из ТЧ (не хардкод Форма1).
Работает СТРОГО внутри тела процедуры — не задевает стр.262 (НачальнаяЗадолженность)."""
path = r"C:\Configuration_downloads\BASERP25\AccumulationRegisters\А_ВзаиморасчетыССотрудниками\Ext\ManagerModule.bsl"
raw = open(path, "rb").read()
has_bom = raw.startswith(b"\xef\xbb\xbf")
text = raw.decode("utf-8-sig")
nl = "\r\n" if "\r\n" in text else "\n"

marker = "Процедура ЗаполнитьИзПередачиНачислений"
i = text.index(marker)
j = text.index("КонецПроцедуры", i) + len("КонецПроцедуры")
proc = text[i:j]
orig = proc

def rep(s, a, b):
    c = s.count(a)
    assert c == 1, f"count={c} (ожид.1) для: {a!r}"
    return s.replace(a, b, 1)

# 1) 1-й подзапрос (Исходное): + ФормаPL
proc = rep(proc,
    f"Н.ОрганизацияБухгалтерия КАК ОрганизацияБухгалтерия{nl}\t|ПОМЕСТИТЬ втСтроки",
    f"Н.ОрганизацияБухгалтерия КАК ОрганизацияБухгалтерия,{nl}\t|\tН.ФормаPL КАК ФормаPL{nl}\t|ПОМЕСТИТЬ втСтроки")
# 2) 2-й подзапрос (Целевое): + ФормаPL
proc = rep(proc,
    f"Н.ОрганизацияБухгалтерия{nl}\t|ИЗ",
    f"Н.ОрганизацияБухгалтерия,{nl}\t|\tН.ФормаPL{nl}\t|ИЗ")
# 3) финальный SELECT: + ФормаPL
proc = rep(proc,
    f"С.ОрганизацияБухгалтерия КАК ОрганизацияБухгалтерия{nl}\t|ИЗ",
    f"С.ОрганизацияБухгалтерия КАК ОрганизацияБухгалтерия,{nl}\t|\tС.ФормаPL КАК ФормаPL{nl}\t|ИЗ")
# 4) GROUP BY: + ФормаPL
proc = rep(proc,
    f"С.ОрганизацияБухгалтерия{nl}\t|{nl}\t|ИМЕЮЩИЕ",
    f"С.ОрганизацияБухгалтерия,{nl}\t|\tС.ФормаPL{nl}\t|{nl}\t|ИМЕЮЩИЕ")
# 5) запись: ФормаPL из строки ТЧ (не хардкод)
proc = rep(proc, "= Перечисления.А_ФормыPL.Форма1;", "= СтрТЗ.ФормаPL;")

assert proc != orig, "процедура не изменилась!"
assert proc.count("Н.ФормаPL КАК ФормаPL") == 1
assert proc.count("С.ФормаPL КАК ФормаPL") == 1
assert proc.count("= СтрТЗ.ФормаPL;") == 1
assert "Перечисления.А_ФормыPL.Форма1" not in proc, "хардкод Форма1 остался в процедуре!"

text2 = text[:i] + proc + text[j:]
# контроль: стр.262 (вне процедуры) НЕ тронута — хардкод Форма1 в файле всё ещё есть (для НачЗадолж)
assert text2.count("Запись.ФормаPL") == text.count("Запись.ФормаPL"), "число Запись.ФормаPL изменилось!"
out = text2.encode("utf-8-sig") if has_bom else text2.encode("utf-8")
open(path, "wb").write(out)
print(f"OK | nl={nl!r} bom={has_bom} | проверки пройдены")
print("Запись.ФормаPL в файле (всего):", text2.count("Запись.ФормаPL"),
      "| из них хардкод Форма1:", text2.count("Запись.ФормаPL             = Перечисления.А_ФормыPL.Форма1;"))
