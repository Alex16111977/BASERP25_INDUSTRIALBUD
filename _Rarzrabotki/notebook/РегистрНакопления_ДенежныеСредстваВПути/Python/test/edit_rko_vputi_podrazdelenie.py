# -*- coding: utf-8 -*-
"""
Детерминированная правка: добавить `ДанныеДокумента.Подразделение КАК Подразделение`
в каждый из 3 блоков функции ТекстЗапросаДенежныеСредстваВПути документа РасходныйКассовыйОрдер.

Скоуп — ТОЛЬКО тело этой функции (между 'Функция ТекстЗапросаДенежныеСредстваВПути('
и следующим 'КонецФункции'). Вставка после строки с 'ДанныеДокумента.Организация ... КАК Организация,'.
Авто-выравнивание КАК по той же колонке. Гард: ровно 3 вставки.
"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)

PATH = r"C:\Configuration_downloads\BASERP25\Documents\РасходныйКассовыйОрдер\Ext\ManagerModule.bsl"

with open(PATH, "r", encoding="utf-8") as f:
    text = f.read()

hdr = "Функция ТекстЗапросаДенежныеСредстваВПути("
i = text.find(hdr)
if i < 0:
    print("[FAIL] функция ТекстЗапросаДенежныеСредстваВПути не найдена"); sys.exit(1)
j = text.find("КонецФункции", i)
if j < 0:
    print("[FAIL] КонецФункции не найден"); sys.exit(1)
head, body, tail = text[:i], text[i:j], text[j:]

# Гард: Подразделение в этой функции ещё нет
if "Подразделение" in body:
    print("[STOP] 'Подразделение' уже присутствует в функции — правка не нужна / уже сделана"); sys.exit(2)

lines = body.split("\n")
out = []
inserted = 0
# Строка вида:  \t|\tДанныеДокумента.Организация<spaces>КАК Организация,
org_re = re.compile(r"^(\s*\|\s*)ДанныеДокумента\.Организация(\s+)КАК Организация,\s*$")
for ln in lines:
    out.append(ln)
    m = org_re.match(ln)
    if m:
        prefix = m.group(1)                       # \t|\t
        # колонка, где начинается 'КАК' в исходной строке
        kak_col = ln.index("КАК Организация,")
        field = prefix + "ДанныеДокумента.Подразделение"
        pad = kak_col - len(field)
        if pad < 1:
            pad = 1
        podr = field + (" " * pad) + "КАК Подразделение,"
        out.append(podr)
        inserted += 1

if inserted != 3:
    print(f"[FAIL] ожидалось 3 вставки, получено {inserted} — НЕ записываю"); sys.exit(3)

new_body = "\n".join(out)
new_text = head + new_body + tail
with open(PATH, "w", encoding="utf-8") as f:
    f.write(new_text)

print(f"[OK] Вставлено {inserted} строк 'Подразделение'. Файл записан: {PATH}")
print("\n=== Контроль: фрагмент функции после правки ===")
for k, ln in enumerate(new_body.split("\n")):
    if "Подразделение" in ln or "Организация," in ln or "ВыдачаДенежныхСредствВДругуюКассу" in ln \
       or "ИнкассацияДенежныхСредствВБанк" in ln or "КонвертацияВалюты" in ln:
        print(f"  {ln}")
