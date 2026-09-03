# -*- coding: utf-8 -*-
"""Пост-патч СКД после skd-compile: ограничить условное оформление областью данных.

skd-compile (JSON DSL) не умеет флаги области применения условного оформления, поэтому
после перекомпиляции подписи колонок в шапке красятся (условие в заголовке считается от
общих итогов). Скрипт вставляет в КАЖДЫЙ элемент conditionalAppearance после </dcsset:presentation>:
    <dcsset:useInFieldsHeader>DontUse</dcsset:useInFieldsHeader>
    <dcsset:useInHeader>DontUse</dcsset:useInHeader>

Использование:
    python skd_postpatch_dontuse_header_2026-09-03.py <путь к Template.xml> [ожидаемое число элементов]

Пример (отчёт А_ОтчетПоСтатьямPL_ДДС_СтатьямРасходам, 3 варианта × 3 элемента = 9):
    python _Rarzrabotki/Python/test/skd_postpatch_dontuse_header_2026-09-03.py ^
      "_Rarzrabotki/Отчеты/А_ОтчетПоСтатьямPL_ДДС_СтатьямРасходам/Templates/ОсновнаяСхемаКомпоновкиДанных/Ext/Template.xml" 9
"""
import re
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

path = sys.argv[1]
expected = int(sys.argv[2]) if len(sys.argv) > 2 else None

raw = open(path, "rb").read()
bom = raw[:3] == b"\xef\xbb\xbf"
s = raw.decode("utf-8-sig")
if "useInFieldsHeader" in s:
    print("уже пропатчен, ничего не делаю")
    sys.exit(0)

pat = re.compile(r'</dcsset:presentation>(\r?\n)([ \t]*)</dcsset:item>')


def rep(m):
    nl, ind = m.group(1), m.group(2)
    return ("</dcsset:presentation>" + nl + ind + "\t<dcsset:useInFieldsHeader>DontUse</dcsset:useInFieldsHeader>"
            + nl + ind + "\t<dcsset:useInHeader>DontUse</dcsset:useInHeader>" + nl + ind + "</dcsset:item>")


s2, n = pat.subn(rep, s)
if expected is not None and n != expected:
    print(f"ОШИБКА: вставок {n}, ожидалось {expected} — файл не записан")
    sys.exit(1)
open(path, "wb").write((b"\xef\xbb\xbf" if bom else b"") + s2.encode("utf-8"))
print(f"пропатчено элементов условного оформления: {n}")
