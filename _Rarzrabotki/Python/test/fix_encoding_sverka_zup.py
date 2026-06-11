# -*- coding: utf-8 -*-
"""Перекодировка исходников отчёта в UTF-8 BOM + CRLF (требование 1С)."""
import io
import os

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "Отчеты")
NAME = "А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП"
FILES = [
    NAME + ".xml",
    os.path.join(NAME, "Ext", "ObjectModule.bsl"),
    os.path.join(NAME, "Ext", "Help.xml"),
    os.path.join(NAME, "Ext", "Help", "ru.html"),
    os.path.join(NAME, "Ext", "Help", "uk.html"),
    os.path.join(NAME, "Templates", "ОсновнаяСхемаКомпоновкиДанных.xml"),
    os.path.join(NAME, "Templates", "ОсновнаяСхемаКомпоновкиДанных", "Ext", "Template.xml"),
]
for rel in FILES:
    p = os.path.normpath(os.path.join(BASE, rel))
    with io.open(p, "r", encoding="utf-8-sig") as fh:
        text = fh.read()
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    with io.open(p, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write(text)
    print("OK", rel)
print("DONE")
