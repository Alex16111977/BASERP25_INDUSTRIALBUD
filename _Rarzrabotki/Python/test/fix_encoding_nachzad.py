# -*- coding: utf-8 -*-
"""Перекодировка исходников обработки А_НачальнаяЗадолженностьПоЗарплатеСозданнаяПоВыплатам
в UTF-8 BOM + CRLF."""
import io
import os

BASE = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "Обработки"))
NAME = "А_НачальнаяЗадолженностьПоЗарплатеСозданнаяПоВыплатам"
FILES = [
    NAME + ".xml",
    os.path.join(NAME, "Ext", "ObjectModule.bsl"),
    os.path.join(NAME, "Forms", "Форма.xml"),
    os.path.join(NAME, "Forms", "Форма", "Ext", "Form.xml"),
    os.path.join(NAME, "Forms", "Форма", "Ext", "Form", "Module.bsl"),
]
for rel in FILES:
    p = os.path.join(BASE, rel)
    with io.open(p, "r", encoding="utf-8-sig") as fh:
        text = fh.read()
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    with io.open(p, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write(text)
    print("OK", rel)
print("DONE")
