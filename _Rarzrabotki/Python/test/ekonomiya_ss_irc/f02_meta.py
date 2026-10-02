# -*- coding: utf-8 -*-
"""Ф0. Метаданные объектов исследования (реквизиты, ТЧ, измерения/ресурсы) в ERP и BuhBud — только чтение."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from common import *


def coll(c):
    return [c.Получить(i) for i in range(c.Количество())]


def typ(conn, md):
    try:
        return conn.String(md.Тип)
    except Exception:
        return "?"


def describe(conn, kind, name):
    out = [f"### {kind}.{name}"]
    M = getattr(conn.Метаданные, kind)
    md = M.Найти(name)
    if md is None:
        out.append("  НЕ НАЙДЕН")
        return out
    for part in ("Измерения", "Ресурсы", "Реквизиты"):
        try:
            c = getattr(md, part)
        except Exception:
            continue
        out.append(f"  [{part}] " + "; ".join(f"{a.Имя}:{typ(conn, a)[:60]}" for a in coll(c)))
    try:
        for t in coll(md.ТабличныеЧасти):
            out.append(f"  [ТЧ {t.Имя}] " + "; ".join(f"{a.Имя}:{typ(conn, a)[:50]}" for a in coll(t.Реквизиты)))
    except Exception:
        pass
    return out


E = erp()
Bh = buh()
res = []
for kind, name in [("Справочники", "А_СтруктураСебестоимости"), ("РегистрыНакопления", "СебестоимостьТоваров"),
                   ("РегистрыНакопления", "А_ПредварительныйБюджетНаМесяц"), ("РегистрыНакопления", "А_ДвиженияДенегИзКазны"),
                   ("Справочники", "А_ОбщиеНазванияНоменклатуры")]:
    res += ["ERP"] + describe(E, kind, name)
for kind, name in [("Справочники", "СтруктураСебестоимости"), ("Документы", "РасчетКомплектаций"),
                   ("Справочники", "ОбщиеНазванияНоменклатуры"), ("Документы", "КомплектацияНоменклатуры"),
                   ("Документы", "ПеремещениеТоваров")]:
    res += ["BUH"] + describe(Bh, kind, name)
p = os.path.join(OUT, "meta.txt")
open(p, "w", encoding="utf-8").write("\n".join(res))
print(p)
