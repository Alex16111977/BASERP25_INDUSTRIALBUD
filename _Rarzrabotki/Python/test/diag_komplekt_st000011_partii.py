# -*- coding: utf-8 -*-
"""Дочитка: реквизиты документов-партий СТ00-000914 (Надходження) и СТ00-000315
(Переміщення) — что именно разыменовывает Субконто2.Подразделение и что менять."""
import sys
from datetime import datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
S = buh.String


def s(x):
    if x is None:
        return "<NULL>"
    try:
        r = S(x)
    except Exception:
        return repr(x)
    return r if r != "" else "<пусто>"


def run_query(text, params):
    q = buh.NewObject("Запрос")
    q.Text = text
    for k, v in params.items():
        q.SetParameter(k, v)
    return q.Execute().Выгрузить()


def dump_doc(meta_name, number, d1, d2, fields):
    t = run_query(
        f"""ВЫБРАТЬ Док.Ссылка КАК Ссылка ИЗ Документ.{meta_name} КАК Док
        ГДЕ Док.Номер = &Номер И Док.Дата МЕЖДУ &Д1 И &Д2""",
        {"Номер": number, "Д1": d1, "Д2": d2})
    print(f"--- Документ.{meta_name} {number}: найдено {t.Количество()}")
    if t.Количество() == 0:
        return None
    ref = t.Получить(0).Ссылка
    print("   ", s(ref), "| Проведен:", ref.Проведен)
    for f in fields:
        try:
            print(f"    {f} = {s(getattr(ref, f))}")
        except Exception as e:
            print(f"    {f}: <нет реквизита> ({type(e).__name__})")
    return ref


print("Реквизиты метаданных (имена, содержащие 'Подраздел'):")
for mn in ("ПоступлениеТоваровУслуг", "ПеремещениеТоваров"):
    meta = getattr(buh.Метаданные.Документы, mn)
    names = []
    for i in range(meta.Реквизиты.Количество()):
        nm = meta.Реквизиты.Получить(i).Имя
        if "одраздел" in nm or nm.startswith("А_") or nm.startswith("_"):
            names.append(nm)
    print(f"  {mn}: {', '.join(names) if names else '<ничего>'}")

print()
d1 = datetime(2026, 6, 1)
d2 = datetime(2026, 6, 30, 23, 59, 59)
post = dump_doc("ПоступлениеТоваровУслуг", "СТ00-000914", d1, d2,
                ["Подразделение", "Склад", "Организация", "Контрагент", "НалоговоеНазначение"])
perem = dump_doc("ПеремещениеТоваров", "СТ00-000315", d1, d2,
                 ["Подразделение", "Подразделение1", "СкладОтправитель", "СкладПолучатель", "Организация"])
