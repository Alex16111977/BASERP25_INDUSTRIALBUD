# -*- coding: utf-8 -*-
"""Запись строк счетов Казны в РС А_ТоварыСчетовПоставщиков (BaseERP) + протокол xlsx.

Вход: RAB/scheta.json (kazna_scheta.py) и RAB/<uid>/stroki.json (raspoznat.py / ИИ).
Набор записей по СчетКазнаУИД заменяется целиком — повторный запуск не плодит строки.
Новые карточки номенклатуры НЕ создаются (sopostavit.py).

    python zagruzit_tovary_schetov.py --dry   # только протокол
    python zagruzit_tovary_schetov.py         # записать
"""
import datetime
import json
import os
import re
import sys

import openpyxl
import win32com.client

from sopostavit import Sopostavitel

sys.stdout.reconfigure(encoding='utf-8')
RAB = os.environ.get("TOVARY_RAB", r"C:\Razrabotki\tovary_schetov_kazna")
DRY = "--dry" in sys.argv
ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def clean(name):
    return re.sub(r"\s+", " ", re.sub(r"^[&*\-\s]+", "", name or "")).strip()[:500]


def main():
    erp = win32com.client.Dispatch("V83.COMConnector").Connect(ERP)
    S = erp.String
    sp = Sopostavitel(erp)
    sch = json.load(open(os.path.join(RAB, "scheta.json"), encoding="utf-8"))
    podr_cache, kontr_cache = {}, {}

    def podr(name):
        if name not in podr_cache:
            podr_cache[name] = erp.Справочники.СтруктураПредприятия.НайтиПоНаименованию(name, True)
        return podr_cache[name]

    def kontr(edrpou):
        if edrpou not in kontr_cache:
            z = erp.NewObject("Запрос")
            z.Text = 'ВЫБРАТЬ ПЕРВЫЕ 1 К.Ссылка ИЗ Справочник.Контрагенты КАК К ГДЕ К.КодПоЕДРПОУ = &К И НЕ К.ПометкаУдаления'
            z.SetParameter("К", edrpou)
            r = z.Execute().Выгрузить()
            kontr_cache[edrpou] = r.Получить(0).Получить(0) if r.Количество() and edrpou else None
        return kontr_cache[edrpou]

    prot = [("Счёт", "Дата", "Контрагент", "Подразделение", "Сумма счёта", "Файл", "Способ", "Строк", "Σ строк",
             "Заказ ЕРП", "Строк без номенклатуры", "Предупреждение")]
    total_rows = 0
    for s in sch:
        p = os.path.join(RAB, s["uid"], "stroki.json")
        r = json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {"stroki": [], "sposob": "нет файла", "file": "", "predupr": "файла нет"}
        zref, zuid, zwarn = sp.zakaz(s)
        rows, bez_nom = [], 0
        for i, x in enumerate(r["stroki"], 1):
            nom, src = sp.nomenklatura(x["nazv"], zuid, s["edrpou"])
            if nom is None:
                bez_nom += 1
            rows.append((i, x, nom))
        warn = "; ".join(w for w in (r.get("predupr", ""), zwarn) if w)
        ssum = round(sum(x["summa"] for _, x, _ in rows), 2)
        prot.append((s["nomer"], s["data_vh"] or s["data"], s["kontragent"], s["podr"], s["summa_scheta"], r.get("file", ""),
                     r.get("sposob", ""), len(rows), ssum, S(zref) if zref is not None else "", bez_nom, warn))
        if DRY:
            continue
        nabor = erp.РегистрыСведений.А_ТоварыСчетовПоставщиков.СоздатьНаборЗаписей()
        nabor.Отбор.СчетКазнаУИД.Установить(s["uid"])
        pd, kt = podr(s["podr"]), kontr(s["edrpou"])
        try:
            dt = datetime.datetime.strptime(s["data_vh"] or s["data"], "%d.%m.%Y")
        except ValueError:
            dt = datetime.datetime(2026, 1, 1)
        for i, x, nom in rows:
            z = nabor.Добавить()
            z.СчетКазнаУИД = s["uid"]
            z.НомерПозиции = i
            if nom is not None:
                z.Номенклатура = nom
            z.НазваниеВСчете = clean(x["nazv"])
            z.Количество = float(x.get("kol") or 0)
            z.Единица = (x.get("ed") or "")[:25]
            z.Цена = float(x.get("cena") or 0)
            z.Сумма = float(x.get("summa") or 0)
            if zref is not None:
                z.ЗаказПоставщику = zref
            z.НомерСчета = s["nomer"][:50]
            z.ДатаСчета = dt
            if kt is not None:
                z.Контрагент = kt
            if pd is not None:
                z.Подразделение = pd
            z.СуммаСчета = s["summa_scheta"]
            z.ИмяФайла = (r.get("file") or "")[:255]
            z.СпособРаспознавания = (r.get("sposob") or "")[:20]
            z.Предупреждение = ("; ".join(w for w in (warn, "" if nom is not None else "номенклатура не найдена") if w))[:500]
        nabor.Записать(True)
        total_rows += len(rows)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Протокол"
    for row in prot:
        ws.append(list(row))
    out = os.path.join(RAB, "protokol.xlsx")
    wb.save(out)
    ok = sum(1 for row in prot[1:] if row[7] and not row[11])
    print(f"счетов {len(prot) - 1}: с строками {sum(1 for x in prot[1:] if x[7])}, без предупреждений {ok}; "
          f"записано строк {total_rows}{' (DRY)' if DRY else ''}; протокол {out}")


if __name__ == "__main__":
    main()
