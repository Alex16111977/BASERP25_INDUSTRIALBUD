# -*- coding: utf-8 -*-
"""Формирование встроенного отчёта «Прогноз виконання СС за потребою» путём формы БСП
(ВариантыОтчетов.СформироватьОтчет, дата — СтандартнаяДатаНачала, отбор по подразделению) и сохранение в xlsx.

    python vygruzka_xlsx.py --date 2026-10-08 --podr "МД IRC 2026" --variant ПрогнозЗаПотребою --out <файл.xlsx>
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
PROJ = json.loads(Path(r"C:\Configuration_downloads\BASERP25\.v8-project.json").read_text(encoding="utf-8-sig"))
REPORT = "А_ПрогнозВиконанняССЗаПотребою"


def connect():
    db = next(d for d in PROJ["databases"] if d["id"] == "erp")
    return win32com.client.Dispatch("V83.COMConnector").Connect(
        f'Srvr="localhost";Ref="{db["ref"]}";Usr="{db["user"]}";Pwd="{db["password"]}"')


def build(erp, date, podr_name, variant):
    S = erp.String
    podr = next(p for p in (erp.Документы.А_ПотребаЗавершенняПроекту.ПодразделенияАктивныхСС().Получить(i)
                            for i in range(erp.Документы.А_ПотребаЗавершенняПроекту.ПодразделенияАктивныхСС().Количество()))
                if S(p).strip() == podr_name)
    rep = getattr(erp.Отчеты, REPORT).Создать()
    schema = rep.СхемаКомпоновкиДанных
    k = rep.КомпоновщикНастроек
    k.Инициализировать(erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema))
    k.ЗагрузитьНастройки(schema.ВариантыНастроек.Найти(variant).Настройки)
    set_date = set_podr = False
    podr_ids = set()
    for j in range(k.Настройки.Отбор.Элементы.Количество()):
        it = k.Настройки.Отбор.Элементы.Получить(j)
        if S(it.ЛевоеЗначение) == "Подразделение":
            podr_ids.add(S(it.ИдентификаторПользовательскойНастройки))
    for i in range(k.ПользовательскиеНастройки.Элементы.Количество()):
        el = k.ПользовательскиеНастройки.Элементы.Получить(i)
        if S(el.ИдентификаторПользовательскойНастройки) in podr_ids:
            el.Использование = True
            el.ПравоеЗначение = podr
            set_podr = True
            continue
        try:
            if S(el.Параметр) == "ДатаОтчета":
                el.Использование = True
                el.Значение = erp.NewObject("СтандартнаяДатаНачала", date + dt.timedelta(hours=3))
                set_date = True
                continue
        except Exception:  # noqa: BLE001 — элемент отбора, а не параметр
            pass
        try:
            if S(el.ЛевоеЗначение) == "Подразделение":
                el.Использование = True
                el.ПравоеЗначение = podr
                set_podr = True
        except Exception:  # noqa: BLE001
            pass
    if not (set_date and set_podr):
        raise SystemExit(f"быстрые настройки не найдены: дата {set_date}, подразделение {set_podr}")
    p = erp.NewObject("Структура")
    p.Вставить("СсылкаОтчета", erp.ОбщегоНазначения.ИдентификаторОбъектаМетаданных("Отчет." + REPORT))
    p.Вставить("СсылкаВарианта", None)
    p.Вставить("КлючВарианта", variant)
    p.Вставить("НастройкиКД", k.Настройки)
    p.Вставить("ФиксированныеНастройкиКД", k.ФиксированныеНастройки)
    p.Вставить("ПользовательскиеНастройкиКД", k.ПользовательскиеНастройки)
    p.Вставить("СхемаМодифицирована", False)
    p.Вставить("КлючСхемы", "")
    p.Вставить("ИмяКлючевойОперации")
    p.Вставить("КомментарийКлючевойОперации")
    p.Вставить("Объект", rep)
    p.Вставить("ПолноеИмя", "Отчет." + REPORT)
    res = erp.ВариантыОтчетов.СформироватьОтчет(p, False, False)
    if not res.Успех:
        raise SystemExit("ОШИБКА БСП: " + S(res.ТекстОшибки))
    return res.ТабличныйДокумент


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", required=True)
    ap.add_argument("--podr", required=True)
    ap.add_argument("--variant", default="ПрогнозЗаПотребою")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    erp = connect()
    td = build(erp, dt.datetime.strptime(a.date, "%Y-%m-%d").replace(hour=12), a.podr, a.variant)
    print(f"{a.podr} / {a.variant}: таблица {td.ВысотаТаблицы}×{td.ШиринаТаблицы}")
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    # Записать выполняет серверный процесс 1С: пишем сразу в целевую папку (временная папка пользователя ему недоступна)
    td.Записать(str(out), erp.ТипФайлаТабличногоДокумента.XLSX)
    print("сохранено:", out, out.stat().st_size, "байт")


if __name__ == "__main__":
    main()
