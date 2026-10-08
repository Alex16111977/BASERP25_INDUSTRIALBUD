# -*- coding: utf-8 -*-
"""Загрузка файла потребы директора в документы А_ПотребаЗавершенняПроекту (BaseERP) тем же кодом, что кнопка формы.

    python zagruzit_potrebu.py <файл.xlsx> "<раздел>=<подразделение>" [...] --moment "2026-10-08 11:36:56" [--apply]

Без --apply: только чтение файла и сверка с Python-эталоном (parse_file_etalon.py).
С --apply: get-or-create документа по (Подразделение, РазделФайла, день даты), заполнение, запись с проведением.
Документы не удаляются.
"""
import argparse
import datetime as dt
import json
import shutil
import sys
import tempfile
from pathlib import Path

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).parent))
from parse_file_etalon import parse  # noqa: E402

PROJ = json.loads(Path(r"C:\Configuration_downloads\BASERP25\.v8-project.json").read_text(encoding="utf-8-sig"))
COM_SHIFT = dt.timedelta(hours=3)  # память com_date_shift_3_hours


def connect():
    db = next(d for d in PROJ["databases"] if d["id"] == "erp")
    return win32com.client.Dispatch("V83.COMConnector").Connect(
        f'Srvr="localhost";Ref="{db["ref"]}";Usr="{db["user"]}";Pwd="{db["password"]}"')


def arr(com_array):
    return [com_array.Получить(i) for i in range(com_array.Количество())]


def section_rows(erp, razdel):
    out = []
    for r in arr(razdel.Строки):
        out.append({k: (erp.String(getattr(r, k)) if k in ("НазваниеТовара", "НазваниеСС", "Единица", "Дома", "Предупреждение")
                        else float(getattr(r, k)))
                    for k in ("НомерСтрокиФайла", "НазваниеТовара", "НазваниеСС", "КолФайла", "Единица", "КоличествоСС",
                              "Цена", "Сумма", "Дома", "Предупреждение")})
    return out


def compare(etalon, erp_secs):
    fails = []
    if [s["name"] for s in etalon] != [s["name"] for s in erp_secs]:
        fails.append(f"разделы: эталон {[s['name'] for s in etalon]} ≠ 1С {[s['name'] for s in erp_secs]}")
        return fails
    for e, c in zip(etalon, erp_secs):
        if len(e["rows"]) != len(c["rows"]):
            fails.append(f"{e['name']}: строк {len(e['rows'])} ≠ {len(c['rows'])}")
            continue
        if abs(e["sum"] - c["sum"]) > 0.005:
            fails.append(f"{e['name']}: Σ {e['sum']} ≠ {c['sum']}")
        if (e["itogo"] is None) != (c["itogo"] is None) or (e["itogo"] and abs(e["itogo"] - c["itogo"]) > 0.01):
            fails.append(f"{e['name']}: ИТОГО {e['itogo']} ≠ {c['itogo']}")
        for er, cr in zip(e["rows"], c["rows"]):
            for k_e, k_c in (("row", "НомерСтрокиФайла"), ("Сумма", "Сумма"), ("Цена", "Цена"), ("КоличествоСС", "КоличествоСС"),
                             ("Количество", "КолФайла"), ("НазваниеСС", "НазваниеСС"), ("НазваниеТовара", "НазваниеТовара"),
                             ("Единица", "Единица"), ("Дома", "Дома")):
                a, b = er[k_e], cr[k_c]
                same = abs(float(a) - float(b)) <= 0.0005 if isinstance(a, (int, float)) else str(a) == str(b)
                if not same:
                    fails.append(f"{e['name']} r{er['row']} {k_c}: эталон {a!r} ≠ 1С {b!r}")
    return fails


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("pairs", nargs="+", help="<маркер раздела>=<подразделение>")
    ap.add_argument("--moment", required=True, help="момент составления списка, местное время ГГГГ-ММ-ДД ЧЧ:ММ:СС")
    ap.add_argument("--comment", default="")
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    copy = Path(tempfile.mkdtemp()) / Path(a.file).name
    shutil.copy(a.file, copy)
    etalon = parse(copy)

    erp = connect()
    adres = erp.ПоместитьВоВременноеХранилище(erp.NewObject("ДвоичныеДанные", str(copy)))
    res = erp.Документы.А_ПотребаЗавершенняПроекту.ПрочитатьФайлПотребы(adres, copy.name)
    if erp.String(res.Ошибка):
        raise SystemExit("ОШИБКА чтения: " + erp.String(res.Ошибка))
    razdely = arr(res.Разделы)
    erp_secs = [{"name": erp.String(r.Имя), "rows": section_rows(erp, r), "sum": round(float(r.Сумма), 2),
                 "itogo": None if r.СуммаИтога is None else round(float(r.СуммаИтога), 2), "err": erp.String(r.Ошибка)}
                for r in razdely]
    for s in erp_secs:
        print(f"1С: {s['name']!r}: строк {len(s['rows'])}, Σ {s['sum']:.2f}, ИТОГО {s['itogo']}, ошибка «{s['err']}»")
    fails = compare(etalon, erp_secs)
    print("СВЕРКА С ЭТАЛОНОМ:", "OK" if not fails else f"{len(fails)} расхождений")
    for f in fails:
        print("  ", f)
    if fails or not a.apply:
        sys.exit(1 if fails else 0)

    moment = dt.datetime.strptime(a.moment, "%Y-%m-%d %H:%M:%S")
    podr_all = {erp.String(p).strip(): p for p in arr(erp.Документы.А_ПотребаЗавершенняПроекту.ПодразделенияАктивныхСС())}
    for pair in a.pairs:
        marker, podr_name = pair.split("=", 1)
        idx = next(i for i, s in enumerate(erp_secs) if s["name"] == marker)
        podr = podr_all[podr_name]
        q = erp.NewObject("Запрос")
        q.Текст = """ВЫБРАТЬ ДокПотр.Ссылка КАК Ссылка ИЗ Документ.А_ПотребаЗавершенняПроекту КАК ДокПотр
            ГДЕ ДокПотр.Подразделение = &Подр И ДокПотр.РазделФайла = &Раздел
            И ДокПотр.Дата МЕЖДУ НАЧАЛОПЕРИОДА(&Д, ДЕНЬ) И КОНЕЦПЕРИОДА(&Д, ДЕНЬ)"""
        q.УстановитьПараметр("Подр", podr)
        q.УстановитьПараметр("Раздел", marker)
        q.УстановитьПараметр("Д", moment + COM_SHIFT)
        found = q.Выполнить().Выгрузить()
        if found.Количество() > 0:
            obj = found.Получить(0).Ссылка.ПолучитьОбъект()
            print(f"{podr_name}: обновляю {erp.String(obj.Номер)}")
        else:
            obj = erp.Документы.А_ПотребаЗавершенняПроекту.СоздатьДокумент()
            obj.Заполнить(None)
            print(f"{podr_name}: новый документ")
        obj.Дата = moment + COM_SHIFT
        obj.Подразделение = podr
        if a.comment:
            obj.Комментарий = a.comment
        itog = erp.Документы.А_ПотребаЗавершенняПроекту.ЗаполнитьПотребу(obj, razdely[idx], False)
        if erp.String(itog.Ошибка):
            raise SystemExit(f"{podr_name}: {erp.String(itog.Ошибка)}")
        obj.Записать(erp.РежимЗаписиДокумента.Проведение)
        print(f"  {erp.String(obj.Ссылка)} | дата {erp.String(obj.Дата)} | строк {int(itog.Строк)} | Σ {float(obj.СуммаДокумента):,.2f}"
              f" | без ОН {int(itog.НеНайдено)} | с предупреждениями {int(itog.Предупреждений)}".replace(",", " "))
        for st in arr(obj.Потреба):
            if erp.String(st.Предупреждение):
                print(f"    r{int(st.НомерСтрокиФайла)} {erp.String(st.НазваниеСС)} → «{erp.String(st.ОбщееНазвание)}»: {erp.String(st.Предупреждение)}")


if __name__ == "__main__":
    main()
