# -*- coding: utf-8 -*-
"""Перенос документов А_ПотребаЗавершенняПроекту из BaseERPRazr (источник, только чтение) в BaseERP (приёмник).

Промт: _Rarzrabotki/notebook/specs/2026-10-08-perenos-potreba-zavershennya-v-rabochuyu-bazu-PROMT.md
    python perenos.py            # пробный прогон
    python perenos.py --apply    # запись (XDTO-сериализация: UUID, номер, дата с временем, флаги — как в источнике)
    python perenos.py --rollback=<backup_target.xml>   # вернуть приёмник: созданные документы пометить на удаление
"""
import argparse
import csv
import datetime as dt
import json
import re
import sys
from pathlib import Path

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).parent
PROJ = json.loads(Path(r"C:\Configuration_downloads\BASERP25\.v8-project.json").read_text(encoding="utf-8-sig"))
DOC = "А_ПотребаЗавершенняПроекту"
REFS_HEAD = {"Подразделение": "СтруктураПредприятия", "Ответственный": "Пользователи"}
REFS_ROWS = {"ОбщееНазвание": "А_ОбщиеНазванияНоменклатуры", "ЕдиницаСС": "УпаковкиЕдиницыИзмерения"}


def connect(ref):
    db = next(d for d in PROJ["databases"] if d["id"] == "erp")
    return win32com.client.Dispatch("V83.COMConnector").Connect(
        f'Srvr="localhost";Ref="{ref}";Usr="{db["user"]}";Pwd="{db["password"]}"')


def docs(base):
    q = base.NewObject("Запрос")
    q.Текст = f"ВЫБРАТЬ Д.Ссылка КАК Ссылка ИЗ Документ.{DOC} КАК Д УПОРЯДОЧИТЬ ПО Д.МоментВремени"
    t = q.Выполнить().Выгрузить()
    return [t.Получить(i).Ссылка for i in range(t.Количество())]


def to_xml(base, obj):
    w = base.NewObject("ЗаписьXML")
    w.УстановитьСтроку()
    base.СериализаторXDTO.ЗаписатьXML(w, obj)
    return w.Закрыть()


def from_xml(base, text):
    r = base.NewObject("ЧтениеXML")
    r.УстановитьСтроку(text)
    return base.СериализаторXDTO.ПрочитатьXML(r)


def date_str(base, ref):
    """Дата документа компонентами из запроса — без COM-конвертации дат (сдвиг −3 ч)."""
    q = base.NewObject("Запрос")
    q.Текст = (f"ВЫБРАТЬ ГОД(Д.Дата) КАК Г, МЕСЯЦ(Д.Дата) КАК М, ДЕНЬ(Д.Дата) КАК Дн, ЧАС(Д.Дата) КАК Ч, "
               f"МИНУТА(Д.Дата) КАК Мин, СЕКУНДА(Д.Дата) КАК С ИЗ Документ.{DOC} КАК Д ГДЕ Д.Ссылка = &Ссылка")
    q.УстановитьПараметр("Ссылка", ref)
    r = q.Выполнить().Выгрузить().Получить(0)
    return f"{int(r.Дн):02d}.{int(r.М):02d}.{int(r.Г)} {int(r.Ч):02d}:{int(r.Мин):02d}:{int(r.С):02d}"


def summary(base, ref):
    o = ref.ПолучитьОбъект()
    S = base.String
    rows = [o.Потреба.Получить(i) for i in range(o.Потреба.Количество())]
    return {"uuid": S(ref.УникальныйИдентификатор()), "Номер": S(o.Номер),
            "Дата": date_str(base, ref), "Подразделение": S(o.Подразделение).strip(),
            "Проведен": bool(o.Проведен), "ПометкаУдаления": bool(o.ПометкаУдаления),
            "СуммаДокумента": round(float(o.СуммаДокумента), 2), "Строк": len(rows),
            "СуммаСтрок": round(sum(float(r.Сумма) for r in rows), 2)}


def target_ref(dst, catalog, uuid, name):
    """(ссылка приёмника, способ) по UUID, иначе по наименованию (однозначно); (None, причина)."""
    mgr = getattr(dst.Справочники, catalog)
    ref = mgr.ПолучитьСсылку(dst.NewObject("УникальныйИдентификатор", uuid))
    if ref.ПолучитьОбъект() is not None:
        return ref, "UUID"
    q = dst.NewObject("Запрос")
    q.Текст = f"ВЫБРАТЬ С.Ссылка КАК Ссылка ИЗ Справочник.{catalog} КАК С ГДЕ С.Наименование = &Имя И НЕ С.ПометкаУдаления"
    q.УстановитьПараметр("Имя", name)
    t = q.Выполнить().Выгрузить()
    if t.Количество() == 1:
        return t.Получить(0).Ссылка, "наименование"
    return None, f"не найдено ({t.Количество()} по наименованию)"


def check_refs(src, dst, ref):
    """Список (поле, uuid источника, uuid приёмника, способ) и список проблем."""
    o = ref.ПолучитьОбъект()
    S = src.String
    seen, mapping, problems = {}, [], []
    items = [(f, cat, getattr(o, f)) for f, cat in REFS_HEAD.items()]
    for i in range(o.Потреба.Количество()):
        r = o.Потреба.Получить(i)
        items += [(f"Потреба[{i + 1}].{f}", cat, getattr(r, f)) for f, cat in REFS_ROWS.items()]
    for field, cat, val in items:
        if val is None or not src.ЗначениеЗаполнено(val):
            continue
        uuid = S(val.УникальныйИдентификатор())
        if uuid in seen:
            continue
        tref, how = target_ref(dst, cat, uuid, S(val.Наименование))
        seen[uuid] = (tref, how)
        if tref is None:
            problems.append(f"{field}: {cat} «{S(val)}» {uuid} — {how}")
        else:
            mapping.append((cat, uuid, dst.String(tref.УникальныйИдентификатор()), how))
    return mapping, problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--rollback")
    a = ap.parse_args()
    src, dst = connect("BaseERPRazr"), connect("BaseERP")
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")

    if a.rollback:
        created = json.loads(Path(a.rollback).with_suffix(".json").read_text(encoding="utf-8")).get("created", [])
        for uuid in created:
            ref = getattr(dst.Документы, DOC).ПолучитьСсылку(dst.NewObject("УникальныйИдентификатор", uuid))
            o = ref.ПолучитьОбъект()
            if o is not None:
                o.ОбменДанными.Загрузка = True
                o.УстановитьПометкуУдаления(True)
                print("помечен на удаление", uuid)
        return

    src_docs, dst_docs = docs(src), docs(dst)
    dst_uuids = {dst.String(r.УникальныйИдентификатор()): r for r in dst_docs}
    # Шаг 1. Бэкапы
    (HERE / f"backup_source_{stamp}.xml").write_text(
        "\n".join(to_xml(src, r.ПолучитьОбъект()) for r in src_docs), encoding="utf-8")
    (HERE / f"backup_source_{stamp}.json").write_text(
        json.dumps([summary(src, r) for r in src_docs], ensure_ascii=False, indent=1), encoding="utf-8")
    bt = HERE / f"backup_target_{stamp}.xml"
    bt.write_text("\n".join(to_xml(dst, r.ПолучитьОбъект()) for r in dst_docs), encoding="utf-8")
    print(f"бэкапы: источник {len(src_docs)} док., приёмник {len(dst_docs)} док. → {HERE}")

    # Шаги 2–3. Ссылки и пробный прогон
    report, plan, skipped = [], [], 0
    for r in src_docs:
        sm = summary(src, r)
        mapping, problems = check_refs(src, dst, r)
        by_name = sum(1 for m in mapping if m[3] != "UUID")
        exists = sm["uuid"] in dst_uuids
        if problems:
            action = "пропуск"
            skipped += 1
        elif exists and to_xml(dst, dst_uuids[sm["uuid"]].ПолучитьОбъект()) == to_xml(src, r.ПолучитьОбъект()):
            action = "без изменений"
        else:
            action = "обновить" if exists else "создать"
            plan.append((r, mapping, action))
        print(f"№ {sm['Номер']} от {sm['Дата']} | {sm['Подразделение']} | Σ {sm['СуммаДокумента']:,.2f} | строк {sm['Строк']} | "
              f"ссылок {len(mapping)} (по наименованию {by_name}) | в приёмнике: {'есть' if exists else 'нет'} → {action}".replace(",", " "))
        for p in problems:
            print("    НЕ НАЙДЕНО:", p)
        report.append({**sm, "ссылок": len(mapping), "по_наименованию": by_name, "проблем": len(problems), "действие": action})
    with open(HERE / "perenos_report.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(report[0].keys()), delimiter=";")
        w.writeheader()
        w.writerows(report)
    print(f"ИТОГО: к записи {len(plan)}, пропуск {skipped}, без изменений {len(src_docs) - len(plan) - skipped}")
    if not a.apply:
        return
    if skipped:
        raise SystemExit("есть пропуски — запись остановлена, нужно решение пользователя")

    # Шаг 4. Запись через XDTO
    created = []
    for r, mapping, action in plan:
        text = to_xml(src, r.ПолучитьОбъект())
        for _cat, u_src, u_dst, how in mapping:
            if u_src != u_dst:
                text = text.replace(u_src, u_dst)
        obj = from_xml(dst, text)
        obj.ОбменДанными.Загрузка = True
        obj.Записать()
        if action == "создать":
            created.append(src.String(r.УникальныйИдентификатор()))
        print(f"  записан № {dst.String(obj.Номер)} ({action})")
    bt.with_suffix(".json").write_text(json.dumps({"created": created}, ensure_ascii=False), encoding="utf-8")

    # Шаг 5. Проверка
    fails = 0
    dst_docs2 = {dst.String(x.УникальныйИдентификатор()): x for x in docs(dst)}
    if len(dst_docs2) != len(dst_docs) + len(created):
        fails += 1
        print(f"FAIL число документов {len(dst_docs2)} ≠ {len(dst_docs)} + {len(created)}")
    for r in src_docs:
        u = src.String(r.УникальныйИдентификатор())
        t = dst_docs2.get(u)
        if t is None:
            fails += 1
            print("FAIL нет в приёмнике", u)
            continue
        s_sum, d_sum = summary(src, r), summary(dst, t)
        same_xml = to_xml(src, r.ПолучитьОбъект()) == to_xml(dst, t.ПолучитьОбъект())
        ok = s_sum == d_sum and same_xml
        fails += 0 if ok else 1
        print(f"{'OK  ' if ok else 'FAIL'} № {d_sum['Номер']}: сводка {'=' if s_sum == d_sum else '≠'}, XML {'=' if same_xml else '≠'}")
        if s_sum != d_sum:
            print("     источник", s_sum, "\n     приёмник", d_sum)
    print("ПРОВЕРКА:", "OK" if not fails else f"{fails} провалов")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
