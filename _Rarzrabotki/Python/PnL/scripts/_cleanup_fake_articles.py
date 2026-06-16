# -*- coding: utf-8 -*-
"""Пометить на удаление НЕ-locked фейк-статьи Справочник.А_Статьи_PL.

Берёт незащищённые (А_РучнаяКорректировка=Ложь) элементы и помечает на удаление те,
чьё имя распознаётся реестром фейков (utils.fake_articles.match_fake). Защищённые
(ручные) статьи НЕ трогает. Запись через ОбменДанными.Загрузка=True (обход гардов ПередЗаписью).

Запуск:
  python scripts/_cleanup_fake_articles.py --dry-run
  python scripts/_cleanup_fake_articles.py
"""
import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import win32com.client
import pythoncom
import config
from utils import fake_articles


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    pythoncom.CoInitialize()
    conn = win32com.client.Dispatch("V83.COMConnector").Connect(config.CONN_ERP)

    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ С.Ссылка КАК Ссылка, С.Наименование КАК Имя "
        "ИЗ Справочник.А_Статьи_PL КАК С "
        "ГДЕ НЕ С.ПометкаУдаления И НЕ С.ЭтоГруппа И НЕ С.А_РучнаяКорректировка"
    )
    sel = q.Выполнить().Выбрать()

    out = []
    marked = 0
    while sel.Следующий():
        nm = sel.Имя
        reason = fake_articles.match_fake(nm)
        if reason:
            marked += 1
            tag = "DRY" if args.dry_run else "DEL"
            out.append(f"  [{tag}] {nm}  <= {reason}")
            if not args.dry_run:
                obj = sel.Ссылка.ПолучитьОбъект()
                obj.ОбменДанными.Загрузка = True
                obj.ПометкаУдаления = True
                obj.Записать()
        else:
            out.append(f"  [KEEP non-locked, NOT fake] {nm}")

    header = "Фейк-статьи на удаление" + (" (DRY-RUN)" if args.dry_run else " (помечено)")
    res = [header] + out + [f"\nИтого помечено: {marked}"]
    Path(__file__).with_name("_cleanup_fake_articles.out").write_text("\n".join(res), encoding="utf-8")
    print("done; marked:", marked, "dry_run:", args.dry_run)


if __name__ == "__main__":
    main()
