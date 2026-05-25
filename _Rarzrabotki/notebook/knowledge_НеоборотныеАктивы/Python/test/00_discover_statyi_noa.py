# -*- coding: utf-8 -*-
"""
00_discover_statyi_noa.py — поиск всех статей АктивовПассивов,
имеющих отношение к НОА (Основные средства / НМА / Расходы /
Вложения в необоротные активы).

Выгрузка: _artifacts/00_statyi_noa.csv (Код, Имя, АктивПассив).
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import connect_erp, save_csv, fail

KEYWORDS = [
    "сновные средства",
    "ематериальн",
    "асходы текущ",
    "асходы будущ",
    "ложения",
    "еоборотн",
    "езерв",
    "Налог",
    "ДооценкаОС",
    "Капитальные",
]

def main():
    erp = connect_erp()
    q = erp.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ
        С.Ссылка,
        С.Код,
        С.Наименование,
        ПРЕДСТАВЛЕНИЕ(С.АктивПассив) КАК АктивПассив,
        С.ПометкаУдаления
    ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов КАК С
    УПОРЯДОЧИТЬ ПО С.Наименование
    """
    try:
        rez = q.Выполнить().Выгрузить()
    except Exception as e:
        print(f"FAIL запрос: {fail(e)}")
        return 1

    out = []
    for i in range(rez.Количество()):
        r = rez.Получить(i)
        nazv = str(r.Наименование)
        nazv_l = nazv.lower()
        if not any(k.lower() in nazv_l for k in KEYWORDS):
            continue
        out.append({
            "Код": str(r.Код).strip(),
            "Наименование": nazv,
            "АктивПассив": str(r.АктивПассив),
            "Удалена": "Да" if r.ПометкаУдаления else "",
        })

    save_csv("00_statyi_noa", out, ["Код", "Наименование", "АктивПассив", "Удалена"])
    print(f"OK, найдено статей: {len(out)}")
    for r in out:
        print(f"  [{r['Код']:>4}] {r['АктивПассив']:>10} | {r['Наименование']:<60}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
