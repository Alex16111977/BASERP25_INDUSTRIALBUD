# -*- coding: utf-8 -*-
"""
03_scan_march2026.py — скан 3 типов документов за Март 2026:
  - АмортизацияОС
  - ПринятиеКУчетуОС
  - ВнутреннееПотреблениеТоваров
по нашей организации (ТОВ ИНДАСТРИАЛБУД).

Для каждого: Σ(ПАП.Приход) − Σ(ПАП.Расход) → если |Δ|>0.01, это асимметричный документ.

Выгрузка: _artifacts/03_scan_march_asymmetric.csv
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import connect_erp, save_csv, fail, money, get_org

DOC_TYPES = ["АмортизацияОС", "ПринятиеКУчетуОС", "ВнутреннееПотреблениеТоваров"]


def scan_doc_type(erp, org, doc_type):
    """Найти все документы данного типа с движениями в ПАП за Март 2026,
    показать Σ Дт vs Σ Кт по каждому."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", org)
    q.Текст = f"""
    ВЫБРАТЬ
        Д.Ссылка КАК Регистратор,
        Д.Номер КАК Номер,
        Д.Дата КАК Дата,
        Д.Проведен КАК Проведен,
        СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА Т.Сумма ИНАЧЕ 0 КОНЕЦ) КАК ΣДт,
        СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) ТОГДА Т.Сумма ИНАЧЕ 0 КОНЕЦ) КАК ΣКт
    ИЗ Документ.{doc_type} КАК Д
        ЛЕВОЕ СОЕДИНЕНИЕ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
        ПО Т.Регистратор = Д.Ссылка
    ГДЕ Д.Организация = &Орг
        И Д.Дата МЕЖДУ ДАТАВРЕМЯ(2026,3,1,0,0,0) И ДАТАВРЕМЯ(2026,3,31,23,59,59)
        И Д.Проведен
        И НЕ Д.ПометкаУдаления
    СГРУППИРОВАТЬ ПО Д.Ссылка, Д.Номер, Д.Дата, Д.Проведен
    """
    try:
        rez = q.Выполнить().Выгрузить()
    except Exception as e:
        return [], fail(e)
    rows = []
    for i in range(rez.Количество()):
        r = rez.Получить(i)
        dt = float(r.ΣДт or 0)
        kt = float(r.ΣКт or 0)
        delta = dt - kt
        rows.append({
            "ТипДокумента": doc_type,
            "Номер":        str(r.Номер).strip(),
            "Дата":         str(r.Дата),
            "ΣДт":          dt,
            "ΣКт":          kt,
            "Δ":            delta,
            "Расхождение":  "ДА" if abs(delta) > 0.01 else "",
        })
    return rows, None


def main():
    erp = connect_erp()
    org = get_org(erp)
    print(f"Организация: {erp.String(org)}")
    print(f"Период: Март 2026\n")

    all_rows = []
    for doc_type in DOC_TYPES:
        print(f"=== {doc_type} ===")
        rows, err = scan_doc_type(erp, org, doc_type)
        if err:
            print(f"  FAIL: {err}\n")
            continue
        if not rows:
            print(f"  (нет документов с движениями в ПАП)\n")
            continue

        asym_count = sum(1 for r in rows if r["Расхождение"])
        tot_delta = sum(r["Δ"] for r in rows)
        print(f"  Всего документов: {len(rows)}, асимметричных: {asym_count}, Σ|Δ|: {money(sum(abs(r['Δ']) for r in rows))}")
        print(f"  Σ ΣДт = {money(sum(r['ΣДт'] for r in rows))}, Σ ΣКт = {money(sum(r['ΣКт'] for r in rows))}, нетто Δ = {money(tot_delta)}")
        if asym_count > 0:
            print(f"  Топ-{min(10, asym_count)} асимметричных:")
            for r in sorted([x for x in rows if x["Расхождение"]], key=lambda x: -abs(x["Δ"]))[:10]:
                print(f"    {r['Номер']:<20} {r['Дата'][:10]:<10} Дт={money(r['ΣДт']):>15} Кт={money(r['ΣКт']):>15} Δ={money(r['Δ']):>13}")
        print()
        all_rows.extend(rows)

    save_csv("03_scan_march_all", all_rows,
             ["ТипДокумента","Номер","Дата","ΣДт","ΣКт","Δ","Расхождение"])
    print(f"\n→ _artifacts/03_scan_march_all.csv ({len(all_rows)} строк)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
