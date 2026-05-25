# -*- coding: utf-8 -*-
"""
01_drill_amortizatsiya_000C_000003.py — drill проблемного документа
"Амортизация ОС (2.1) 000Ц-000003 от 31.03.2026" (UUID 2d9182ad-3062-11f1-a2e7-eb06fbf9b98b).

Источник плуга в Контроле баланса:
  Дт «Расходы текущего периода» = 241 388.07
  Кт «Основные средства»        = 139 327.52
  Разница                       = 102 060.55

Цель:
  1. Достать ВСЕ движения этого документа во ВСЕХ регистрах накопления,
     которые он двигает (особенно РН.ПрочиеАктивыПассивы).
  2. Сложить Σ(signed) по (Статья, Подразделение, Аналитика) внутри ПАП
     — это покажет реальную асимметрию.
  3. Объяснить откуда 102 060.55.

Выгрузка:
  _artifacts/01_drill_movements.csv — все движения по регистрам
  _artifacts/01_pap_summary.csv     — свод по ПАП (Статья × Подразделение × signed)
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import connect_erp, save_csv, fail, money, get_type_name

DOC_UUID = "2d9182ad-3062-11f1-a2e7-eb06fbf9b98b"


def get_doc_ref(erp):
    uid = erp.NewObject("УникальныйИдентификатор", DOC_UUID)
    return erp.Документы.АмортизацияОС.ПолучитьСсылку(uid)


def list_all_movements(erp, doc_ref):
    """Вернуть ВСЕ регистры накопления, которые двигает этот документ."""
    # Через метаданные документа узнаем какие регистры он может двигать
    md = doc_ref.Метаданные()
    res = []
    for rec in md.Движения:
        reg_name = rec.Имя
        try:
            rs = erp.РегистрыНакопления[reg_name].СоздатьНаборЗаписей()
            rs.Отбор.Регистратор.Установить(doc_ref)
            rs.Прочитать()
            kol = rs.Количество()
            if kol > 0:
                res.append((reg_name, kol))
        except Exception:
            continue
    return res


def dump_register_movements(erp, doc_ref, reg_name):
    """Выгрузить все строки движений из заданного РН."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Р", doc_ref)
    q.Текст = f"""
    ВЫБРАТЬ
        Т.НомерСтроки КАК НомерСтроки,
        Т.Период      КАК Период,
        ПРЕДСТАВЛЕНИЕ(Т.ВидДвижения) КАК ВидДвижения,
        ПРЕДСТАВЛЕНИЕ(Т.Организация) КАК Организация,
        ПРЕДСТАВЛЕНИЕ(Т.Подразделение) КАК Подразделение,
        ПРЕДСТАВЛЕНИЕ(Т.Статья) КАК Статья,
        Т.Статья.Код  КАК Код,
        ПРЕДСТАВЛЕНИЕ(Т.Аналитика) КАК Аналитика,
        ПРЕДСТАВЛЕНИЕ(Т.Источник) КАК Источник,
        Т.Сумма       КАК Сумма
    ИЗ РегистрНакопления.{reg_name} КАК Т
    ГДЕ Т.Регистратор = &Р
    УПОРЯДОЧИТЬ ПО Т.НомерСтроки
    """
    try:
        rez = q.Выполнить().Выгрузить()
    except Exception as e:
        return None, fail(e)
    rows = []
    for i in range(rez.Количество()):
        r = rez.Получить(i)
        rows.append({
            "Регистр": reg_name,
            "НомерСтроки": int(r.НомерСтроки) if r.НомерСтроки else 0,
            "ВидДвижения": str(r.ВидДвижения),
            "Подразделение": str(r.Подразделение),
            "Статья": str(r.Статья) if reg_name == "ПрочиеАктивыПассивы" else "",
            "Код": str(r.Код).strip() if reg_name == "ПрочиеАктивыПассивы" else "",
            "Аналитика": str(r.Аналитика) if reg_name == "ПрочиеАктивыПассивы" else "",
            "Источник": str(r.Источник) if reg_name == "ПрочиеАктивыПассивы" else "",
            "Сумма": float(r.Сумма) if r.Сумма else 0.0,
        })
    return rows, None


def main():
    erp = connect_erp()
    doc_ref = get_doc_ref(erp)
    if not erp.ЗначениеЗаполнено(doc_ref):
        print(f"FAIL: документ {DOC_UUID} не найден")
        return 1
    print(f"Документ: {erp.String(doc_ref)}")

    # 1) Все регистры с движениями
    print("\n=== Регистры с движениями ===")
    regs = list_all_movements(erp, doc_ref)
    for r, k in regs:
        print(f"  {r:50} | {k:>4} строк")

    # 2) Выгружаем ПрочиеАктивыПассивы детально + остальные коротко
    print("\n=== РН.ПрочиеАктивыПассивы — построчно ===")
    pap_rows, err = dump_register_movements(erp, doc_ref, "ПрочиеАктивыПассивы")
    if err:
        print(f"FAIL ПАП: {err}")
        return 2
    if pap_rows is None or len(pap_rows) == 0:
        print("(нет движений в РН.ПрочиеАктивыПассивы)")
    else:
        for r in pap_rows:
            print(f"  #{r['НомерСтроки']:>3} {r['ВидДвижения']:>7} {r['Подразделение']:<30} "
                  f"[{r['Код']:>4}] {r['Статья']:<40} "
                  f"Анал={r['Аналитика']:<40} Ист={r['Источник']:<25} "
                  f"Сумма={money(r['Сумма']):>16}")

    # 3) Свод по ПАП: Статья × signed
    print("\n=== Свод по ПАП (Статья × signed) ===")
    agg = {}
    for r in pap_rows or []:
        sign = 1 if "Приход" in r["ВидДвижения"] else -1
        key = (r["Код"], r["Статья"], r["Подразделение"])
        agg[key] = agg.get(key, 0) + sign * r["Сумма"]
    summary = []
    for (kod, statya, podr), s in sorted(agg.items()):
        summary.append({
            "Код": kod, "Статья": statya, "Подразделение": podr,
            "Σ_signed": s,
            "Σ_signed_pretty": money(s),
        })
        print(f"  [{kod:>4}] {statya:<35} {podr:<30} signed={money(s):>16}")

    # 4) Полный CSV-дамп всех регистров
    print("\n=== Полный CSV-дамп ===")
    all_rows = list(pap_rows or [])
    for reg_name, _ in regs:
        if reg_name == "ПрочиеАктивыПассивы":
            continue
        other_rows, err = dump_register_movements(erp, doc_ref, reg_name)
        if err:
            print(f"  [WARN] {reg_name}: {err}")
            continue
        all_rows.extend(other_rows or [])
    save_csv("01_drill_movements", all_rows,
             ["Регистр","НомерСтроки","ВидДвижения","Подразделение",
              "Статья","Код","Аналитика","Источник","Сумма"])
    save_csv("01_pap_summary", summary,
             ["Код","Статья","Подразделение","Σ_signed","Σ_signed_pretty"])
    print(f"  → _artifacts/01_drill_movements.csv ({len(all_rows)} строк)")
    print(f"  → _artifacts/01_pap_summary.csv ({len(summary)} строк)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
