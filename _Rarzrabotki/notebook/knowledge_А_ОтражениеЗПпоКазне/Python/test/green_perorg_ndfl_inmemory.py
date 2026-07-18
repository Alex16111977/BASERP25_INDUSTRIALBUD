"""
GREEN-харнесс per-org НДФЛ/«Удержание» (ТЧ-уровень, in-memory, БЕЗ записи).

Заполняет А_ОтражениеЗПпоКазне в памяти (оркестратор ЗаполнитьДляОтражения...) и проверяет:
  1. НачисленныйНДФЛ Ф1: ОрганизацияБухгалтерия НЕ пустая (GREEN=0 пустых);
  2. gross-up «Удержание» Ф1 (А_ЭтоУдержание, ВзносыВсего=0): ОргБух НЕ пустая;
  3. per (ФЛ, Орг): Σ НачисленныйНДФЛ.Сумма == Σ НалогиБухгалтерия.СуммаНДФЛ того же дока (точный per-org сплит);
  4. «Удержание» per (ФЛ, Орг) зеркалит НДФЛ per (ФЛ, Орг).

ДО правок+загрузки кода: RED (ОргБух у НДФЛ пустая). ПОСЛЕ: GREEN.

Запуск: C:\\Python313\\python.exe green_perorg_ndfl_inmemory.py [НОМЕР]
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

НОМЕР = sys.argv[1] if len(sys.argv) > 1 else "000000001"


def nm(v, erp):
    return v.Наименование if erp.ЗначениеЗаполнено(v) else "<ПУСТО>"


def main():
    erp = win32com.client.Dispatch("V83.COMConnector").Connect(
        'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
    )
    Ф1 = erp.Перечисления.А_ФормыPL.Форма1
    Ф1x = erp.XMLСтрока(Ф1)

    ref = erp.Документы.А_ОтражениеЗПпоКазне.НайтиПоНомеру(НОМЕР)
    if not erp.ЗначениеЗаполнено(ref):
        print(f"FAIL: документ №{НОМЕР} не найден")
        return
    об = ref.ПолучитьОбъект()
    print(f"Документ №{НОМЕР} от {об.Дата} | Организация={nm(об.Организация, erp)}")

    # Источник истины: НалогиБухгалтерия per (ФЛ, Орг) — ДО заполнения (исходная ТЧ дока)
    нб = {}
    for i in range(об.НалогиБухгалтерия.Количество()):
        s = об.НалогиБухгалтерия.Получить(i)
        if s.СуммаНДФЛ == 0:
            continue
        фл = s.Сотрудник.ФизическоеЛицо if erp.ЗначениеЗаполнено(s.Сотрудник) else None
        флн = nm(фл, erp)
        орг = nm(s.Организация, erp)
        нб[(флн, орг)] = round(нб.get((флн, орг), 0) + s.СуммаНДФЛ, 2)

    # Заполнение в памяти (4 процедуры внутри оркестратора) — БЕЗ записи
    об.ЗаполнитьДляОтражения_ДокументОтражениеЗарплатыВФинансовомУчете()

    # НачисленныйНДФЛ Ф1 per (ФЛ, ОргБух)
    ндфл = {}
    ndfl_empty = 0
    for i in range(об.НачисленныйНДФЛ.Количество()):
        s = об.НачисленныйНДФЛ.Получить(i)
        if erp.XMLСтрока(s.ФормаPL) != Ф1x:
            continue
        if not erp.ЗначениеЗаполнено(s.ОрганизацияБухгалтерия):
            ndfl_empty += 1
        флн = nm(s.ФизическоеЛицо, erp)
        орг = nm(s.ОрганизацияБухгалтерия, erp)
        ндфл[(флн, орг)] = round(ндфл.get((флн, орг), 0) + s.Сумма, 2)

    # gross-up «Удержание» Ф1 per (ФЛ, ОргБух)
    удерж = {}
    ud_empty = 0
    for i in range(об.НачисленнаяЗарплатаИВзносыПоФизлицам.Количество()):
        s = об.НачисленнаяЗарплатаИВзносыПоФизлицам.Получить(i)
        if erp.XMLСтрока(s.ФормаPL) != Ф1x:
            continue
        sp = s.СпособОтраженияЗарплатыВБухучете
        if not erp.ЗначениеЗаполнено(sp) or not sp.А_ЭтоУдержание:
            continue
        if s.ВзносыВсего != 0:
            continue
        if not erp.ЗначениеЗаполнено(s.ОрганизацияБухгалтерия):
            ud_empty += 1
        флн = nm(s.ФизическоеЛицо, erp)
        орг = nm(s.ОрганизацияБухгалтерия, erp)
        удерж[(флн, орг)] = round(удерж.get((флн, орг), 0) + s.Сумма, 2)

    # multi-org ФЛ по источнику
    by_fl = {}
    for (флн, орг), v in нб.items():
        by_fl.setdefault(флн, {})[орг] = v
    multi = {fl: o for fl, o in by_fl.items() if len(o) >= 2}

    print(f"\nИсточник НалогиБухгалтерия: ФЛ с НДФЛ={len(by_fl)}, multi-org={len(multi)}")
    print(f"НачисленныйНДФЛ Ф1: строк-групп={len(ндфл)}, с ПУСТОЙ ОргБух={ndfl_empty}  (GREEN=0)")
    print(f"«Удержание» Ф1:     строк-групп={len(удерж)}, с ПУСТОЙ ОргБух={ud_empty}  (GREEN=0)")

    print("\n=== Сверка per (multi-org ФЛ, Орг): НДФЛ-ТЧ vs НалогиБухгалтерия vs «Удержание» ===")
    ok = True
    for fl in sorted(multi):
        for орг in sorted(multi[fl]):
            src = multi[fl][орг]
            n = ндфл.get((fl, орг), 0.0)
            u = удерж.get((fl, орг), 0.0)
            match_n = abs(n - src) < 0.01
            match_u = abs(u - src) < 0.01
            if not (match_n and match_u):
                ok = False
            flag = "OK" if (match_n and match_u) else "<<< MISMATCH"
            print(f"  {fl[:24]:24s} | {орг:30s} | НБ={src:>10.2f} | НДФЛ={n:>10.2f} | Удерж={u:>10.2f} | {flag}")

    verdict = (ndfl_empty == 0 and ud_empty == 0 and ok)
    print(f"\nИТОГ: {'GREEN ✅' if verdict else 'RED ❌'} "
          f"(пустых НДФЛ={ndfl_empty}, пустых Удерж={ud_empty}, сверка per-org={'OK' if ok else 'MISMATCH'})")


if __name__ == "__main__":
    main()
