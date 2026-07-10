# -*- coding: utf-8 -*-
"""Smoke-тест обработки «Загрузка работ в СС для базы ЕРП» (BaseERP).

Эталон считается openpyxl из того же Excel (лист «Проект СС», колонка C = "Робота",
пустое G -> пропуск; Этап<-E, Работа<-G, Единица<-I, Количество<-J, Цена<-K, Сумма<-M),
затем движок .epf грузит ТЧ Работы тестового элемента А_СтруктураСебестоимости
«__ТЕСТ Загрузка работ ЕРП» (get-or-create, НЕ удаляется), после чего сверяются
строки/суммы/этапы/единицы и проверяется идемпотентность (повтор: 0 новых работ/этапов).
Единицы резолвятся в базовые (владелец НаборыУпаковок.БазовыеЕдиницыИзмерения) по тем же
правилам, что и движок: точное имя -> нормализованное -> без хвостовой точки -> полное имя.

Запуск: C:\\Python313\\python.exe test_zagruzka_rabot_erp.py [путь_к_xlsx]
"""
import sys
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Загрузка СС\Загрузка работ в СС для базы ЕРП.epf"
DEFAULT_XLSX = (r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Загрузка СС\Виробництво\1"
                r"\IRC 15м2 НОВИЙ ШАблон  СС Виробництво 16-06-2026_Коррект (1).xlsx")
TEST_NAME = "__ТЕСТ Загрузка работ ЕРП"
TOL = 0.005

ERRORS = []


def check(cond, msg):
    print(("OK   " if cond else "FAIL ") + msg)
    if not cond:
        ERRORS.append(msg)


def norm(s):
    s = ("" if s is None else str(s)).strip()
    while "  " in s:
        s = s.replace("  ", " ")
    return s


def r2(x):
    return float(Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP))


def strip_dot(s):
    while s.endswith("."):
        s = s[:-1].strip()
    return s


def build_unit_resolver(erp):
    """Зеркало ПостроитьКэшЕдиниц движка: ключ -> каноничное Наименование."""
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ Наименование, НаименованиеПолное, Ссылка ИЗ Справочник.УпаковкиЕдиницыИзмерения\n"
              "ГДЕ Владелец = ЗНАЧЕНИЕ(Справочник.НаборыУпаковок.БазовыеЕдиницыИзмерения)")
    tab = q.Execute().Выгрузить()
    items = [((tab.Получить(i).Наименование or ""), (tab.Получить(i).НаименованиеПолное or ""))
             for i in range(tab.Количество())]
    keys = {}
    for prio in (1, 2, 3, 4):
        for naim, full in items:
            nn = norm(naim)
            if prio == 1:
                if naim != nn:
                    continue
                k = nn.lower()
            elif prio == 2:
                k = nn.lower()
            elif prio == 3:
                k = strip_dot(nn).lower()
            else:
                k = strip_dot(norm(full)).lower()
            if k and k not in keys:
                keys[k] = naim
    def resolve(name):
        k = norm(name).lower()
        if not k:
            return ""
        return keys.get(k) or keys.get(strip_dot(k)) or ""
    return resolve


def read_excel(path):
    """Эталон: та же фильтрация и маппинг, что у клиента формы."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Проект СС"]
    rows, total, skipped = [], 0, 0
    exp = {"rows": 0, "kol": 0.0, "summa": 0.0, "stages": {}}
    for row in ws.iter_rows(min_row=2):
        if norm(row[2].value) != "Робота":
            continue
        total += 1
        rabota = str(row[6].value).strip() if row[6].value is not None else ""
        if not rabota:
            skipped += 1
            continue
        etap = str(row[4].value).strip() if row[4].value is not None else ""
        edinica = str(row[8].value).strip() if row[8].value is not None else ""
        kol = float(row[9].value or 0)
        cena = float(row[10].value or 0)
        summa = float(row[12].value or 0)
        rows.append({"Этап": etap, "Работа": rabota, "Единица": edinica,
                     "Количество": kol, "Цена": cena, "Сумма": summa})
        # ожидание в БД: построчное округление под квалификаторы ТЧ (2/2)
        exp["rows"] += 1
        exp["kol"] += r2(kol)
        exp["summa"] += r2(summa)
        key = norm(etap).lower()
        exp["stages"][key] = exp["stages"].get(key, 0.0) + r2(summa)
    wb.close()
    print(f"Excel: строк «Робота» = {total}, взято = {len(rows)}, пропущено пустых = {skipped}")
    print(f"Excel-эталон (окр. 2/2): строк={exp['rows']}, ΣКол={exp['kol']:.2f}, ΣСумма={exp['summa']:.2f}, этапов={len(exp['stages'])}")
    return rows, exp


def names_set(erp, query, param=None):
    q = erp.NewObject("Запрос")
    q.Text = query
    if param:
        q.SetParameter(*param)
    tab = q.Execute().Выгрузить()
    return {tab.Получить(i).Наименование for i in range(tab.Количество())}


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    print("Файл:", xlsx)
    rows, exp = read_excel(xlsx)

    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

    # --- тестовый элемент: get-or-create (НЕ удалять) ---
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ Ссылка КАК СсылкаЭлемента ИЗ Справочник.А_СтруктураСебестоимости\n"
              "ГДЕ Наименование = &Н И НЕ ЭтоГруппа")
    q.SetParameter("Н", TEST_NAME)
    sel = q.Execute().Выбрать()
    if sel.Следующий():
        структура = sel.СсылкаЭлемента
        print("Тест-элемент найден:", erp.String(структура))
    else:
        obj = erp.Справочники.А_СтруктураСебестоимости.СоздатьЭлемент()
        obj.Наименование = TEST_NAME
        obj.Записать()
        структура = obj.Ссылка
        print("Тест-элемент создан:", erp.String(структура))

    # --- снимок работ/этапов ДО (для списка созданных) ---
    гр_q = erp.NewObject("Запрос")
    гр_q.Text = "ВЫБРАТЬ Ссылка КАК СсылкаГруппы ИЗ Справочник.Номенклатура ГДЕ Код = &Код И ЭтоГруппа"
    гр_q.SetParameter("Код", "0Ц-00000016")
    гр_sel = гр_q.Execute().Выбрать()
    гр_sel.Следующий()
    группа = гр_sel.СсылкаГруппы
    Q_RABOTY = ("ВЫБРАТЬ Наименование ИЗ Справочник.Номенклатура\n"
                "ГДЕ НЕ ЭтоГруппа И Родитель В ИЕРАРХИИ(&Гр)")
    Q_ETAPY = "ВЫБРАТЬ Наименование ИЗ Справочник.А_ЭтапыРабот"
    работы_до = names_set(erp, Q_RABOTY, ("Гр", группа))
    этапы_до = names_set(erp, Q_ETAPY)

    # --- движок из .epf ---
    proc = erp.ВнешниеОбработки.Создать(EPF, False)

    массив = erp.NewObject("Массив")
    for r in rows:
        st = erp.NewObject("Структура", "Этап, Работа, Единица, Количество, Цена, Сумма")
        st.Вставить("Этап", r["Этап"])
        st.Вставить("Работа", r["Работа"])
        st.Вставить("Единица", r["Единица"])
        st.Вставить("Количество", float(r["Количество"]))
        st.Вставить("Цена", float(r["Цена"]))
        st.Вставить("Сумма", float(r["Сумма"]))
        массив.Добавить(st)

    def прогон(метка):
        res = proc.ЗагрузитьРаботыИзМассива(структура, массив)
        print(f"--- Протокол ({метка}) ---")
        print(f"  Загружено={res.Загружено}, СозданоРабот={res.СозданоРабот}, "
              f"СозданоЭтапов={res.СозданоЭтапов}, Пропущено={res.Пропущено}, "
              f"КонтрольПройден={res.КонтрольПройден}, Ошибка={res.Ошибка}")
        if res.Ошибка:
            print("  ТЕКСТ ОШИБКИ:", res.ТекстОшибки)
        for i in range(res.Контроль.Количество()):
            print("  " + res.Контроль.Получить(i))
        for i in range(res.Инфо.Количество()):
            print("  " + res.Инфо.Получить(i))
        return res

    # === прогон 1 ===
    res1 = прогон("прогон 1")
    check(not res1.Ошибка, "прогон 1: без ошибки")
    check(bool(res1.КонтрольПройден), "прогон 1: КонтрольПройден")
    check(res1.Загружено == exp["rows"], f"прогон 1: Загружено {res1.Загружено} == {exp['rows']}")

    # --- сверка с БД ---
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ Р.Этап.Наименование КАК Этап, СУММА(Р.Количество) КАК Кол, СУММА(Р.Сумма) КАК Сум,\n"
              "	КОЛИЧЕСТВО(*) КАК Строк\n"
              "ИЗ Справочник.А_СтруктураСебестоимости.Работы КАК Р\n"
              "ГДЕ Р.Ссылка = &Структура\n"
              "СГРУППИРОВАТЬ ПО Р.Этап.Наименование")
    q.SetParameter("Структура", структура)
    tab = q.Execute().Выгрузить()
    db_rows, db_kol, db_sum, db_stages = 0, 0.0, 0.0, {}
    for i in range(tab.Количество()):
        row = tab.Получить(i)
        db_rows += int(row.Строк)
        db_kol += float(row.Кол)
        db_sum += float(row.Сум)
        db_stages[norm(row.Этап).lower()] = float(row.Сум)

    check(db_rows == exp["rows"], f"БД: строк ТЧ {db_rows} == {exp['rows']}")
    check(abs(db_kol - exp["kol"]) <= TOL, f"БД: ΣКоличество {db_kol:.3f} ~ {exp['kol']:.2f}")
    check(abs(db_sum - exp["summa"]) <= TOL, f"БД: ΣСумма {db_sum:.2f} ~ {exp['summa']:.2f}")
    check(set(db_stages) == set(exp["stages"]), f"БД: набор этапов совпал ({len(db_stages)})")
    for k in sorted(exp["stages"]):
        if k in db_stages:
            check(abs(db_stages[k] - exp["stages"][k]) <= TOL,
                  f"БД: этап «{k}» {db_stages[k]:.2f} ~ {exp['stages'][k]:.2f}")

    # --- сверка единиц построчно (те же правила резолвинга, что в движке) ---
    resolve = build_unit_resolver(erp)
    expected_units = [resolve(r["Единица"]) for r in rows]
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ Р.НомерСтроки КАК НомерСтроки, Р.Единица.Наименование КАК ЕдиницаНаим\n"
              "ИЗ Справочник.А_СтруктураСебестоимости.Работы КАК Р\n"
              "ГДЕ Р.Ссылка = &Структура\n"
              "УПОРЯДОЧИТЬ ПО Р.НомерСтроки")
    q.SetParameter("Структура", структура)
    tab = q.Execute().Выгрузить()
    db_units = [(tab.Получить(i).ЕдиницаНаим or "") for i in range(tab.Количество())]
    check(len(db_units) == len(expected_units), f"БД: строк для сверки единиц {len(db_units)}")
    mismatches = [(i + 1, expected_units[i], db_units[i])
                  for i in range(min(len(db_units), len(expected_units)))
                  if db_units[i] != expected_units[i]]
    check(not mismatches, f"БД: единицы построчно совпали (расхождений {len(mismatches)})")
    for num, e, d in mismatches[:10]:
        print(f"     строка {num}: ожидалось «{e}», в БД «{d}»")
    filled = sum(1 for u in db_units if u)
    not_found = {}
    for r, e in zip(rows, expected_units):
        if r["Единица"] and not e:
            not_found[r["Единица"]] = not_found.get(r["Единица"], 0) + 1
    print(f"Единицы: заполнено {filled} из {len(db_units)}; не найдено в базовых: "
          + (", ".join(f"«{k}»×{v}" for k, v in sorted(not_found.items())) or "нет"))

    # === прогон 2: идемпотентность ===
    res2 = прогон("прогон 2, идемпотентность")
    check(not res2.Ошибка, "прогон 2: без ошибки")
    check(bool(res2.КонтрольПройден), "прогон 2: КонтрольПройден")
    check(res2.Загружено == res1.Загружено, f"прогон 2: те же строки ({res2.Загружено})")
    check(res2.СозданоРабот == 0, "прогон 2: СозданоРабот = 0")
    check(res2.СозданоЭтапов == 0, "прогон 2: СозданоЭтапов = 0")

    # --- созданные этапы/работы (первый прогон этого файла) ---
    работы_после = names_set(erp, Q_RABOTY, ("Гр", группа))
    этапы_после = names_set(erp, Q_ETAPY)
    новые_этапы = sorted(этапы_после - этапы_до)
    новые_работы = sorted(работы_после - работы_до)
    print(f"Создано этапов в этом запуске: {len(новые_этапы)}")
    for n in новые_этапы:
        print("  этап +", n)
    print(f"Создано работ в этом запуске: {len(новые_работы)} (первые 20):")
    for n in новые_работы[:20]:
        print("  работа +", n)

    print()
    if ERRORS:
        print(f"ИТОГ: ПРОВАЛЕНО {len(ERRORS)} проверок:")
        for e in ERRORS:
            print("  -", e)
        return 1
    print("ИТОГ: SMOKE ПРОЙДЕН")
    return 0


if __name__ == "__main__":
    sys.exit(main())
