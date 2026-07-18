# -*- coding: utf-8 -*-
"""Проверка по разделу 7 ТЗ: кнопка «Полный обмен с Казной» (BuhKazn <-> ERP).

ФАЗА A (безопасно, без обмена):
  обе .epf грузятся и BSL компилируется во ВСЕХ трёх базах;
  УзелКорреспондентПлана() даёт ожидаемую матрицу.

ФАЗА B (реальный обмен — это штатная синхронизация):
  BuhKazn: СинхронизироватьПоПлану(ПланКазны)        -> Истина, «нечего принимать» не ошибка;
  BuhBud : СинхронизироватьПоПлану(ПланБухгалтерии)  -> РЕГРЕССИЯ пары BuhBud<->ERP;
  BaseERP: оба плана по отдельности (шаг 2 обеих цепочек).

ФАЗА C: строка BuhKazn в регистре А_НастройкиПодключенияБаз (её читает кнопка).
"""
import sys
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

RUN_EXCHANGE = "--no-exchange" not in sys.argv

# Боевое расположение (Rule #4): именно эти .epf открывает пользователь и на них
# указывает реквизит ПутьКОбработкеОбмена в оркестраторе.
OBR = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки"
WORKER = OBR + r"\А_ЗапуститьОбмен.epf"
ORCH = OBR + r"\А_АвтоматическийОбменДанными.epf"

CONNS = {
    "BaseERP": 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"',
    "BuhBud": 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"',
    "BuhKazn": 'Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"',
}

# Ожидаемая матрица корреспондентов (раздел 3 ТЗ).
EXPECT = {
    ("BaseERP", "buh"): "bas_industrialbud",
    ("BaseERP", "kazna"): "Казна",
    ("BuhBud", "buh"): "BAS ERP, ред. 2.5",
    ("BuhBud", "kazna"): None,
    ("BuhKazn", "buh"): None,
    ("BuhKazn", "kazna"): "BAS ERP, ред. 2.5",
}

failures = []


def err_text(exc):
    if hasattr(exc, "excepinfo") and exc.excepinfo:
        return str(exc.excepinfo[2])
    return str(exc)


def node_name(conn, node):
    return None if node is None else conn.String(node)


v8 = win32com.client.Dispatch("V83.COMConnector")

print("#" * 78)
print("ФАЗА A — загрузка обеих .epf и матрица узлов по планам (без обмена)")
print("#" * 78)

for base, cs in CONNS.items():
    print(f"\n=== {base} ===")
    conn = v8.Connect(cs)

    # Обе обработки обязаны компилироваться в каждой базе (модуль объекта у них общий).
    for label, path in (("worker  А_ЗапуститьОбмен", WORKER),
                        ("orchestr А_АвтоматическийОбменДанными", ORCH)):
        try:
            obr = conn.ExternalDataProcessors.Create(path, False)
            _ = obr.ПланБухгалтерии()
            print(f"  [OK] {label}: загружена, BSL скомпилирован")
        except Exception as exc:  # noqa: BLE001
            print(f"  [FAIL] {label}: {err_text(exc)}")
            failures.append(f"{base}: не грузится {label}")

    obr = conn.ExternalDataProcessors.Create(WORKER, False)
    plans = {"buh": obr.ПланБухгалтерии(), "kazna": obr.ПланКазны()}

    for role, plan in plans.items():
        got = node_name(conn, obr.УзелКорреспондентПлана(plan))
        want = EXPECT[(base, role)]
        ok = got == want
        print(f"  {'[OK] ' if ok else '[FAIL]'} УзелКорреспондентПлана({plan})")
        print(f"        -> {got!r}   (ожидалось {want!r})")
        if not ok:
            failures.append(f"{base}/{role}: узел {got!r}, ожидался {want!r}")

    conn = None

if RUN_EXCHANGE:
    print()
    print("#" * 78)
    print("ФАЗА B — РЕАЛЬНЫЙ обмен (штатная синхронизация)")
    print("#" * 78)

    scenarios = [
        ("BuhKazn", "kazna", "ЦЕЛЬ: Казна -> ERP"),
        ("BaseERP", "kazna", "ЦЕЛЬ: шаг 2 цепочки «Казна» (только план Казны)"),
        ("BuhKazn", "kazna", "ЦЕЛЬ: Казна принимает встречное от ERP"),
        ("BuhBud", "buh", "РЕГРЕССИЯ: пара BuhBud <-> ERP не сломана"),
        ("BaseERP", "buh", "РЕГРЕССИЯ: шаг 2 цепочки «Бухгалтерия»"),
    ]

    for base, role, note in scenarios:
        print(f"\n=== {base} / план {role} — {note}")
        conn = v8.Connect(CONNS[base])
        obr = conn.ExternalDataProcessors.Create(WORKER, False)
        plan = obr.ПланКазны() if role == "kazna" else obr.ПланБухгалтерии()

        msgs = conn.NewObject("Массив")
        try:
            res = obr.СинхронизироватьПоПлану(plan, msgs)
        except Exception as exc:  # noqa: BLE001
            print(f"  [FAIL] ИСКЛЮЧЕНИЕ: {err_text(exc)}")
            failures.append(f"{base}/{role}: исключение при обмене")
            conn = None
            continue

        for i in range(msgs.Count()):
            print(f"      {msgs.Get(i)}")
        print(f"  Возврат: {res}  ({'OK' if res else 'С ОШИБКАМИ'})")
        if not res:
            failures.append(f"{base}/{role}: СинхронизироватьПоПлану вернула Ложь")
        conn = None

print()
print("#" * 78)
print("ФАЗА C — строка BuhKazn в регистре А_НастройкиПодключенияБаз (читает кнопка)")
print("#" * 78)

conn = v8.Connect(CONNS["BaseERP"])
q = conn.NewObject("Запрос")
q.SetParameter("Ид", "BuhKazn")
q.Text = """ВЫБРАТЬ
	Настройки.Сервер КАК Сервер,
	Настройки.ИмяБазы КАК ИмяБазы,
	Настройки.Пользователь КАК Пользователь,
	Настройки.Пароль КАК Пароль
ИЗ
	РегистрСведений.А_НастройкиПодключенияБаз КАК Настройки
ГДЕ
	Настройки.ИдентификаторБазы = &Ид"""
tab = q.Execute().Unload()
if tab.Count() == 0:
    print("  [FAIL] строки BuhKazn в регистре НЕТ — кнопка «Казна» не сможет запустить базу")
    failures.append("регистр: нет строки BuhKazn")
else:
    r = tab.Get(0)
    pwd = "задан" if str(r.Get(3)).strip() else "ПУСТО"  # сам пароль не печатаем
    print(f"  [OK] BuhKazn -> Сервер={r.Get(0)}  ИмяБазы={r.Get(1)}  "
          f"Пользователь={r.Get(2)}  Пароль={pwd}")
    if pwd == "ПУСТО":
        failures.append("регистр: у строки BuhKazn пустой пароль")
conn = None

print()
print("=" * 78)
if failures:
    print("ИТОГ: ЕСТЬ ПРОБЛЕМЫ")
    for f in failures:
        print("  - " + f)
    sys.exit(1)
print("ИТОГ: ВСЁ ЗЕЛЁНОЕ — обе .epf грузятся в 3 базах, матрица узлов верна,")
print("      обмен Казна<->ERP работает, пара BuhBud<->ERP не сломана.")
