# -*- coding: utf-8 -*-
"""Rule #-1: проверка запроса УзелКорреспондентПлана на трёх живых базах.

Имена планов обмена НЕ набираются руками (грабля 2 — кириллическая мина
«Казна-ЧЕЙ-ство»): резолвятся ИЗ МЕТАДАННЫХ базы, печатаются codepoints,
и точные имена пишутся в plan_names.txt для подстановки в BSL.

Грабля 6: conn.ExchangePlans["ИмяПлана"] через COM НЕ работает
("This object does not support enumeration") -> доступ только через getattr.

Ожидание (раздел 3 ТЗ):
  BaseERP : Бухгалтерия -> bas_industrialbud ; Казна -> Казна
  BuhBud  : Бухгалтерия -> BAS ERP, ред. 2.5 ; Казна -> плана НЕТ
  BuhKazn : Бухгалтерия -> плана НЕТ         ; Казна -> BAS ERP, ред. 2.5
"""
import sys
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONNS = {
    "BaseERP": 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"',
    "BuhBud": 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"',
    "BuhKazn": 'Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"',
}

# 1:1 текст запроса, который пойдёт в BSL (функция УзелКорреспондентПлана).
QUERY_TMPL = """ВЫБРАТЬ ПЕРВЫЕ 1
	Узлы.Ссылка КАК Узел
ИЗ
	ПланОбмена.{plan} КАК Узлы
ГДЕ
	Узлы.Ссылка <> &ЭтотУзел
	И НЕ Узлы.ПометкаУдаления"""

# Ожидаемое написание имени плана Казны (грабля 2): «Казначейство» + BASERP.
KAZNA_CP = [0x41A, 0x430, 0x437, 0x43D, 0x430, 0x447, 0x435, 0x439,
            0x441, 0x442, 0x432, 0x43E]


def codepoints(text):
    return " ".join(hex(ord(ch)) for ch in text)


def err_text(exc):
    if hasattr(exc, "excepinfo") and exc.excepinfo:
        return str(exc.excepinfo[2])
    return str(exc)


def this_node(conn, plan):
    """ЭтотУзел плана. Только getattr — индексация [] по COM падает (грабля 6)."""
    return getattr(conn.ExchangePlans, plan).ThisNode()


v8 = win32com.client.Dispatch("V83.COMConnector")

# --- Шаг 1: точные имена планов резолвим ОДИН РАЗ из метаданных BaseERP ---
# (в BaseERP оба плана присутствуют и фильтры однозначны; в BuhBud фильтр
#  "…Бухгалтерия20" неоднозначен — там есть и ОбменРозница20Бухгалтерия20).
erp = v8.Connect(CONNS["BaseERP"])
md_erp = erp.Metadata.ExchangePlans
erp_names = [md_erp.Get(i).Name for i in range(md_erp.Count())]

plan_names = {
    "kazna": next(n for n in erp_names if n.endswith("BASERP")),
    "buh": next(n for n in erp_names
                if n.startswith("Обмен") and n.endswith("Бухгалтерия20")),
}
erp = None

results = {}   # (база, роль) -> имя корреспондента / None

for base, cs in CONNS.items():
    print("=" * 78)
    print(f"БАЗА {base}")
    conn = v8.Connect(cs)

    md = conn.Metadata.ExchangePlans
    names = [md.Get(i).Name for i in range(md.Count())]

    for role, plan in (("Бухгалтерия", plan_names["buh"]),
                       ("Казна", plan_names["kazna"])):

        # Как в BSL: сперва Метаданные.ПланыОбмена.Найти(ИмяПлана).
        if md.Find(plan) is None:
            print(f"  {role:12s}: плана {plan} НЕТ в этой базе "
                  f"(Найти -> Неопределено, исключения нет)  OK")
            results[(base, role)] = None
            continue

        q = conn.NewObject("Запрос")
        q.Text = QUERY_TMPL.format(plan=plan)
        q.SetParameter("ЭтотУзел", this_node(conn, plan))
        try:
            tab = q.Execute().Unload()
        except Exception as exc:  # noqa: BLE001
            print(f"  {role:12s}: [FAIL] {plan} -> {err_text(exc)}")
            results[(base, role)] = "FAIL"
            continue

        if tab.Count() == 0:
            print(f"  {role:12s}: план {plan} есть, корреспондента НЕТ")
            results[(base, role)] = None
            continue

        node = conn.String(tab.Get(0).Get(0))
        print(f"  {role:12s}: {plan}")
        print(f"  {'':12s}  корреспондент -> {node}   (строк={tab.Count()})")
        results[(base, role)] = node

    conn = None

print()
print("=" * 78)
print("ИМЕНА ПЛАНОВ ИЗ МЕТАДАННЫХ (грабля 2 — проверка написания):")
for role, plan in plan_names.items():
    print(f"  {role:6s} = {plan}")
    print(f"         codepoints: {codepoints(plan)}")

kz = plan_names.get("kazna", "")
ok_kz = [ord(c) for c in kz[:12]] == KAZNA_CP
print(f"\n  Проверка Казны: первые 12 codepoints == ожидаемым ({'OK' if ok_kz else 'НЕ СОВПАЛО!'})")
print(f"  Буква №8: {kz[7]!r} = {hex(ord(kz[7]))} (должно быть 'й' = 0x439)")

print()
print("=" * 78)
print("СВОДКА:")
for (base, role), node in results.items():
    print(f"  {base:8s} / {role:12s} -> {node}")

with open("plan_names.txt", "w", encoding="utf-8") as f:
    f.write(plan_names.get("buh", "") + "\n")
    f.write(plan_names.get("kazna", "") + "\n")
print("\nТочные имена записаны в plan_names.txt (для подстановки в BSL без набора руками).")
