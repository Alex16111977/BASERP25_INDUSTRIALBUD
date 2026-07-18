# -*- coding: utf-8 -*-
"""Воспроизводит ПровестиБалансСвод за декабрь 2025 (3 активные Свод_*),
эмулирует Свернуть(19 dim) и находит коллизию ключа регистра:
2 строки, которые Свернуть РАЗЛИЧАЕТ (НЕОПРЕДЕЛЕНО vs пустая ссылка),
а РегистрСведений схлопывает в один ключ → "запись с такими ключевыми полями
существует". READ-ONLY (никаких проведений/правок)."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client
from collections import defaultdict

conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# период декабрь 2025 (серверно)
from datetime import datetime
M = datetime(2025, 12, 1, 12, 0, 0)
qb = conn.NewObject("Запрос")
qb.Текст = ("ВЫБРАТЬ НАЧАЛОПЕРИОДА(&М,МЕСЯЦ) КАК НМ, "
            "КОНЕЦПЕРИОДА(КОНЕЦПЕРИОДА(&М,МЕСЯЦ),ДЕНЬ) КАК КД")
qb.УстановитьПараметр("М", M); rb = qb.Выполнить().Выбрать(); rb.Следующий()
НМ, КД = rb.НМ, rb.КД
print(f"період: {НМ} .. {КД}")

qO = conn.NewObject("Запрос")
qO.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Справочник.Организации '
            'ГДЕ КодПоЕДРПОУ = "40645273"')
sO = qO.Выполнить().Выбрать(); sO.Следующий(); ORG = sO.С

# Искл (ИсключенныеСтатьи)
qI = conn.NewObject("Запрос")
qI.Текст = ('ВЫБРАТЬ Ссылка КАК Ссылка ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов '
            'ГДЕ Наименование В (&Имена)')
им = conn.NewObject("Массив")
for n in ("Собственные средства","Доходы текущего периода","Расходы текущего периода"):
    им.Добавить(n)
qI.УстановитьПараметр("Имена", им)
tI = qI.Выполнить().Выгрузить()
Искл = conn.NewObject("Массив")
for i in range(tI.Количество()):
    Искл.Добавить(tI.Получить(i).Ссылка)

# _РефыИсточников + статьи денег (как в Свод_ДенежныеСредства)
qSrc = conn.NewObject("Запрос")
qSrc.Текст = ("ВЫБРАТЬ Ссылка КАК С, ПРЕДСТАВЛЕНИЕ(Ссылка) КАК П "
              "ИЗ Перечисление.ИсточникиУправленческогоБаланса")
tSrc = qSrc.Выполнить().Выгрузить()
src = {}
for i in range(tSrc.Количество()):
    src[str(tSrc.Получить(i).П).strip()] = tSrc.Получить(i).С
def рефы(преды):
    return [src[p] for p in преды if p in src]
МассивПредст = ["Денежные средства (безналичные)","Денежные средства (наличные)",
                "Денежные средства в пути","Денежные средства у подотчетных лиц"]
МассивИст = conn.NewObject("Массив")
for p in МассивПредст:
    if p in src: МассивИст.Добавить(src[p])
ИстБезнал = src["Денежные средства (безналичные)"]
ИстНал = src["Денежные средства (наличные)"]
ИстПодотч = src["Денежные средства у подотчетных лиц"]
print(f"ИстБезнал={conn.XMLСтрока(ИстБезнал)} ИстНал={conn.XMLСтрока(ИстНал)} "
      f"ИстПодотч={conn.XMLСтрока(ИстПодотч)}")
qSt = conn.NewObject("Запрос")
qSt.Текст = ("ВЫБРАТЬ Ссылка КАК С, Наименование КАК Н "
             "ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование В (&И)")
иСт = conn.NewObject("Массив")
for n in ("Денежные средства (безналичные)","Денежные средства (наличные)",
          "Денежные средства (у подотчетных лиц)"):
    иСт.Добавить(n)
qSt.УстановитьПараметр("И", иСт)
tSt = qSt.Выполнить().Выгрузить()
пэСс = conn.ПланыВидовХарактеристик.СтатьиАктивовПассивов.ПустаяСсылка()
СтБезнал = СтНал = СтПодотч = пэСс
for i in range(tSt.Количество()):
    н = str(tSt.Получить(i).Н); c = tSt.Получить(i).С
    if н == "Денежные средства (безналичные)": СтБезнал = c
    elif н == "Денежные средства (наличные)": СтНал = c
    elif н == "Денежные средства (у подотчетных лиц)": СтПодотч = c
print(f"СтБезнал={conn.XMLСтрока(СтБезнал)} СтНал={conn.XMLСтрока(СтНал)} "
      f"СтПодотч={conn.XMLСтрока(СтПодотч)}")

DIMS = ["Организация","Подразделение","Статья","Номенклатура","Контрагент",
        "Партнер","Склад","ОбъектыЭксплуатации","Договор","ФизическоеЛицо",
        "ДенежныеСредства","Source","Расхождение","Аналитика1","Аналитика2",
        "Аналитика3","ОбъектРасчетов","НематериальныйАктив","ДокументДвижения"]

def cell(v):
    """('UNDEF'|'EMPTY:<тип>'|'<uuid>'|'B:0/1', filled?)"""
    if v is None:
        return ("UNDEF", False)
    try:
        if isinstance(v, bool):
            return (f"B:{int(v)}", True)
    except Exception:
        pass
    try:
        зап = conn.ЗначениеЗаполнено(v)
    except Exception:
        зап = True
    if not зап:
        try:
            return (f"EMPTY:{conn.XMLСтрока(v)[:8]}", False)
        except Exception:
            return ("EMPTY:?", False)
    try:
        return (conn.XMLСтрока(v), True)
    except Exception:
        return (str(v), True)

rows = []  # (func, {dim:rawcell})

def собрать(func, qtext, params, alias_to_dim):
    q = conn.NewObject("Запрос")
    for k, val in params.items():
        q.УстановитьПараметр(k, val)
    q.Текст = qtext
    try:
        t = q.Выполнить().Выгрузить()
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        raise RuntimeError(f"FAIL {func}: {msg}")
    кол = t.Колонки
    имена = [кол.Получить(j).Имя for j in range(кол.Количество())]
    for i in range(t.Количество()):
        r = t.Получить(i)
        d = {}
        for nm in имена:
            dim = alias_to_dim.get(nm, nm)
            if dim in DIMS:
                d[dim] = cell(getattr(r, nm))
        rows.append((func, d))
    print(f"  {func}: строк={t.Количество()}")

P = {"Орг": ORG, "НачМес": НМ, "КонМес": КД, "Искл": Искл}

# --- Свод_СебестоимостьТоваров (точный текст ObjectModule.bsl) ---
QSEB = open("Documents/А_ФинРез_Баланс/Ext/ObjectModule.bsl", encoding="utf-8").read()
def вырез(text, нач_сигнатура):
    i = text.index(нач_сигнатура)
    j = text.index('Запрос.Текст =', i)
    k = text.index('"', j)
    # читаем строковый литерал BSL (до закрывающей " не экранированной "")
    m = k + 1
    buf = []
    while m < len(text):
        ch = text[m]
        if ch == '"':
            if m + 1 < len(text) and text[m+1] == '"':
                buf.append('"'); m += 2; continue
            break
        buf.append(ch); m += 1
    raw = "".join(buf)
    # снять префикс continuation '|'
    out = []
    for ln in raw.split("\n"):
        s = ln.lstrip()
        if s.startswith("|"):
            s = s[1:]
        out.append(s)
    return "\n".join(out)

qseb = вырез(QSEB, "Функция Свод_СебестоимостьТоваров(")
собрать("Себест", qseb, P, {})

# --- Свод_ДенежныеСредства ---
qden = вырез(QSEB, "Функция Свод_ДенежныеСредства(")
Pden = dict(P); Pden.update({"МассивИст": МассивИст, "ИстБезнал": ИстБезнал,
    "ИстНал": ИстНал, "ИстПодотч": ИстПодотч, "СтБезнал": СтБезнал,
    "СтНал": СтНал, "СтПодотч": СтПодотч})
собрать("ДенСр", qden, Pden, {})

# --- Свод_РасчетыСПартнерами ---
qrasch = вырез(QSEB, "Функция Свод_РасчетыСПартнерами(")
собрать("Расчёты", qrasch, P, {})

print(f"\nВСЕГО строк (до Свернуть) = {len(rows)}")

# Эмуляция Свернуть(19) → ключ по СЫРЫМ значениям (UNDEF != EMPTY != ref)
# Регистр-ключ → 'filled?' (UNDEF и EMPTY -> '∅')
def exact_key(d):
    return tuple(d.get(x, ("EMPTY:nil", False))[0] for x in DIMS)
def reg_key(d):
    out = []
    for x in DIMS:
        val, filled = d.get(x, ("EMPTY:nil", False))
        out.append(val if filled else "∅")
    return tuple(out)

# Свернуть: уникальные exact_key (одна строка на exact-комбо)
свернуто = {}
for func, d in rows:
    свернуто.setdefault(exact_key(d), (func, d))
print(f"После Свернуть (уник. exact-ключей) = {len(свернуто)}")

# Коллизия регистра: reg_key с >1 РАЗНЫХ exact_key
по_рег = defaultdict(list)
for ek, (func, d) in свернуто.items():
    по_рег[reg_key(d)].append((ek, func, d))
коллизии = {rk: v for rk, v in по_рег.items() if len(v) > 1}
print(f"\n=== КОЛЛИЗИЙ ключа регистра (Свернуть различает, регистр — нет): {len(коллизии)} ===")
for rk, lst in list(коллизии.items())[:20]:
    # человекочитаемо
    показ = {DIMS[i]: rk[i] for i in range(len(DIMS)) if rk[i] != "∅"}
    print(f"\nКЛЮЧ регистра: {показ}")
    for ek, func, d in lst:
        разн = {DIMS[i]: d.get(DIMS[i], ('EMPTY:nil', False))[0]
                for i in range(len(DIMS))
                if d.get(DIMS[i], ('EMPTY:nil', False))[0] != rk[i]}
        print(f"  [{func}] отличия(сырое vs ∅): {разн}")
if not коллизии:
    print("Коллизий нет — дубль не воспроизведён этой моделью.")
print("=" * 64)
