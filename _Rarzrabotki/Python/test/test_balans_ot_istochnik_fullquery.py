# -*- coding: utf-8 -*-
"""Rule #-1: извлечь ФАКТИЧЕСКИЙ текст запросов Свод_ОплатаТруда и
Свод_ПрочиеАктивыПассивы_Прямой из отредактированного BSL и выполнить через COM
с реальными параметрами января — проверить отсутствие синтакс-ошибок и суммы по статьям."""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def f(x):
    try: return float(x)
    except: return 0.0

BSL = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\angry-poincare-d9bb04\Documents\А_ФинРез_Баланс\Ext\ObjectModule.bsl"
src = open(BSL, encoding="utf-8", newline="").read().replace("\r\n", "\n")
lines = src.split("\n")

def extract_query(func_name):
    """Извлечь текст Запрос.Текст = "..."; внутри функции func_name."""
    # найти строку с объявлением функции
    start = next(i for i, ln in enumerate(lines) if ("Функция " + func_name) in ln or ("Процедура " + func_name) in ln)
    # найти Запрос.Текст = после start
    qi = next(i for i in range(start, len(lines)) if "Запрос.Текст =" in lines[i])
    # текст начинается со следующей строки (открывается ")
    out = []
    i = qi + 1
    while i < len(lines):
        ln = lines[i]
        # снять ведущий таб
        s = ln[1:] if ln.startswith("\t") else ln
        if s.startswith('"'):           # первая строка запроса: \t"ВЫБРАТЬ
            s = s[1:]
        elif s.startswith("|"):         # продолжение: |<...>
            s = s[1:]
        end = s.rstrip().endswith('";')
        if end:
            s = s.rstrip()[:-2]         # убрать ";
        out.append(s)
        if end:
            break
        i += 1
    text = "\n".join(out).replace('""', '"')
    return text

q_ot = extract_query("Свод_ОплатаТруда")
q_pr = extract_query("Свод_ПрочиеАктивыПассивы_Прямой")
print(f"Свод_ОплатаТруда: извлечено {len(q_ot)} символов; начало: {q_ot[:40]!r}; конец: {q_ot[-40:]!r}")
print(f"Прямой:           извлечено {len(q_pr)} символов; начало: {q_pr[:40]!r}; конец: {q_pr[-40:]!r}")

# --- параметры (январь 2026) ---
qo = erp.NewObject("Запрос")
qo.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1 Организация КАК О ИЗ Документ.А_ФинРез_Баланс ГДЕ Проведен УПОРЯДОЧИТЬ ПО Дата"
Орг = qo.Выполнить().Выгрузить()[0].О

def stat(name):
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Н", name)
    q.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование = &Н"
    r = q.Выполнить().Выгрузить(); return r[0].С if r.Количество() else None

import datetime
def run(qtext, params):
    q = erp.NewObject("Запрос")
    for k, v in params.items():
        q.УстановитьПараметр(k, v)
    q.Текст = qtext
    return q.Выполнить().Выгрузить()

# Искл = массив статей "Собственные средства" (1С Массив, не Python list)
arrИмена = erp.NewObject("Массив")
arrИмена.Добавить("Собственные средства")
qi = erp.NewObject("Запрос")
qi.Текст = ("ВЫБРАТЬ Ссылка КАК С ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов "
            "ГДЕ Наименование В (&И)")
qi.УстановитьПараметр("И", arrИмена)
Искл = qi.Выполнить().Выгрузить().ВыгрузитьКолонку("С")

# НачМес/КонМес через 1С-литералы нельзя в параметре (tz) — но эти запросы юзают параметры.
# КонМес у документа = КонецДня(КонецМесяца). Для теста передадим даты как в проведении.
# ВНИМАНИЕ tz: используем datetime с полуднем для границ нельзя (нужен точный конец).
# Поэтому границы вычислим в 1С и передадим как значения.
qb = erp.NewObject("Запрос")
qb.Текст = "ВЫБРАТЬ НАЧАЛОПЕРИОДА(ДАТАВРЕМЯ(2026,1,15), МЕСЯЦ) КАК НМ, КОНЕЦПЕРИОДА(ДАТАВРЕМЯ(2026,1,15), МЕСЯЦ) КАК КМ"
b = qb.Выполнить().Выгрузить()[0]
НачМес, КонМес = b.НМ, b.КМ
print(f"НачМес={НачМес}  КонМес={КонМес}")

common = {"Орг": Орг, "НачМес": НачМес, "КонМес": КонМес, "Искл": Искл, "ОТ": stat("Оплата труда")}
pr_params = dict(common)
pr_params["ПУ"] = stat("Прибыли и убытки")
pr_params["РасхТекПер"] = stat("Расходы текущего периода")
pr_params["ДохТекПер"] = stat("Доходы текущего периода")

print("\n--- Выполняю Свод_ОплатаТруда (январь) ---")
try:
    tz_ot = run(q_ot, common)
    by = {}
    for r in tz_ot:
        nm = r.Статья.Наименование
        by[nm] = by.get(nm, 0.0) + f(r.СуммаКонечныйОстаток)
    for k, v in by.items():
        print(f"   Статья={k}: Σ КО={v:,.2f}")
    ot_total = sum(by.values())
    print(f"   ОТ ИТОГО = {ot_total:,.2f}  (эталон +246 187,13)  {'OK' if abs(ot_total-246187.13)<0.1 else 'FAIL'}")
except Exception as e:
    ei = getattr(e, 'excepinfo', None)
    print(f"   FAIL OT: {ei[2] if ei else e}")

print("\n--- Выполняю Свод_ПрочиеАктивыПассивы_Прямой (январь) ---")
try:
    tz_pr = run(q_pr, pr_params)
    kas = 0.0
    for r in tz_pr:
        if r.Статья.Наименование == "Денежные средства (наличные)":
            kas += f(r.СуммаКонечныйОстаток)
    print(f"   касса наличная (вклад Прямого) Σ КО = {kas:,.2f}  (ожидаем -9 677 001,50)  {'OK' if abs(kas+9677001.50)<0.1 else 'FAIL'}")
    print(f"   всего строк Прямого: {tz_pr.Количество()}")
except Exception as e:
    ei = getattr(e, 'excepinfo', None)
    print(f"   FAIL Прямой: {ei[2] if ei else e}")
