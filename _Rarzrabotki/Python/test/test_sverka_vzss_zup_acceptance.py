# -*- coding: utf-8 -*-
"""Acceptance: COM-зеркало отчёта А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.
Период: декабрь 2025. Проверки: согласованность Σ, поведение флага, дубликаты ДРФО."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MAPPING = "T0"  # вердикт test_sverka_vzss_zup_pretest2_signs.py (identity)
TOL = 0.01

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
zup = v8.Connect('Srvr="localhost";Ref="zup2";Usr="cfo";Pwd="2442"')

def fetch(conn, text, cols):
    q = conn.NewObject("Запрос")
    q.Text = text
    t = q.Execute().Выгрузить()
    rows = []
    for i in range(t.Количество()):
        r = t.Получить(i)
        rows.append(tuple(getattr(r, c) for c in cols))
    return rows

erp_rows = {}
dup_erp = 0
for fl_drfo, n, p, r, k in fetch(erp, """ВЫБРАТЬ
	Ост.ФизическоеЛицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаКВыплатеНачальныйОстаток) КАК НачОст,
	СУММА(Ост.СуммаКВыплатеПриход) КАК Приход,
	СУММА(Ост.СуммаКВыплатеРасход) КАК Расход,
	СУММА(Ост.СуммаКВыплатеКонечныйОстаток) КАК КонОст
ИЗ
	РегистрНакопления.ЗарплатаКВыплате.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , ) КАК Ост
СГРУППИРОВАТЬ ПО Ост.ФизическоеЛицо.КодПоДРФО""", ("ДРФО", "НачОст", "Приход", "Расход", "КонОст")):
    key = (fl_drfo or "").strip()
    vals = tuple(float(x or 0) for x in (n, p, r, k))
    if key in erp_rows:
        dup_erp += 1
        old = erp_rows[key]
        vals = tuple(a + b for a, b in zip(old, vals))
    erp_rows[key] = vals

zup_rows = {}
for drfo, n, p, r, k in fetch(zup, """ВЫБРАТЬ
	Ост.Физлицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаУпрНачальныйОстаток) КАК НачОст,
	СУММА(Ост.СуммаУпрПриход) КАК Приход,
	СУММА(Ост.СуммаУпрРасход) КАК Расход,
	СУММА(Ост.СуммаУпрКонечныйОстаток) КАК КонОст
ИЗ
	РегистрНакопления.ВзаиморасчетыСРаботниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , Физлицо.КодПоДРФО <> "") КАК Ост
СГРУППИРОВАТЬ ПО Ост.Физлицо.КодПоДРФО""", ("ДРФО", "НачОст", "Приход", "Расход", "КонОст")):
    key = (drfo or "").strip()
    vals = tuple(float(x or 0) for x in (n, p, r, k))
    if MAPPING == "T1":
        vals = (-vals[0], vals[2], vals[1], -vals[3])
    if key in zup_rows:
        old = zup_rows[key]
        vals = tuple(a + b for a, b in zip(old, vals))
    zup_rows[key] = vals

# Корректировка ЗУП: + «Утримано по бухгалтерії» (00001) к Начислениям и Выплатам
_utr_rows = {}
for drfo, s in fetch(zup, """ВЫБРАТЬ
	Удержания.Физлицо.КодПоДРФО КАК ДРФО,
	СУММА(Удержания.Результат) КАК СуммаУтрБух
ИЗ
	Документ.НачислениеЗарплатыРаботникам.Удержания КАК Удержания
ГДЕ
	Удержания.Ссылка.Дата МЕЖДУ ДАТАВРЕМЯ(2025, 12, 1) И ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59)
	И Удержания.Ссылка.Проведен = ИСТИНА
	И Удержания.Результат <> 0
	И Удержания.ВидРасчета.Код = "00001"
	И Удержания.Физлицо.КодПоДРФО <> ""
СГРУППИРОВАТЬ ПО
	Удержания.Физлицо.КодПоДРФО""", ("ДРФО", "СуммаУтрБух")):
    key = (drfo or "").strip()
    _utr_rows[key] = _utr_rows.get(key, 0.0) + float(s or 0)
for key, s in _utr_rows.items():
    n, p, r, k = zup_rows.get(key, (0.0, 0.0, 0.0, 0.0))
    zup_rows[key] = (n, p + s, r + s, k)

membership = set()
for (drfo,) in fetch(zup, """ВЫБРАТЬ РАЗЛИЧНЫЕ
	ФЛ.КодПоДРФО КАК ДРФО
ИЗ Справочник.ФизическиеЛица КАК ФЛ
ГДЕ ФЛ.КодПоДРФО <> "" И НЕ ФЛ.ПометкаУдаления""", ("ДРФО",)):
    membership.add((drfo or "").strip())

all_keys = set(erp_rows) | set(zup_rows)
matched = set(erp_rows) & set(zup_rows)
flag_keys = {k for k in all_keys if k in membership}

Z4 = (0.0, 0.0, 0.0, 0.0)
sum_d = [0.0] * 4
worst = []
for k in all_keys:
    e = erp_rows.get(k, Z4)
    z = zup_rows.get(k, Z4)
    d = tuple(a - b for a, b in zip(e, z))
    for i in range(4):
        sum_d[i] += d[i]
    worst.append((abs(d[3]), k, d[3]))
worst.sort(reverse=True)

print(f"Строк: всего {len(all_keys)}, matched {len(matched)}, ЕРП-only {len(set(erp_rows) - set(zup_rows))}, ЗУП-only {len(set(zup_rows) - set(erp_rows))}")
print(f"Дубликаты ДРФО в ЕРП (слиты): {dup_erp}")
print(f"С флагом 'Только сотрудники из ЗУП': {len(flag_keys)} строк (скрыто {len(all_keys) - len(flag_keys)})")
print(f"Σ Разница: НачОст={sum_d[0]:,.2f} Начисл={sum_d[1]:,.2f} Выпл={sum_d[2]:,.2f} КонОст={sum_d[3]:,.2f}")

se = [sum(v[i] for v in erp_rows.values()) for i in range(4)]
sz = [sum(v[i] for v in zup_rows.values()) for i in range(4)]
ok_sum = all(abs(se[i] - sz[i] - sum_d[i]) <= TOL for i in range(4))
zup_only_keys = set(zup_rows) - set(erp_rows)
ok_flag = all(k in membership for k in zup_only_keys & flag_keys) and (flag_keys <= membership)
print(f"\nТоп-10 |Δ КонОст|:")
for absd, k, d in worst[:10]:
    print(f"  {k}: {d:,.2f}")

print(f"\nCHECK Σ(ЕРП)-Σ(ЗУП)==Σ(Разница): {'OK' if ok_sum else 'FAIL'}")
print(f"CHECK флаг подмножество членства: {'OK' if ok_flag else 'FAIL'}")
verdict = ok_sum and ok_flag
print("ACCEPTANCE: " + ("PASS" if verdict else "FAIL"))
sys.exit(0 if verdict else 1)
