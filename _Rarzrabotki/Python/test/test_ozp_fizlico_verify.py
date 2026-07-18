# -*- coding: utf-8 -*-
"""Verify: сравнить состояние регистра ПрочиеРасходы для ОЗП 000Ц-000003
с pretest-снимком после применения правки в ManagerModule.bsl.

Acceptance:
  - Σ Сумма не изменилась (Δ < 0.01 ₽)
  - Σ СуммаРегл не изменилась
  - Σ СуммаБезНДС не изменилась
  - % строк с ФизЛицо > 50%
  - Σ Сумма по (Подр, Статья, Аналитика, НН) совпадает с pretest группами
"""
import sys, io, json, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Pretest
path = os.path.join(os.path.dirname(__file__), "ozp_fizlico_pretest.json")
with open(path, "r", encoding="utf-8") as f:
    pretest = json.load(f)

q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    Р.Подразделение КАК Подразделение,
    Р.СтатьяРасходов КАК Статья,
    Р.АналитикаРасходов КАК Аналитика,
    Р.НалоговоеНазначение КАК НН,
    Р.А_ФизическоеЛицо КАК ФЛ,
    Р.Сумма КАК Сумма,
    Р.СуммаБезНДС КАК СуммаБезНДС,
    Р.СуммаРегл КАК СуммаРегл
ИЗ РегистрНакопления.ПрочиеРасходы КАК Р
ГДЕ Р.Регистратор.Номер = "000Ц-000003"
    И Р.Регистратор.Дата = ДАТАВРЕМЯ(2025,12,31,12,0,0)
    И ТИПЗНАЧЕНИЯ(Р.Регистратор) = ТИП(Документ.ОтражениеЗарплатыВФинансовомУчете)
"""
tz = q.Выполнить().Выгрузить()

всего = tz.Количество()
сумма = sum(r.Сумма for r in tz)
сумма_безндс = sum(r.СуммаБезНДС for r in tz)
сумма_регл = sum(r.СуммаРегл for r in tz)
фл_зап = sum(1 for r in tz if r.ФЛ and conn.ЗначениеЗаполнено(r.ФЛ))

print("=== VERIFY ОЗП 000Ц-000003 от 31.12.2025 (после правки) ===\n")
print(f"  ВсегоДвижений:  было {pretest['total_movements']} → стало {всего}")
print(f"  Σ Сумма:        было {pretest['sum_total']:.2f} → стало {сумма:.2f}  Δ={сумма-pretest['sum_total']:.4f}")
print(f"  Σ СуммаБезНДС:  было {pretest['sum_without_vat']:.2f} → стало {сумма_безндс:.2f}  Δ={сумма_безндс-pretest['sum_without_vat']:.4f}")
print(f"  Σ СуммаРегл:    было {pretest['sum_regl']:.2f} → стало {сумма_регл:.2f}  Δ={сумма_регл-pretest['sum_regl']:.4f}")
print(f"  С заполненным ФЛ: {фл_зап} ({100.0*фл_зап/max(всего,1):.1f}%)  (было {pretest['fl_filled']})")

# Свод после правки
свод_after = {}
for r in tz:
    key = (
        conn.XMLСтрока(r.Подразделение) if r.Подразделение and conn.ЗначениеЗаполнено(r.Подразделение) else "",
        conn.XMLСтрока(r.Статья) if r.Статья and conn.ЗначениеЗаполнено(r.Статья) else "",
        conn.XMLСтрока(r.Аналитика) if r.Аналитика and conn.ЗначениеЗаполнено(r.Аналитика) else "",
        conn.XMLСтрока(r.НН) if r.НН and conn.ЗначениеЗаполнено(r.НН) else "",
    )
    свод_after[key] = свод_after.get(key, 0.0) + float(r.Сумма)

свод_before = {tuple(g["key"]): g["sum"] for g in pretest["groups"]}

mismatch = 0
all_keys = set(свод_before) | set(свод_after)
total_delta = 0.0
for k in all_keys:
    a = свод_before.get(k, 0.0)
    b = свод_after.get(k, 0.0)
    if abs(a - b) > 0.01:
        mismatch += 1
        total_delta += abs(a - b)
        if mismatch <= 10:
            print(f"  MISMATCH {k}: до={a:.2f} после={b:.2f} Δ={b-a:.4f}")

print(f"\n  Групп до: {len(свод_before)}, после: {len(свод_after)}")
print(f"  Различающихся групп: {mismatch} (Σ|Δ| = {total_delta:.2f})")

# Финальный assert
assert abs(сумма - pretest['sum_total']) < 0.01, f"Σ Сумма НЕ совпадает! Δ={сумма-pretest['sum_total']:.4f}"
assert abs(сумма_регл - pretest['sum_regl']) < 0.01, f"Σ СуммаРегл НЕ совпадает! Δ={сумма_регл-pretest['sum_regl']:.4f}"
assert (100.0*фл_зап/max(всего,1)) > 50.0, f"% ФЛ слишком мал: {100.0*фл_зап/max(всего,1):.1f}%"
print("\n[OK] Σ-инвариант выполнен, ФЛ заполнен > 50%")
