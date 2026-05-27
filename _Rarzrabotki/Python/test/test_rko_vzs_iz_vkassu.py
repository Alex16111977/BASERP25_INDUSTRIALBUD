"""
Тест: проведение РКО N0000052986 со ссылкой на ВКассу 000Ц-000009 → движения
А_ВзаиморасчетыССотрудниками в разрезе подразделений (СтрПредпр).

Действия:
1. Сменить ВКассу.Касса на "2 Касса Спецтехника" (матч с РКО).
2. Установить РКО.Ведомость = ВКассу.
3. Провести РКО.
4. Проверить движения регистра: Σ=348 800, разрез по СтрПредпр.
"""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EXPECTED_SUM = 348800.00
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)


def find_doc(meta_name, number):
    q = erp.NewObject("Запрос")
    q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка ИЗ Документ.{meta_name} КАК Д ГДЕ Д.Номер = "{number}"'
    rs = q.Выполнить()
    if rs.Пустой(): return None
    sel = rs.Выбрать(); sel.Следующий()
    return sel.Ссылка


# 1. Эталоны
rko_ref = find_doc("РасходныйКассовыйОрдер", "N0000052986")
vk_ref  = find_doc("ВедомостьНаВыплатуЗарплатыВКассу", "000Ц-000009")
if rko_ref is None or vk_ref is None:
    print("FAIL: эталоны не найдены")
    sys.exit(1)
print(f"РКО:    {rko_ref}")
print(f"ВКассу: {vk_ref}")

# 2. Сменить Кассу в ВКассу
kassa_target = erp.Справочники.Кассы.НайтиПоНаименованию("2 Касса Спецтехника", True)
if not kassa_target or kassa_target.Пустая():
    print("FAIL: касса '2 Касса Спецтехника' не найдена")
    sys.exit(1)

vk = vk_ref.ПолучитьОбъект()
kassa_old = vk.Касса
if (kassa_old and not kassa_old.Пустая() and
    erp.string(kassa_old.УникальныйИдентификатор()) != erp.string(kassa_target.УникальныйИдентификатор())):
    vk.Касса = kassa_target
    vk.ОбменДанными.Загрузка = True
    vk.Записать(erp.РежимЗаписиДокумента.Запись)
    print(f"OK: ВКассу.Касса '{kassa_old.Наименование}' → '{kassa_target.Наименование}'")
else:
    print(f"OK: ВКассу.Касса уже = '{kassa_target.Наименование}'")

# 3. Установить РКО.Ведомость и провести
rko = rko_ref.ПолучитьОбъект()
# Если РКО уже проведено — распровести
if rko.Проведен:
    rko.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
rko = rko_ref.ПолучитьОбъект()
rko.Ведомость = vk_ref
try:
    rko.Записать(erp.РежимЗаписиДокумента.Проведение)
    print(f"OK: РКО проведён (Ведомость = ВКассу)")
except Exception as e:
    info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
    print(f"FAIL: проведение упало: {info}")
    sys.exit(1)

# 4. Проверить движения регистра
q = erp.NewObject("Запрос")
q.Текст = '''ВЫБРАТЬ
	Р.Подразделение.Наименование КАК СтрПредпр,
	Р.ФизическоеЛицо.Наименование КАК ФЛ,
	Р.СуммаВзаиморасчетов КАК Сум
ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками КАК Р
ГДЕ Р.Регистратор = &Док
УПОРЯДОЧИТЬ ПО Р.Подразделение, Р.ФизическоеЛицо'''
q.УстановитьПараметр("Док", rko_ref)
rs = q.Выполнить()

cnt = 0
sum_total = 0.0
under_lines = {}  # СтрПредпр → Σ
sel = rs.Выбрать()
while sel.Следующий():
    cnt += 1
    s = float(sel.Сум)
    sum_total += s
    spp_name = str(sel.СтрПредпр) if sel.СтрПредпр else "—"
    under_lines.setdefault(spp_name, 0.0)
    under_lines[spp_name] += s

print(f"\nДвижения регистра А_ВзаиморасчетыССотрудниками:")
print(f"  Кол строк:        {cnt}")
print(f"  Σ:                {sum_total:,.2f}")
print(f"  Разрез по СтрПредпр:")
for spp, s in sorted(under_lines.items()):
    print(f"    {spp}: Σ={s:,.2f}")

fail = False
if abs(sum_total - EXPECTED_SUM) > 0.01:
    print(f"FAIL: Σ движений {sum_total:,.2f} != ожидаемое {EXPECTED_SUM:,.2f}")
    fail = True
if cnt < 2:
    print(f"FAIL: ожидали ≥2 строки (разрез по СтрПредпр), получили {cnt}")
    fail = True
if len(under_lines) < 2:
    print(f"FAIL: ожидали ≥2 уникальных СтрПредпр (разрез), получили {len(under_lines)}: {list(under_lines.keys())}")
    fail = True

if fail:
    sys.exit(1)
print(f"\nPASS: движения А_ВзСС корректны (Σ={sum_total:,.2f}, {cnt} строк, {len(under_lines)} подразделений)")
