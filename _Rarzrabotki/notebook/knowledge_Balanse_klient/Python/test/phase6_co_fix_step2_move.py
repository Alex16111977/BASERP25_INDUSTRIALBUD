# -*- coding: utf-8 -*-
"""
Phase 6 fix — Step 2: Перенос строки ЮЕйДрім из Авансов в Долги
1. Backup полных ТЧ обоих документов в JSON
2. Добавить строку в 0Ц-00000082 (Долги) с теми же значениями
3. Удалить строку #1 из 0Ц-00000083 (Авансы)
4. Записать оба документа с проведением
5. Проверить что РСКПС за партнёра ЮЕйДрім теперь имеет ДолгУпр=+36408 (с положительным знаком как долг клиента)
"""
import sys, io, json, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, money, ARTIFACTS_DIR

erp = connect_erp()
S = erp.String

def dump_tc(obj, name):
    """Дамп ТЧ в список словарей."""
    tc = getattr(obj, name)
    md_tc = obj.Метаданные().ТабличныеЧасти.Найти(name)
    cols = [r.Имя for r in md_tc.Реквизиты]
    rows = []
    for i in range(tc.Количество()):
        row = tc.Получить(i)
        d = {}
        for col in cols:
            try:
                v = getattr(row, col)
                if v is None:
                    d[col] = None
                elif isinstance(v, (str, int, float, bool)):
                    d[col] = v
                else:
                    try:
                        sv = S(v) if erp.ЗначениеЗаполнено(v) else ""
                        uuid = ""
                        try:
                            if erp.ЗначениеЗаполнено(v):
                                uuid = str(S(v.УникальныйИдентификатор()))
                        except: pass
                        d[col] = {"name": str(sv), "uuid": uuid}
                    except:
                        d[col] = "<obj>"
            except: pass
        rows.append(d)
    return cols, rows

# === Найти оба документа ===
q = erp.NewObject("Запрос")
q.Текст = """ВЫБРАТЬ Ссылка ИЗ Документ.ВводОстатков ГДЕ Номер = "0Ц-00000083" """
sel = q.Выполнить().Выбрать(); sel.Следующий()
refAv = sel.Ссылка

q.Текст = """ВЫБРАТЬ Ссылка ИЗ Документ.ВводОстатков ГДЕ Номер = "0Ц-00000082" """
sel = q.Выполнить().Выбрать(); sel.Следующий()
refDol = sel.Ссылка

print(f"Авансы: {S(refAv)}")
print(f"Долги:  {S(refDol)}")

objAv = refAv.ПолучитьОбъект()
objDol = refDol.ПолучитьОбъект()

# === Backup ===
print("\n[1] Backup ТЧ обоих документов")
cols_av, rows_av = dump_tc(objAv, "РасчетыСПартнерами")
cols_dol, rows_dol = dump_tc(objDol, "РасчетыСПартнерами")
backup = {
    "Avansy_0C_00000083": {"cols": cols_av, "count": len(rows_av), "rows": rows_av},
    "Dolgi_0C_00000082":  {"cols": cols_dol, "count": len(rows_dol), "rows": rows_dol},
}
backup_path = os.path.join(ARTIFACTS_DIR, "phase6_co_backup_before_move.json")
with open(backup_path, "w", encoding="utf-8") as f:
    json.dump(backup, f, ensure_ascii=False, indent=2, default=str)
print(f"    Backup: {backup_path}")
print(f"    Авансы:  {len(rows_av)} строк")
print(f"    Долги:   {len(rows_dol)} строк")

# === Найти строку ЮЕйДрім в Авансах ===
print("\n[2] Поиск строки ЮЕйДрім в Авансах")
tcAv = objAv.РасчетыСПартнерами
src_row = None
src_idx = None
for i in range(tcAv.Количество()):
    row = tcAv.Получить(i)
    try:
        p = str(S(row.Партнер) or "")
    except: p = ""
    if "ЮЕйДрім" in p:
        src_row = row
        src_idx = i
        break
if src_row is None:
    print("[FAIL] Строка не найдена"); sys.exit(1)
print(f"    Индекс: {src_idx}, Партнёр={p}, Сумма={src_row.Сумма}")

# Сохранить значения для копирования
src_values = {}
for col in cols_av:
    try:
        src_values[col] = getattr(src_row, col)
    except: pass

# === Проверить нет ли уже такой строки в Долгах ===
print("\n[3] Проверка дубликата в Долгах")
tcDol = objDol.РасчетыСПартнерами
dup = False
for i in range(tcDol.Количество()):
    row = tcDol.Получить(i)
    try:
        p = str(S(row.Партнер) or "")
        d = str(S(row.Договор) or "")
    except: continue
    if "ЮЕйДрім" in p and "13/06/22-1" in d:
        dup = True
        print(f"    [!] УЖЕ ЕСТЬ строка в Долгах #{i+1}: Партнёр={p}, Договор={d}, Сумма={row.Сумма}")
        break
if dup:
    print("    [STOP] Дубль — прерываю операцию для ручной проверки")
    sys.exit(2)
print("    Дубликата нет, продолжаю")

# === Добавить строку в Долги ===
print("\n[4] Добавление строки в Долги")
new_row = objDol.РасчетыСПартнерами.Добавить()
for col, v in src_values.items():
    try:
        setattr(new_row, col, v)
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        print(f"    [warn] не удалось установить {col}: {info[2] if info else e}")
print(f"    Добавлена строка с Партнёр={S(new_row.Партнер)}, Сумма={new_row.Сумма}")
print(f"    Долги теперь: {objDol.РасчетыСПартнерами.Количество()} строк")

# === Удалить строку из Авансов ===
print("\n[5] Удаление строки из Авансов")
objAv.РасчетыСПартнерами.Удалить(src_idx)
print(f"    Авансы теперь: {objAv.РасчетыСПартнерами.Количество()} строк")

# === Записать оба документа с проведением ===
print("\n[6] Запись документов")
try:
    objDol.Записать(erp.РежимЗаписиДокумента.Проведение)
    print(f"    [OK] Долги перепроведены")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] Долги: {info[2] if info else e}")
    sys.exit(3)

try:
    objAv.Записать(erp.РежимЗаписиДокумента.Проведение)
    print(f"    [OK] Авансы перепроведены")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] Авансы: {info[2] if info else e}")
    sys.exit(4)

# === Verify: РСКПС за ЮЕйДрім ===
print("\n[7] Verify РСКПС по ЮЕйДрім (по обоим документам, дата 31.10.2025)")
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Док1", refAv)
q.УстановитьПараметр("Док2", refDol)
q.Текст = """
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор) КАК Док,
    ПРЕДСТАВЛЕНИЕ(АП.Партнер) КАК Партнёр,
    Р.ВидДвижения КАК ВидДв,
    Р.ДолгУпр КАК ДолгУпр,
    Р.ПредоплатаУпр КАК Аванс,
    Р.Период КАК Период
ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам КАК Р
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Р.ДокументРегистратор В (&Док1, &Док2)
    И АП.Партнер.Наименование ПОДОБНО "%ЮЕйДрім%"
УПОРЯДОЧИТЬ ПО Период
"""
r = q.Выполнить().Выгрузить()
print(f"    Строк РСКПС по ЮЕйДрім: {r.Количество()}")
for i in range(r.Количество()):
    rec = r.Получить(i)
    print(f"    {str(rec.Док)[:50]:<52}  {S(rec.ВидДв):<8}  Долг={rec.ДолгУпр:>+12,.2f}  Аванс={rec.Аванс:>+12,.2f}")

print("\n[8] Шаг 2 выполнен. Далее — перепровести А_ФинРез_Баланс декабря.")
