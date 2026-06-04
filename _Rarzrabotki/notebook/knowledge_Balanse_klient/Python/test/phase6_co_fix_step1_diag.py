# -*- coding: utf-8 -*-
"""
Phase 6 fix — Step 1: Диагностика
1. Прочитать строку ЮЕйДрім в ТЧ ВводОстатков 0Ц-00000083 (Авансы клиентов)
2. Найти/проверить есть ли соответствующий ВводОстатков Задолженности клиентов для ЦО за 31.10.2025
3. Изучить структуру ТЧ обоих типов документов
"""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, money, ARTIFACTS_DIR
import os

erp = connect_erp()
S = erp.String

# === 1. Найти документ Авансов 0Ц-00000083 ===
q = erp.NewObject("Запрос")
q.Текст = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ВводОстатков ГДЕ Номер = "0Ц-00000083" """
sel = q.Выполнить().Выбрать(); sel.Следующий()
docAv = sel.Ссылка
print(f"[1] Документ Авансов: {S(docAv)}")
print(f"    Дата: {docAv.Дата}")
print(f"    ХозОп: {S(docAv.ХозяйственнаяОперация)}")

# Прочитать ТЧ
objAv = docAv.ПолучитьОбъект()
tcAv = objAv.РасчетыСПартнерами
print(f"\n[2] ТЧ.РасчетыСПартнерами содержит {tcAv.Количество()} строк")

# Найти колонки через метаданные
md = objAv.Метаданные()
tcMd = md.ТабличныеЧасти.Найти("РасчетыСПартнерами")
cols = []
for r in tcMd.Реквизиты:
    cols.append(r.Имя)
print(f"    Колонки ТЧ: {cols}")

# Найти строку ЮЕйДрім
print(f"\n[3] Все строки ТЧ:")
print(f"{'#':>3}  {'Партнёр':<35}{'Контрагент':<30}{'Договор':<40}{'Сумма':>14}{'СуммаДолг':>14}{'СуммаАванс':>14}")
print("-" * 152)
yu_row = None
yu_idx = None
for i in range(tcAv.Количество()):
    row = tcAv.Получить(i)
    p = ""
    try: p = str(S(row.Партнер) or "")[:33]
    except: pass
    k = ""
    try: k = str(S(row.Контрагент) or "")[:28]
    except: pass
    d = ""
    try: d = str(S(row.Договор) or "")[:38]
    except: pass
    s = 0
    try: s = float(row.Сумма)
    except: pass
    sd = 0
    try: sd = float(row.СуммаДолга)
    except: pass
    sa = 0
    try: sa = float(row.СуммаПредоплаты)
    except: pass

    print(f"{i+1:>3}  {p:<35}{k:<30}{d:<40}{money(s):>14}{money(sd):>14}{money(sa):>14}")

    if "ЮЕйДрім" in p or "ЮЕйДрім" in k or "13/06/22-1" in d:
        yu_row = row
        yu_idx = i

print(f"\n[4] Найдена строка ЮЕйДрім: индекс={yu_idx}")
if yu_row is None:
    print("[FAIL] Строка не найдена")
    sys.exit(1)

# Полный дамп строки
print(f"\n[5] Полный дамп строки {yu_idx+1}:")
backup_row = {}
for col in cols:
    try:
        v = getattr(yu_row, col)
        if v is None:
            sv = None
        elif isinstance(v, (str, int, float, bool)):
            sv = v
        else:
            try:
                sv_str = S(v) if erp.ЗначениеЗаполнено(v) else ""
                # UUID-форма для восстановления
                uuid_str = ""
                try:
                    if erp.ЗначениеЗаполнено(v):
                        uuid_str = str(S(v.УникальныйИдентификатор()))
                except: pass
                sv = {"name": str(sv_str), "uuid": uuid_str, "type": str(v.Метаданные().ПолноеИмя()) if erp.ЗначениеЗаполнено(v) else ""}
            except Exception as e:
                sv = f"<{type(v).__name__}>"
        backup_row[col] = sv
        if v is not None and v != "" and v != 0:
            print(f"      {col}: {sv}")
    except: pass

# Сохранить backup строки
backup_path = os.path.join(ARTIFACTS_DIR, "phase6_co_yueydrim_row_backup.json")
with open(backup_path, "w", encoding="utf-8") as f:
    json.dump(backup_row, f, ensure_ascii=False, indent=2, default=str)
print(f"\n[6] Backup строки сохранён: {backup_path}")

# === 6. Поиск ВводОстатковЗадолженности для ЦО за 31.10.2025 ===
print(f"\n[7] Поиск документа Ввод остатков задолженности клиентов для ЦО на 31.10.2025:")
q2 = erp.NewObject("Запрос")
q2.Текст = """
ВЫБРАТЬ
    Д.Ссылка КАК Ссылка,
    Д.Номер КАК Номер,
    Д.Дата КАК Дата,
    Д.Проведен КАК Проведен,
    Д.Организация КАК Орг,
    Д.ХозяйственнаяОперация КАК ХозОп
ИЗ Документ.ВводОстатков КАК Д
ГДЕ Д.Организация = ЗНАЧЕНИЕ(Справочник.Организации.ПустаяСсылка)
    ИЛИ Д.Дата = ДАТАВРЕМЯ(2025,10,31,0,0,0)
"""
# Не нужно фильтровать пока — посмотрим что есть
q2.Текст = """
ВЫБРАТЬ
    Д.Ссылка КАК Ссылка,
    Д.Номер КАК Номер,
    Д.Дата КАК Дата,
    Д.Проведен КАК Проведен,
    Д.ХозяйственнаяОперация КАК ХозОп,
    ПРЕДСТАВЛЕНИЕ(Д.Организация) КАК ОргИмя
ИЗ Документ.ВводОстатков КАК Д
ГДЕ Д.Дата = ДАТАВРЕМЯ(2025,10,31,0,0,0)
    И Д.Организация.КодПоЕДРПОУ = "40645273"
УПОРЯДОЧИТЬ ПО Д.Номер
"""
r = q2.Выполнить().Выгрузить()
print(f"    Найдено документов ВводОстатков на 31.10.2025 (ТОВ ІНДАСТРІАЛБУД): {r.Количество()}")
docDolg = None
for i in range(r.Количество()):
    rec = r.Получить(i)
    print(f"      {rec.Номер}  Дата={rec.Дата}  Проведен={rec.Проведен}  ХозОп={S(rec.ХозОп)}")
    if "адолженност" in str(S(rec.ХозОп)).lower() and "клиент" in str(S(rec.ХозОп)).lower():
        docDolg = rec.Ссылка
        print(f"        >>> Этот документ задолженности клиентов")

if docDolg is None:
    print(f"\n[8] Документа ВводОстатковЗадолженностиКлиентов на 31.10.2025 НЕТ — придётся создать новый")
else:
    print(f"\n[8] Документ задолженности существует: {S(docDolg)}")
    # Прочитать его ТЧ
    objD = docDolg.ПолучитьОбъект()
    print(f"    ТЧ.РасчетыСПартнерами: {objD.РасчетыСПартнерами.Количество()} строк")
    tcMdD = objD.Метаданные().ТабличныеЧасти.Найти("РасчетыСПартнерами")
    cols_d = [r.Имя for r in tcMdD.Реквизиты]
    print(f"    Колонки ТЧ Долга: {cols_d}")
