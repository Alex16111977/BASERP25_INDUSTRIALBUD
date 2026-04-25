"""Найти регистры и отчёты с подразделением Астарта.Тищенки (декабрь 2025)."""
import sys, datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

conn = connect_erp()
md = conn.Метаданные

# 1) перебрать все регистры накопления с измерением Подразделение и проверить
podr_ref = conn.Справочники.СтруктураПредприятия.НайтиПоНаименованию("Астарта. Тищенки", True)
print("Подразделение ref OK:", podr_ref.Наименование)

print("\n=== Регистры накопления с измерением Подразделение ===")
for r in md.РегистрыНакопления:
    has_podr = False
    for d in r.Измерения:
        if str(d.Имя) == "Подразделение":
            has_podr = True
            break
    if not has_podr:
        continue
    # проверим есть ли данные по Астарте в декабре 2025
    q = conn.NewObject("Запрос")
    q.Текст = f"ВЫБРАТЬ ПЕРВЫЕ 1 1 КАК К ИЗ РегистрНакопления.{r.Имя} ГДЕ Подразделение = &П И Период МЕЖДУ &С И &ПО"
    q.УстановитьПараметр("П", podr_ref)
    q.УстановитьПараметр("С", datetime.datetime(2025, 12, 1))
    q.УстановитьПараметр("ПО", datetime.datetime(2025, 12, 31, 23, 59, 59))
    try:
        has = q.Выполнить().Пустой()
        if not has:
            print(f"  ✓ {r.Имя} — ЕСТЬ ДАННЫЕ")
    except Exception as e:
        pass

print("\n=== Отчеты с 'ДоходыИРасходы' ===")
for r in md.Отчеты:
    nm = str(r.Имя)
    if "Доход" in nm or "Расход" in nm:
        print(f"  - {nm} ({r.Синоним})")
