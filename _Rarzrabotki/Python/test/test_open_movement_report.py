# -*- coding: utf-8 -*-
"""
Тест: проверка что Отчет.ДвиженияДокумента принимает наш документ как параметр.
"""
import sys, io
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
)

q = erp.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
ИЗ Документ.А_ПереносДвиженийИзКазны
ГДЕ Проведен И НЕ ПометкаУдаления
"""
sel = q.Выполнить().Выбрать()
sel.Следующий()
ref = sel.Ссылка

print("Документ найден:", str(ref))
print("Тип:", type(ref).__name__)

# Создаём отчёт и пытаемся применить параметр через тот же путь, что использует ПриСозданииНаСервере
report = erp.Отчеты.ДвиженияДокумента.Создать()
print("\nОтчёт создан")

# Параметр СКД "ДокументВладелец"
nastroyki = report.КомпоновщикНастроек.Настройки
parametr_id = erp.NewObject("ПараметрКомпоновкиДанных", "ДокументВладелец")
param = nastroyki.Параметры.НайтиЗначениеПараметра(parametr_id)
if param is None:
    print("[FAIL] Параметр ДокументВладелец не найден в СКД!")
    print("Доступные:")
    for p in nastroyki.Параметры:
        print(" -", str(p.Параметр))
    sys.exit(1)

param.Значение = ref
param.Использование = True
print(f"[OK] Параметр установлен: {ref}")

# Сформировать
result = erp.NewObject("ТабличныйДокумент")
try:
    компМакет = erp.NewObject("КомпоновщикМакетаКомпоновкиДанных")
    макет = компМакет.Выполнить(report.СхемаКомпоновкиДанных, nastroyki)
    proc = erp.NewObject("ПроцессорКомпоновкиДанных")
    proc.Инициализировать(макет, None, None, True)
    output = erp.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
    output.УстановитьДокумент(result)
    output.Вывести(proc, True)
    print(f"\n[OK] Отчёт сформирован: {result.ВысотаТаблицы} строк")
except Exception as e:
    print(f"[FAIL] {e}")
