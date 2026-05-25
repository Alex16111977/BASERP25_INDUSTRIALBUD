# -*- coding: utf-8 -*-
"""
TEST 2 — Перепроведение 2-х малых документов через ПерепровестиДокументы().

Цель: проверить что Записать(Проведение) внутри COM-обработки корректно
создаёт РегистраторРасчётов в РСППС (защита от feedback_com_repost_skips_registrator_raschetov).

Метод:
  1. Анализ → ТЧ заполнена
  2. Очистить ТЧ кроме 2 малых:
     - Списание 000005683 от 29.04.2026 (Логистика, Δ=-780)
     - Списание 000005519 от 24.04.2026 (МД ПРООН, Δ=-1959.32)
  3. Запомнить baseline РегистраторРасчётов для них
  4. Вызвать ПерепровестиДокументы()
  5. Проверить:
     - Обоих Обработан=Истина и НовоеСостояние="ОК"
     - В РСППС РегистраторРасчётов снова есть (не пустой)
     - Δ ≈ 0

Acceptance: 2 документа Обработан=Истина с Δ=0, РегистраторРасчётов сохранён.
"""
import sys, io, datetime as dt
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp

ERF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ОбработкаДисбалансаПоСтатьямБаланса.erf"

erp = connect_erp()
report = erp.ВнешниеОтчеты.Создать(ERF_PATH, False)

report.НачалоПериода = dt.datetime(2026, 4, 1)
report.ОкончаниеПериода = dt.datetime(2026, 4, 30, 23, 59, 59)
report.АнализРасхождений()

tch = report.ДокументыРасхождения
print(f"После Анализа: {tch.Количество()} строк ТЧ")

# Оставить только 2 малых СписанияБезн ДС
# 000005683 — Логистика 780, 000005519 — МД ПРООН 1959.32
TARGET_NUMBERS = {"000005683", "000005519"}
to_delete = []
for i in range(tch.Количество()):
    row = tch.Получить(i)
    doc_name = str(erp.String(row.Документ)) if row.Документ else ""
    # Проверяем что это Списание и номер в TARGET
    is_spisanie = "Списание" in doc_name
    has_target_num = any(num in doc_name for num in TARGET_NUMBERS)
    keep = is_spisanie and has_target_num
    if not keep:
        to_delete.append(i)

# Удаляем в обратном порядке
for idx in reversed(to_delete):
    tch.Удалить(idx)

print(f"После фильтра: {tch.Количество()} строк (ожидали 4: 2 документа × 2 статьи)")
# Каждое Списание даёт 2 строки в ТЧ: одну по статье ЗПП и одну по ВыданныеАвансы
# (документ списывает аванс + не пишет долг → плуг на обе статьи)
unique_docs = set()
for i in range(tch.Количество()):
    row = tch.Получить(i)
    unique_docs.add(str(erp.String(row.Документ)))

print(f"Уникальных документов в ТЧ: {len(unique_docs)} (ожидали 2)")
for doc in sorted(unique_docs):
    print(f"  - {doc}")
if len(unique_docs) != 2:
    print("FAIL: ожидали 2 уникальных документа")
    sys.exit(1)

# Baseline РегистраторРасчётов
def check_registrator_raschetov(doc_ref):
    """Возвращает количество строк в РСППС с этим ДокументРегистратор."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Док", doc_ref)
    q.Текст = """
    ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол
    ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам
    ГДЕ ДокументРегистратор = &Док
    """
    sel = q.Выполнить().Выбрать()
    sel.Следующий()
    return int(sel.Кол)

print("\n=== Baseline РегистраторРасчётов (до перепроведения) ===")
for i in range(tch.Количество()):
    row = tch.Получить(i)
    cnt = check_registrator_raschetov(row.Документ)
    print(f"  {erp.String(row.Документ)}: РСППС-строк = {cnt}")

# Перепровести
print("\n=== Запуск ПерепровестиДокументы() ===")
report.ПерепровестиДокументы()

# Проверка
print("\n=== Acceptance ===")
errors = 0
for i in range(tch.Количество()):
    row = tch.Получить(i)
    doc_name = str(erp.String(row.Документ))
    status = "OK" if row.Обработан and abs(row.Дельта) < 0.01 else "FAIL"
    print(f"  {status} {doc_name}: Обработан={row.Обработан}, Δ={row.Дельта:+.2f}, состояние='{row.НовоеСостояние}'")
    if row.ОшибкаТекст:
        print(f"    Ошибка: {row.ОшибкаТекст}")
    if status == "FAIL":
        errors += 1

    # Проверка РегистраторРасчётов после
    cnt_after = check_registrator_raschetov(row.Документ)
    if cnt_after == 0:
        print(f"  FAIL РегистраторРасчётов: {doc_name} — РСППС пустой после перепроведения!")
        print(f"  ⚠️ КРИТИЧНО: проявилась memory feedback_com_repost_skips_registrator_raschetov")
        errors += 1
    else:
        print(f"  OK РегистраторРасчётов: {doc_name} — РСППС-строк = {cnt_after}")

if errors == 0:
    print("\n*** TEST 2 PASS — перепроведение через COM-обработку работает корректно ***")
    print("*** Можно запускать Test 3 на полную выборку ***")
    sys.exit(0)
else:
    print(f"\n*** TEST 2 FAIL ({errors} errors) — STOP, не запускать Test 3 ***")
    sys.exit(1)
