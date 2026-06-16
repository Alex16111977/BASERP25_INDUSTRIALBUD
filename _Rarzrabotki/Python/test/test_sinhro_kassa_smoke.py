# -*- coding: utf-8 -*-
# Smoke-тест обробки СинхронизироватьДеньгиКасса: Фаза 1 (сверка) + Фаза 2 (аналіз) + фільтр каси.
# Фаза 3 (перепроведення/обмін) НЕ виконується (тільки UI користувачем).
import sys, datetime
import win32com.client
import pywintypes
sys.stdout.reconfigure(encoding="utf-8")

EPF = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\trusting-noyce-fbeddb\_Rarzrabotki\Обработки\СинхронизироватьДеньгиКасса.epf"
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

def err(e):
    if hasattr(e, "excepinfo") and e.excepinfo:
        return str(e.excepinfo[2])
    return str(e)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)
print("Підключено до BaseERP")

try:
    proc = erp.ВнешниеОбработки.Создать(EPF, False)
except Exception:
    proc = erp.ВнешниеОбработки.Создать(EPF)
print("Обробку завантажено")

# Дату формуємо СЕРВЕРНО через англ. алиас erp.Date (без tz-зсуву pywintypes)
try:
    d1 = erp.Date(2025, 12, 1, 0, 0, 0)
    d2 = erp.Date(2025, 12, 31, 0, 0, 0)
except Exception as e:
    print("erp.Date недоступне, fallback pywintypes (можливий tz-зсув):", err(e))
    d1 = pywintypes.Time(datetime.datetime(2025, 12, 1, 12, 0, 0))
    d2 = pywintypes.Time(datetime.datetime(2025, 12, 31, 12, 0, 0))

proc.НачалоПериода = d1
proc.ОкончаниеПериода = d2
print("Перевірка реквізитів періоду: Начало=%s Окончание=%s" % (proc.НачалоПериода, proc.ОкончаниеПериода))

# ---- ФАЗА 1: сверка без фільтра ----
print("\n=== ФАЗА 1: СравнитьОстатки (грудень 2025, без фільтра) ===")
try:
    res = proc.СравнитьОстатки()
    print("Результат:", erp.String(res))
except Exception as e:
    print("FAIL СравнитьОстатки:", err(e)); sys.exit(1)

тр = proc.ТаблицаРасхождений
print("Рядків розбіжностей:", тр.Количество())
top_uid = ""
top_kassa = None
for i in range(min(тр.Количество(), 8)):
    s = тр.Получить(i)
    print(f"  [{i}] {erp.String(s.Касса)[:30]:30} | НачЕРП={float(s.НачОстатокЕРП):>13.2f} НачКазна={float(s.НачОстатокКазна):>13.2f}"
          f" | ПрихЕРП={float(s.ПриходЕРП):>12.2f} ПрихКазна={float(s.ПриходКазна):>12.2f}"
          f" | РасхЕРП={float(s.РасходЕРП):>12.2f} РасхКазна={float(s.РасходКазна):>12.2f} | Δ={float(s.Разница):>12.2f}")
    if i == 0:
        top_uid = erp.String(s.КассаUID)
        top_kassa = s.Касса

# ---- ФАЗА 2: аналіз верхнього рядка ----
print("\n=== ФАЗА 2: АнализироватьДокументы(0) ===")
if тр.Количество() > 0:
    try:
        res2 = proc.АнализироватьДокументы(0)
        print("Результат:", erp.String(res2))
        тд = proc.ТаблицаДокументов
        print("Рядків документів:", тд.Количество())
        for i in range(min(тд.Количество(), 14)):
            s = тд.Получить(i)
            print(f"  [{i}] {erp.String(s.ТипДокумента)[:24]:24} | СумЕРП={float(s.СуммаЕРП):>11.2f} СумКазна={float(s.СуммаКазна):>11.2f}"
                  f" | Дія='{erp.String(s.Действие)}' | {erp.String(s.Статус)[:48]}")
    except Exception as e:
        print("FAIL АнализироватьДокументы:", err(e))
else:
    print("Немає розбіжностей для аналізу")

# ---- ФІЛЬТР каси: повторна сверка тільки по верхній касі ----
print("\n=== ФАЗА 1 з ФільтрКасса (валідація гілки &БезФильтраКассы) ===")
if top_uid:  # КассаUID непорожній → каса є посиланням ЕРП
    try:
        proc.ФильтрКасса = top_kassa
        res3 = proc.СравнитьОстатки()
        print("Результат:", erp.String(res3))
        тр2 = proc.ТаблицаРасхождений
        print("Рядків (очікується <=1 по фільтрованій касі):", тр2.Количество())
        for i in range(тр2.Количество()):
            s = тр2.Получить(i)
            print(f"  [{i}] {erp.String(s.Касса)[:30]:30} | Δ={float(s.Разница):>12.2f}")
    except Exception as e:
        print("FAIL фільтр:", err(e))
else:
    print("Верхня каса не зіставлена з ЕРП (тільки в Казні) — фільтр-тест пропущено")

print("\nГОТОВО smoke (Фаза 3 не виконувалась — тільки UI).")
