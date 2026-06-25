import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ERF = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\stupefied-neumann-346fd7\_build_pfpp\А_ПланФактныйПроизводствоПолный.erf"
СС_ИМЯ = "МД МХП ОРІЛЬ / СТІЛ МД МХП ОРІЛЬ"

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

rep = erp.ВнешниеОтчеты.Создать(ERF)
сс = erp.Справочники.А_СтруктураСебестоимости.НайтиПоНаименованию(СС_ИМЯ)
assert not сс.Пустая(), "СС не найдена"

период = datetime.datetime(2026, 4, 30)
д = rep.ПолучитьДанныеАнализа(период, сс)

def g(k):
    return д.Получить(k) if hasattr(д, "Получить") else getattr(д, k)

# структура 1С: доступ через свойство по ключу
def val(k):
    try:
        return getattr(д, k)
    except Exception:
        return д.Свойство(k)

print("=== ПЛАН (контроль Оріль) ===")
checks = {
    "РаботыГрн": 753000,
    "ЧасВир": 1927,
    "ЧасМонт": 868,
    "СтавкаВир": 200,
    "СтавкаМонт": 350,
    "ВирГрнПлан": 385400,
    "МонтГрнПлан": 303800,  # 868ч×350 (карта); Excel 302400 исключал 4ч фикс-стоимости (864ч)
}
ok = True
for k, exp in checks.items():
    got = round(float(val(k)), 2)
    flag = "OK" if abs(got - exp) < 0.5 else "FAIL"
    if flag == "FAIL":
        ok = False
    print(f"  {k:14} = {got:>14}  (эталон {exp})  {flag}")

мв = round(float(val("МатВир")), 2)
мм = round(float(val("МатМонт")), 2)
print(f"  {'МатВир':14} = {мв:>14}  (эталон ~1711992)  {'OK' if abs(мв-1711992)<2 else 'FAIL'}")
print(f"  {'МатМонт':14} = {мм:>14}  (эталон 206605)  {'OK' if abs(мм-206605)<0.5 else 'FAIL'}")

print("\n=== ФАКТ (информативно, период <= 30.04.2026) ===")
for k in ["ФактЧасПроект", "ФактЧасМесяц", "ЗПнарахПроект", "ЗПнарахМесяц",
          "ВирГрнФактПроект", "МатВирФакт", "МатМонтФакт", "СумГрнФактПроект"]:
    print(f"  {k:18} = {round(float(val(k)),2)}")

print("\n=== РЕНДЕР ТабличногоДокумента ===")
тд = erp.NewObject("ТабличныйДокумент")
try:
    rep.СформироватьТабличныйДокумент(д, тд)
except Exception as e:
    info = getattr(e, "excepinfo", None)
    msg = info[2] if info and len(info) > 2 else str(e)
    with open(r"C:\Configuration_downloads\BASERP25\.claude\worktrees\stupefied-neumann-346fd7\_pfpp_err.txt", "w", encoding="utf-8") as f:
        f.write(str(msg))
    print("  RENDER ERROR written to _pfpp_err.txt")
    raise
print(f"  ВысотаТаблицы = {тд.ВысотаТаблицы}, ШиринаТаблицы = {тд.ШиринаТаблицы}")
assert тд.ВысотаТаблицы > 3, "пустой документ"
assert тд.ШиринаТаблицы >= 51, "недостаточно колонок"

print("\nИТОГ:", "SMOKE OK" if ok and abs(мв-1711992)<2 and abs(мм-206605)<0.5 else "SMOKE FAIL")
