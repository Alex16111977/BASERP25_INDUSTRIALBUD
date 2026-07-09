import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПланФактныйПроизводствоПолный.erf"
erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
rep = erp.ВнешниеОтчеты.Создать(ERF)

from collections import defaultdict
def _пуст(ссылка):
    try: return ссылка.Пустая()
    except Exception: return ссылка is None

def проверить(имя, код=None):
    if код:
        сс = erp.Справочники.А_СтруктураСебестоимости.НайтиПоКоду(код)   # дубли имён → по коду
    else:
        сс = erp.Справочники.А_СтруктураСебестоимости.НайтиПоНаименованию(имя)
    assert not сс.Пустая(), имя
    тз = rep.ПолучитьДанные(None, сс.Ссылка)   # Период=Неопределено → весь срок
    print(f"\n=== {имя} (код {сс.Код}): строк ТЗ = {тз.Количество()} ===")
    agg = defaultdict(lambda: [0.0,0.0,0.0,0.0])
    matfact_byэтап = defaultdict(float); matfact_tot=0.0; matfact_nf=0.0
    for i in range(тз.Количество()):
        r = тз.Получить(i)
        k = (str(r.Блок), str(r.Категория))
        agg[k][0]+=float(r.ПланЧасы); agg[k][1]+=float(r.ФактЧасы)
        agg[k][2]+=float(r.ПланГрн);  agg[k][3]+=float(r.ФактГрн)
        if str(r.Блок).startswith("2") and float(r.ФактГрн)>0:
            этап = "" if _пуст(r.Этап) else str(r.Этап)
            matfact_byэтап[этап]+=float(r.ФактГрн); matfact_tot+=float(r.ФактГрн)
            if этап=="": matfact_nf+=float(r.ФактГрн)
    for k in sorted(agg):
        v=agg[k]
        print(f"  {k[0]:14}/{k[1]:12} ПланЧ={v[0]:8.0f} ФактЧ={v[1]:8.0f} ПланГрн={v[2]:12.0f} ФактГрн={v[3]:12.0f}")
    if matfact_tot>0:
        print(f"  -- факт-матеріали по етапах (ФактГрн={matfact_tot:.0f}):")
        for et,c in sorted(matfact_byэтап.items(), key=lambda x:-x[1])[:9]:
            print(f"     {(et if et else '(без етапу)'):24} {c:14.0f}")
        print(f"  -- разнесено по этапам: {(matfact_tot-matfact_nf)/matfact_tot*100:.1f}% ({matfact_tot-matfact_nf:.0f}/{matfact_tot:.0f})")
    return matfact_tot, matfact_nf

mt, mnf = проверить("МД ПРООН Черкаси  ДСНС / СТІЛ Цех МД ПРООН", код="000000003")
проверить("МД МХП ОРІЛЬ / СТІЛ МД МХП ОРІЛЬ")
print("\nИТОГ:", "OK" if (mt==0 or (mt-mnf)/mt>0.8) else "FAIL")
