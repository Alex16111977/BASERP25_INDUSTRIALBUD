import sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
kazna = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')

# === V4: остатки/обороты наличной кассы Казны (спека §4.2) ===
# Фильтр Наличные внутри 5-го параметра виртуальной таблицы ОстаткиИОбороты
QUERY = """
ВЫБРАТЬ
    ДС.БанковскийСчетКасса КАК БанковскийСчетКасса,
    ДС.СуммаНачальныйОстаток КАК СуммаНачальныйОстаток,
    ДС.СуммаПриход КАК СуммаПриход,
    ДС.СуммаРасход КАК СуммаРасход,
    ДС.СуммаКонечныйОстаток КАК СуммаКонечныйОстаток
ИЗ
    РегистрНакопления.ДенежныеСредства.ОстаткиИОбороты(ДАТАВРЕМЯ(2025,12,1,0,0,0), ДАТАВРЕМЯ(2025,12,31,23,59,59), , ,
        ВидДенежныхСредств = ЗНАЧЕНИЕ(Перечисление.ВидыДенежныхСредств.Наличные)) КАК ДС
"""

def run(period_label, query_text):
    q = kazna.NewObject("Запрос")
    q.Text = query_text
    try:
        r = q.Execute().Выгрузить()
    except Exception as e:
        if hasattr(e, 'excepinfo') and e.excepinfo:
            print(f"[{period_label}] FAIL: {e.excepinfo[2]}")
        else:
            print(f"[{period_label}] FAIL: {e}")
        return None
    n = r.Количество()
    print(f"[{period_label}] OK, строк (касс) = {n}")
    shown = 0
    for i in range(n):
        row = r.Получить(i)
        ref = row.БанковскийСчетКасса
        if kazna.ЗначениеЗаполнено(ref):
            kod = kazna.String(ref.Код)
            naim = kazna.String(ref.Наименование)
        else:
            kod = "<пусто>"
            naim = "<пусто>"
        nach = row.СуммаНачальныйОстаток
        prih = row.СуммаПриход
        rash = row.СуммаРасход
        kon = row.СуммаКонечныйОстаток
        print(f"  Код={kod} | Наим={naim} | Нач={nach} Прих={prih} Расх={rash} Кон={kon}")
        shown += 1
        if shown >= 8:
            print(f"  ... (показано {shown} из {n})")
            break
    return r

print("=== ДЕКАБРЬ 2025 (фильтр Наличные в 5-м параметре ОстаткиИОбороты) ===")
res = run("дек2025", QUERY)

# Если декабрь пуст — пробуем январь 2026
if res is None or res.Количество() == 0:
    print()
    print("=== ЯНВАРЬ 2026 (декабрь пуст/упал — fallback) ===")
    QUERY_JAN = QUERY.replace("ДАТАВРЕМЯ(2025,12,1,0,0,0)", "ДАТАВРЕМЯ(2026,1,1,0,0,0)").replace(
        "ДАТАВРЕМЯ(2025,12,31,23,59,59)", "ДАТАВРЕМЯ(2026,1,31,23,59,59)")
    run("янв2026", QUERY_JAN)
