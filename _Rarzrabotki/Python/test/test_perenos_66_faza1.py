"""Verify ФАЗА 1: загрузка .epf через COM + ЗагрузитьИзBuhBud + анализ превью (без создания документов)."""
import sys
from datetime import datetime
from collections import Counter
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Перенос остатков сотрудников.epf"
ДАТА = datetime(2025, 12, 31, 23, 59, 59)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Создать экземпляр внешней обработки (безопасный режим = Ложь, чтобы разрешить COMConnector)
проц = erp.ВнешниеОбработки.Создать(EPF, False)
проц.ДатаОстатков = ДАТА
проц.ЗагрузитьИзBuhBud()

тч = проц.Задолженность
n = тч.Количество()
st = Counter()
sum_all = 0.0
sum_ok = 0.0
neg = 0
for i in range(n):
    стр = тч.Получить(i)
    статус = S(стр.Статус)
    summ = float(стр.Сумма)
    st[статус] += 1
    sum_all += summ
    if статус == "OK":
        sum_ok += summ
    if summ < 0:
        neg += 1

print(f"Превью строк: {n}")
print(f"Статусы: {dict(st)}")
print(f"Σ всех строк: {sum_all:,.2f}")
print(f"Σ OK-строк:   {sum_ok:,.2f}")
print(f"Отрицательных (Кт-Дт<0): {neg}")

# Проверки
assert n > 0, "Превью пусто!"
# Σ всех строк должна совпасть с эталоном BuhBud для орг в балансе (pretest = 204 900,64 на 31.12.2025)
assert abs(sum_all - 204900.64) < 1.0, f"Σ превью {sum_all} != эталон BuhBud 204 900,64"
assert st.get("OK", 0) > 0, "Нет ни одной OK-строки"
print("\n" + "=" * 80)
print(f"✅ ФАЗА 1 OK — превью загружено из 66 (в иерархии), Σ={sum_all:,.2f} = эталон BuhBud, OK={st.get('OK',0)}")
print("=" * 80)
