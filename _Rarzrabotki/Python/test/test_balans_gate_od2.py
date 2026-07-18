# -*- coding: utf-8 -*-
"""GATE OD-2: чи заповнюється ИдентификаторФинЗаписи у ПАП (ключ JOIN Етапів 2-4).
Read-only. PASS і FAIL обидва валідні — це ворота, не тест на «зелене».
FAIL → ескалація фінансисту, БЕЗ silent-fallback (канон OD-2 Вимога 1)."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom
from datetime import datetime

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
D = datetime(2026, 1, 31, 12, 0, 0)   # зріз на кінець січня, опівдні (TZ-safe)
pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(CONN)

GATE_PASS = False
try:
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ
      КОЛИЧЕСТВО(*) КАК Всего,
      СУММА(ВЫБОР КОГДА Б.ИдентификаторФинЗаписи <> "" ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Заполнено
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Остатки(&Д) КАК Б
    """
    q.УстановитьПараметр("Д", D)
    r = q.Выполнить().Выбрать(); r.Следующий()
    vsego = int(r.Всего); zap = int(r.Заполнено or 0)
    print(f"ПАП.Остатки(31.01.2026): всього={vsego}, з ИдентификаторФинЗаписи<>''={zap}")
    GATE_PASS = vsego > 0 and zap > 0 and (zap / max(vsego, 1)) > 0.5
except Exception as e:
    msg = ""
    if hasattr(e, "excepinfo") and e.excepinfo:
        msg = str(e.excepinfo[2])
    print(f"ПАП.ИдентификаторФинЗаписи — запит НЕ пройшов: {msg or e}")
    GATE_PASS = False

print("=" * 64)
if GATE_PASS:
    print("GATE OD-2: PASS — ИдентификаторФинЗаписи придатний як ключ JOIN.")
    print("Етапи 2-4 РОЗБЛОКОВАНО. Наступний крок — окремий план stages 2-4")
    print("(writing-plans з канону §8, ключ = ИдентификаторФинЗаписи).")
else:
    print("GATE OD-2: FAIL — ИдентификаторФинЗаписи НЕ придатний як ключ JOIN.")
    print("Етапи 2-4 ЗАБЛОКОВАНО (канон OD-2 Вимога 1). БЕЗ silent-fallback.")
    print("ЕСКАЛАЦІЯ ФІНАНСИСТУ — альтернативні ключі (вибір за фінансистом):")
    print("  (а) Регистратор + Контрагент;")
    print("  (б) АналитикаУчетаПоПартнерам / СтатьиРасходов")
    print("      (як штатна Обработка.ДвиженияАктивовПассивы — розвідка підтвердила,")
    print("       що ПАП НЕ заповнює ИдентификаторФинЗаписи, звʼязок там через ці аналітики).")
    print("Після рішення фінансиста — повторити пре-тест для нового ключа, тоді план 2-4.")
    print()
    print("ВАЖЛИВО: Етап 1 (v1.3) — самодостатній робочий deliverable (OD-8),")
    print("закритий і НЕ залежить від результату цього GATE.")
print("[DECISION RECORDED]")
