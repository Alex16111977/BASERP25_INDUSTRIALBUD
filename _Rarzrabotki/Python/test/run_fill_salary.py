# -*- coding: utf-8 -*-
import win32com.client, sys, traceback, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

try:
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
    S = conn.String
    print("Connected OK")

    # Find document
    q = conn.NewObject("Запрос")
    q.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка, "
        "Док.Номер КАК Номер, "
        "Док.ПериодРегистрации КАК Период "
        "ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док "
        "ГДЕ НЕ Док.ПометкаУдаления "
        "УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
    )
    res = q.Execute()
    sel = res.Choose()

    if not sel.Next():
        print("ERROR: Document not found for Dec 2025!")
        sys.exit(1)

    print("Document:", S(sel.Номер), "Period:", S(sel.Период))
    ref = sel.Ссылка
    obj = ref.GetObject()

    kaz = obj.А_РаспределениеКазна.Count()
    nal = obj.А_РаспределениеНалоги.Count()
    print(f"A_Kazna: {kaz}, A_Nalogi: {nal}")

    bfiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
    bnach = obj.НачисленнаяЗарплатаИВзносы.Count()
    bndfl = obj.НачисленныйНДФЛ.Count()
    print(f"BEFORE: NachFiz={bfiz}, Nach={bnach}, NDFL={bndfl}")

    # Call the procedure
    print("Calling procedure...")
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print("CALL: OK (no errors)")

    afiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
    anach = obj.НачисленнаяЗарплатаИВзносы.Count()
    andfl = obj.НачисленныйНДФЛ.Count()
    print(f"AFTER: NachFiz={afiz}, Nach={anach}, NDFL={andfl}")

    if afiz > 0:
        print("--- NachZPPoFiz (first 5) ---")
        for i in range(min(5, afiz)):
            r = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i)
            print(f"  [{i}] FL={S(r.ФизическоеЛицо)}, Dept={S(r.ПодразделениеПредприятия)}, Sum={r.Сумма}, Contr={r.ВзносыВсего}")

    if andfl > 0:
        print("--- NachNDFL (first 5) ---")
        for i in range(min(5, andfl)):
            r = obj.НачисленныйНДФЛ.Get(i)
            print(f"  [{i}] FL={S(r.ФизическоеЛицо)}, Dept={S(r.ПодразделениеПредприятия)}, Sum={r.Сумма}, Op={S(r.ВидОперации)}")

    if anach > 0:
        print("--- NachZPIVznosy (first 5) ---")
        for i in range(min(5, anach)):
            r = obj.НачисленнаяЗарплатаИВзносы.Get(i)
            print(f"  [{i}] Dept={S(r.ПодразделениеПредприятия)}, Sum={r.Сумма}, Contr={r.ВзносыВсего}")

    print("\nTEST PASSED")

except Exception as e:
    print(f"ERROR: {e}")
    traceback.print_exc()
    sys.exit(1)
