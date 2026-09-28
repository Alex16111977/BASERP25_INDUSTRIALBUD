# -*- coding: utf-8 -*-
"""Приймання: «Створити документи» у ТЧ «Документи» РасчетКомплектаций (BuhBud, 28.09.2026).

Промт: _Rarzrabotki/notebook/knowledge_Бухгалтерия_РасчетКомплектаций/PROMPT_dorabotka_SozdanieDokumentovSpisaniya.md
Scratch-РК (маркер у коментарі) — копія шапки й рядків складів РК №10 без документів. Бойові РК не змінюються.
Документи попереднього прогону помічаються на видалення (не Удалить()) і відвʼязуються від scratch-РК.
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MARKER = "SMOKE_DOC_KOMPL_SOZD_v1"
v8 = win32com.client.Dispatch("V83.COMConnector")
b = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = b.String
FAILS = []


def check(cond, msg):
    print(("PASS " if cond else "FAIL ") + msg)
    if not cond:
        FAILS.append(msg)


def uid(ref):
    return S(ref.УникальныйИдентификатор())


def find_rk(where, **params):
    q = b.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ Р.Ссылка КАК Ссылка ИЗ Документ.РасчетКомплектаций КАК Р ГДЕ НЕ Р.ПометкаУдаления И " + where
    for k, v in params.items():
        q.SetParameter(k, v)
    r = q.Execute().Выгрузить()
    return r.Получить(0).Ссылка if r.Количество() else None


TABLES = ("ДокументиКомплектації", "ДокументиМалоценки")
COLS = ("ДокументПоНормам", "ДокументДодаткова")

# --- scratch-РК: get-or-create, шапка й склади з РК №10 ---
src = find_rk("Р.Номер = &Н", Н="00000000010").ПолучитьОбъект()
ref = find_rk("ВЫРАЗИТЬ(Р.Комментарий КАК Строка(100)) = &М", М=MARKER)
doc = ref.ПолучитьОбъект() if ref is not None else b.Документы.РасчетКомплектаций.СоздатьДокумент()
for f in ("Дата", "Подразделение", "Организация", "ПодразделениеОрганизации", "НоменклатураКомплект",
          "ДатаОтгрузки", "ДатаСписанияПоМонтажу", "Спецификация", "Направление"):
    setattr(doc, f, getattr(src, f))
doc.Комментарий = MARKER
for tch in TABLES:
    # попередні scratch-документи — на видалення
    t = getattr(doc, tch)
    for i in range(t.Количество()):
        for col in COLS:
            d = getattr(t.Получить(i), col)
            if b.ЗначениеЗаполнено(d):
                o = d.ПолучитьОбъект()
                if not o.ПометкаУдаления:
                    o.УстановитьПометкуУдаления(True)
    t.Очистить()
    st = getattr(src, tch)
    for i in range(st.Количество()):
        n = t.Добавить()
        s = st.Получить(i)
        n.Склад = s.Склад
        n.Организация = s.Организация
        if tch == "ДокументиМалоценки":
            n.ФизЛицо = s.ФизЛицо
            n.СпособОтраженияРасходов = s.СпособОтраженияРасходов
doc.Записать()
print("scratch-РК:", S(doc.Ссылка))

# --- 1. створення по всіх рядках обох ТЧ ---
created = {}
for tch in TABLES:
    t = getattr(doc, tch)
    nums = b.NewObject("Массив")
    for i in range(t.Количество()):
        nums.Добавить(i + 1)
    res = doc.СоздатьДокументыСписания(tch, nums)
    msgs = [res.Сообщения.Получить(i) for i in range(res.Сообщения.Количество())]
    for m in msgs:
        print("   ", m)
    created[tch] = int(res.Создано)
    check(int(res.Создано) == 2 * t.Количество() and int(res.Пропущено) == 0,
          f"{tch}: створено {res.Создано} з {2 * t.Количество()}, пропущено {res.Пропущено}")
doc.Записать()   # запис РК проходить контроль складу
check(True, "РК записано після створення (контроль «той ли склад» пройдено)")
check(doc.РасхожденияСкладовДокументов().Количество() == 0, "розбіжностей складу немає")

# --- 2. шапки ---
nds = uid(b.Справочники.НалоговыеНазначенияАктивовИЗатрат.НДС_Облагаемая)
montazh = S(b.Перечисления.А_ВидСклада.Монтаж)   # у перечислень немає GUID
for tch in TABLES:
    mal = tch == "ДокументиМалоценки"
    t = getattr(doc, tch)
    for i in range(t.Количество()):
        row = t.Получить(i)
        skl = row.Склад
        is_m = S(skl.А_ВидСклада) == montazh
        base = doc.ДатаСписанияПоМонтажу if is_m else doc.ДатаОтгрузки
        hh = (19 if mal else 20) if is_m else (8 if mal else 9)
        mm = 0 if is_m or not mal else 30
        for col in COLS:
            d = getattr(row, col)
            po = col == "ДокументПоНормам"
            tag = f"{tch}[{i + 1}].{col} {S(d)}"
            ok = (uid(d.Склад) == uid(skl) and uid(d.Организация) == uid(skl.А_Организация)
                  and d.ПоНорме == po and not d.Проведен
                  and ("в нормі" if po else "понад норму") in S(d.Комментарий)
                  and uid(d.Подразделение) == uid(skl.Подразделение)
                  and S(d.Направление) == "Производство")
            exp_time = f"{hh}:{mm:02d}:{0 if po else 1:02d}"
            ok_date = S(d.Дата)[:10] == S(base)[:10] and S(d.Дата).endswith(exp_time)
            if mal:
                ok_tch = d.МалоценныеАктивы.Количество() == 0
                owner_ok = uid(d.ПодразделениеОрганизации.Владелец) == uid(d.Организация)
                check(ok and ok_date and ok_tch and owner_ok,
                      f"{tag}: шапка ПМА (дата {S(d.Дата)}, підрозділ орг. «{S(d.ПодразделениеОрганизации)}»)")
                if S(d.Организация).find("ІНДЕПТ") >= 0:
                    check("Цех металоконструкцій" in S(d.ПодразделениеОрганизации),
                          f"{tag}: ІС → «Виробничий (Цех металоконструкцій)»")
            else:
                ok_k = (S(d.ВидОперации) == "Комплектація" and float(d.Количество) == 0.5
                        and S(d.СчетУчетаБУ.Код) == "26" and uid(d.НалоговоеНазначение) == nds
                        and uid(d.Номенклатура) == uid(doc.НоменклатураКомплект)
                        and d.Комплектующие.Количество() == 0)
                check(ok and ok_date and ok_k, f"{tag}: шапка комплектації (дата {S(d.Дата)}, «{S(d.Комментарий)}»)")

# --- 3. повторний запуск нічого не створює ---
for tch in TABLES:
    t = getattr(doc, tch)
    nums = b.NewObject("Массив")
    for i in range(t.Количество()):
        nums.Добавить(i + 1)
    res = doc.СоздатьДокументыСписания(tch, nums)
    check(int(res.Создано) == 0, f"{tch}: повторно створено 0 (є {res.Создано})")

# --- 4. порожня дата списання по монтажу → рядок «Відвантажено» пропущено (лише в памʼяті) ---
obj = doc.Ссылка.ПолучитьОбъект()
t = obj.ДокументиКомплектації
idx = None
for i in range(t.Количество()):
    sk = t.Получить(i).Склад
    if S(sk.А_ВидСклада) == montazh:
        idx = i
        break
check(idx is not None, "у scratch-РК є склад «Відвантажено»")
if idx is not None:
    row = t.Получить(idx)
    row.ДокументПоНормам = b.Документы.КомплектацияНоменклатуры.ПустаяСсылка()
    row.ДокументДодаткова = b.Документы.КомплектацияНоменклатуры.ПустаяСсылка()
    obj.ДатаСписанияПоМонтажу = b.NewObject("ОписаниеТипов", "Дата").ПривестиЗначение(None)
    nums = b.NewObject("Массив")
    nums.Добавить(idx + 1)
    res = obj.СоздатьДокументыСписания("ДокументиКомплектації", nums)
    msg = res.Сообщения.Получить(0) if res.Сообщения.Количество() else ""
    check(int(res.Создано) == 0 and int(res.Пропущено) == 1 and "Дата списання по монтажу" in msg,
          f"порожня дата монтажу → пропуск із причиною ({msg})")

print()
print("ИТОГ:", "FAIL " + str(len(FAILS)) if FAILS else "ВСЕ PASS")
sys.exit(1 if FAILS else 0)
