# -*- coding: utf-8 -*-
"""Z3, часть 1: эталон 15 м по M5 §3.1 + выбор бухгалтера/прораба, дерево узлов строителя, цены норм §4.5,
варианты домов §3.7 (копия узла с цепочкой родителей), разузлование §4.2–4.4. В памяти, ничего не пишет в 1С."""
import hashlib, uuid, copy, math
from base15 import *

SS_ALL = RAW["ss"]
SS = [l for l in SS_ALL if Dd(l["q"]) != 0]
ZERO = [l for l in SS_ALL if Dd(l["q"]) == 0]
LINE = {l["n"]: l for l in SS_ALL}
for l in SS_ALL:
    l["q"] = Dd(l["q"]); l["sum"] = Dd(l["sum"])
    l["on"] = (l["on"] or "").strip(); l["pos"] = (l["pos"] or "").strip()

ONC = collections.defaultdict(list)                 # ОН (код) -> активные карточки
for c in RAW["on_cards"]:
    if not c["Пом"]:
        ONC[c["ОНКод"]].append(c["Код"])

FQ = collections.defaultdict(lambda: collections.defaultdict(lambda: Z))   # списано: код -> дом -> кол
FS = collections.defaultdict(lambda: collections.defaultdict(lambda: Z))   # списано: код -> дом -> сумма без ПДВ
for f in FACT:
    FQ[f["code"]][f["house"]] += f["q"]; FS[f["code"]][f["house"]] += f["s"]
LQ = collections.defaultdict(lambda: collections.defaultdict(lambda: Z))   # лежит на складах дома
LS = collections.defaultdict(lambda: collections.defaultdict(lambda: Z))
for c in list(st.keys()):
    for h in H15:
        q, s = bal(c, "H:" + h)
        if q or s:
            LQ[c][h] = q; LS[c][h] = s
USED15 = {c for c in FQ if any(FQ[c][h] for h in H15)}

UKEY_CARD = {c: CARD[c]["ukey"] for c in CARD}


def ukey_eq(a, b):
    a, b = (a or "").strip(), (b or "").strip()
    norm = lambda x: "м" if x == "пог.м" else x
    return norm(a) == norm(b)


# ───────────────────────── цены закупки (без ПДВ) с защитой от ошибки единицы (Z1 §1.4)
ANY = collections.defaultdict(list)
for r in RAW["receipts_any"]:
    q = Dd(r["Кол"])
    if q <= 0:
        continue
    net = Dd(r["Сумма"]) - (Dd(r["НДС"]) if r["ВклНДС"] else Z)
    ANY[r["Код"]].append({"date": pdate(r["Дата"]), "price_n": net / q, "doc": r["Номер"], "podr": r["Подр"]})


def fact_cost(c):
    q = sum(FQ[c][h] for h in H15); s = sum(FS[c][h] for h in H15)
    if q > 0:
        return s / q
    a = st[c]["P"]
    return a[1] / a[0] if a[0] > 0 else None


def price_last(c):
    """(цена без ПДВ, основа). Основа: ПТУ проекта / ПТУ любая / себестоимость факта; ошибка единицы — цена ≠ себестоимости ×5."""
    fc = fact_cost(c)
    for src, lst in (("ПТУ проєкту", sorted(PURCH.get(c, []), key=lambda x: x["date"], reverse=True)),
                     ("ПТУ інша", sorted(ANY.get(c, []), key=lambda x: x["date"], reverse=True))):
        for i, rec in enumerate(lst):
            p = rec["price_n"]
            if fc and fc > 0 and (p / fc > 5 or p / fc < Decimal("0.2")):
                continue                                  # ошибка единицы — берём предыдущую правдоподобную
            return p, src + ("" if i == 0 else " (попередня: остання з помилкою одиниці)"), rec["date"]
    if fc:
        return fc, "собівартість факту", None
    return None, "немає", None


# ───────────────────────── §3.1 «Заповнити з СС»: кандидаты ТОЛЬКО из базы
GROUP_ON = {"гвинтові кріплення"}
SEPTIC_N = 168
AUTO = {}
for l in SS:
    cards = ONC.get(l["on_code"], [])
    used = [c for c in cards if c in USED15]
    if K(l["on"]) in GROUP_ON:
        AUTO[l["n"]] = {"cat": "Група", "card": None}
    elif l["n"] == SEPTIC_N:
        AUTO[l["n"]] = {"cat": "Вузол (позиція-комплект)", "card": None, "used": used}
    elif len(cards) == 1:
        AUTO[l["n"]] = {"cat": "зелена: єдина картка ОН", "card": cards[0]}
    elif len(used) == 1:
        AUTO[l["n"]] = {"cat": "жовта: єдина списана на будинки типу", "card": used[0]}
    elif len(cards) == 0:
        AUTO[l["n"]] = {"cat": "вибрати: у ОН немає карток", "card": None}
    elif len(used) == 0:
        AUTO[l["n"]] = {"cat": "вибрати: жодна не списувалась", "card": None, "cards": len(cards)}
    else:
        AUTO[l["n"]] = {"cat": "вибрати: списувались 2+", "card": None, "used": used}

# ───────────────────────── выбор бухгалтера (симуляция; прораб подсказывает по назначению позиции)
# n -> список строк: (код | None, доля количества позиции, k «1 од. картки = k од. рядка», источник k, примечание)
M = lambda code, note="", k=1, ksrc=None, share=1: (code, Decimal(str(share)), Decimal(str(k)), ksrc, note)
MANUAL = {
    1: [M("00000007615", "розмір 100*100*3")], 2: [M("00-00021868", "розмір 50*50*2")], 3: [M("00000010533", "дв 40*3,5 = труба 40*3,5")],
    4: [M(None, "картки «89*3,5» у базі немає — завести або виключити")],
    5: [M("00-00020143", "лист 4")], 6: [M("00-00020550", "лист 8")], 7: [M("00-00008024", "лист 14")],
    8: [M("00000011364", "М48")], 10: [M("00000010963", "М24")], 11: [M("00000002877", "М24×200")],
    12: [M("00000016107", "40*40*3")], 15: [M("00000013053", "50*50*4")],
    16: [M("00000008418", "М12×60")], 17: [M("00-00010035", "М12×90")],
    18: [M("00000005990", "гайка М12 ISO 4032 (на №1–№5); на №6 — DIN 934")], 19: [M("00000005781", "шайба М12 DIN125")],
    20: [M("00-00010036", "гровер = шайба пружинна DIN127")],
    27: [M("00000000624", "CORGON18 (МІКС-1) балон 40 л", 1, "Вручну: «балон» = «шт» (1 балон = 1 шт)")],
    28: [M("00000006053", "диск 230 мм — картка в ОН «мшп»; жовта пропозиція «Круг 350» відхилена (інший розмір)")],
    41: [M("00-00000859", "вітробар'єр стін = мембрана дифузійна; Strotex — поза нормою (питання 3 M5: «дозволені заміни»)")],
    42: [M("00-00007736", "пароізоляція — всередині стін/стелі (прораб: половина позиції)", share="0.5"),
         M("00000021537", "гідроізоляція підпокрівельна — у покрівлі (прораб: половина позиції)", share="0.5")],
    52: [M("00-00009980", "меблевий кутик = «Плінтус на стільницю 3м» (ОН нульового рядка 173)", 3, "Текст картки: «3м» → 1 шт = 3 м")],
    57: [M("00000000596", "клей монтажний білий = Lacrysil (ОН «Піна клей»)")],
    59: [M("00-00021971", "кутник 30×30")], 60: [M("00-00021972", "кутник 50×50")],
    61: [M("00000014244", "клей-піна 860 мл")], 78: [M("БП-00019002", "піна монтажна")],
    62: [M("00-00016723", "ПК 8 = S-8")], 63: [M("00-00016722", "ПК 20 покрівля ≈ Т-18")], 64: [M("00000011151", "ПК 10 стіни = ПС-10")],
    66: [M("00-00021685", "ринва Profil 90")], 67: [M("00-00021686", "труба водостічна 75")],
    70: [M("00-00021693", "«Відвід 45» на фасаді = коліно водостоку; жовта «Коліно каналіз. 110*87» відхилена")],
    71: [M("00-00021697", "держак ринви")], 72: [M("00-00021696", "кріплення труби водостоку = держак труби; жовта «Комплект кріплення 47–52» відхилена")],
    73: [M("00-00021695", "з'єднувач ринви")],
    79: [M("00-00022083", "корпус ЩРН-П-36 (на №5–№6 — бокс 36 авт.)")],
    80: [M("00000018518", "3х4")], 81: [M("00000018382", "3х2,5")], 82: [M("00000018520", "3х1,5")],
    83: [M("00-00017553", "гофра 20 мм (не списувалась)")],
    84: [M("00-00003557", "10 А")], 85: [M("00-00003558", "16 А")], 86: [M("00-00007600", "у СС «автомат 25А», фактично на всіх 6 будинках ставили ввідний 2р 40 А; картка 25 А не списувалась")],
    88: [M("00-00005046", "100×100×45")], 92: [M("00-00005050", "коробка монтажна")],
    89: [M("00-00005047", "WGn-2")], 90: [M("00-00005048", "WGn-3")], 91: [M("00-00005049", "WGn-5")],
    93: [M("00000000615", "32 А IP44")], 104: [M("00-00022088", "із заземленням")], 105: [M("00000019254", "IP44 з кришкою")],
    106: [M("00-00005055", "подвійна")],
    95: [M("00-00022085", "LED 12W")], 96: [M("00-00005051", "панель 36W")], 97: [M("00-00005052", "рамка")],
    98: [M("00-00005053", "аварійний")], 99: [M("00-00005054", "стінний")], 100: [M("00000007125", "лампа 7 Вт проти 15 Вт у СС")],
    101: [M("00000022393", "2-клавішний")], 103: [M("00-00000721", "1-клавішний")],
    107: [M("00-00015564", "кабель Hemstedt — на 5 будинках з 6")],
    110: [M("00-00013627", "гратка ДВ 102 ВК")],
    114: [M(None, "труби д32 у базі не списувались; жовта «110/3000» відхилена")],
    115: [M(None, "труби д32 у базі не списувались; жовта «110/3000» відхилена")],
    116: [M("00-00013067", "відрізок 0,5 м з труби 3 м", 3, "Текст картки: «/3000» → 1 шт = 3 м")],
    117: [M("00-00013067", "труба 3 м", 3, "Текст картки: «/3000» → 1 шт = 3 м")],
    118: [M("00000010982", "відвід д110 45° (ОН «Коліно обвідне»); жовта «110*87» відхилена")],
    119: [M(None, "відводу д32 у базі немає; жовта «110*87» відхилена")],
    120: [M("00-00005134", "назва збігається (не списувалась)")],
    121: [M("00-00000555", "ексцентричний 110/50")], 122: [M("00000021706", "трійник 110-50 87°")],
    123: [M("00-00015629", "трійник д32 (не списувався)")],
    126: [M("00-00005955", "сифон раковини")], 127: [M("00-00005957", "сифон піддону")],
    128: [M("00-00015630", "хомут 110–118 мм (каналізація д110); жовта «47–52» відхилена")],
    129: [M("00-00002506", "комплект 47–52 мм замість 31–37 у СС — перевірити розмір")],
    130: [M("00-00017797", "PP-R 20")],
    131: [M("00000000106", "відвід ППР д20 = куточок 90 PP-R D20 (ОН нульового рядка 169); жовта «110*87» відхилена")],
    132: [M("00-00019481", "обвід PP-R D20")], 133: [M("00-00009835", "трійник ПП-Р 20")], 134: [M("00000000101", "трійник 20×1/2 Н")],
    135: [M("00-00007659", "перехід ППР 20×1/2 зов = муфта D20×1/2 зовн. (ОН «Муфта»)")],
    137: [M("00000000341", "гнучкий шланг 1/2 ВВ 30 см")],
    138: [M("00-00005442", "перехід латунний = футорка латунна")],
    140: [M("00-00006152", "з'єднувач компресійний 25 = муфта ПЕ 25 затискна (ОН «Муфта»)")],
    141: [M("00-00022163", "труба ПЕ ф25×2,0 (на №3 — ф25×2,3)")],
    142: [M("00-00015635", "хомут PP-R д20 = кліпса ППР 20 (ОН «кліпса»)")],
    155: [M("00-00005561", "дозатор мила = диспенсер рідкого мила (інша ОН)")],
    156: [M("00-00005366", "дозатор рушників = диспенсер рушників (інша ОН)")],
}
# единица карточки ≠ единице строки, k задаёт бухгалтер фразой «1 [од. картки] = k [од. рядка]» (источник — текст карточки или факт)
KFIX = {44: (Decimal(1), "Вручну: «пач» = «пак» — синонім, якого немає в КлючЕдиницы; факт 9,016 = норма"),
        45: (Decimal(1), "Вручну: «пач» = «пак» — синонім, якого немає в КлючЕдиницы; факт 14,016 = норма")}
# предложения V2 «заміна аналогом», которые прораб отклонил (другой размер / другое назначение)
REJECT_V2 = {("00-00015629", "00000000100"): "трійник д32 ≠ трійник PP-R д20 (інший розмір)",
             ("00-00015630", "00-00015642"): "хомут 110–118 мм (каналізація д110) ≠ хомут 47–51 мм (інший діаметр)"}
# жёлтые, которые бухгалтер отклонил (правда отклонена — в MANUAL), считаем отдельно
for n, a in AUTO.items():
    a["manual"] = n in MANUAL
    a["yellow_rejected"] = a["cat"].startswith("жовта") and n in MANUAL and (MANUAL[n][0][0] != a["card"])

# ───────────────────────── узлы строителя: дерево и строки СС
# id -> (наименование, родитель, кол-во в родителе, ед. выпуска, кол-во выпуска)
NODES = {
    "N01": ("Каркас модуля", None, 1, "компл.", 1),
    "N01a": ("Болтові з'єднання і анкери каркаса", "N01", 1, "компл.", 1),
    "N01b": ("Зварювання і фарбування каркаса", "N01", 1, "компл.", 1),
    "N02": ("Підлога (на 1 м²)", None, Decimal("18"), "м²", 1),
    "N03": ("Стіни і перегородки", None, 1, "компл.", 1),
    "N03a": ("Внутрішня обшивка стін (на 1 м²)", "N03", Decimal("80.1"), "м²", 1),
    "N04": ("Стеля", None, 1, "компл.", 1),
    "N05": ("Покрівля", None, 1, "компл.", 1),
    "N05a": ("Водостік 15 м", "N05", 1, "компл.", 1),
    "N06": ("Фасад", None, 1, "компл.", 1),
    "N07": ("Вікна й двері", None, 1, "компл.", 1),
    "N08": ("Електрика", None, 1, "компл.", 1),
    "N08a": ("Електрощит", "N08", 1, "шт", 1),
    "N08b": ("Резервне живлення (генератор)", "N08", 1, "компл.", 1),
    "N09": ("Опалення, вентиляція і клімат", None, 1, "компл.", 1),
    "N10": ("Водопостачання і сантехніка", None, 1, "компл.", 1),
    "N11": ("Каналізація всередині", None, 1, "компл.", 1),
    "N12": ("Зовнішні мережі", None, 1, "компл.", 1),
    "N12a": ("Септик", "N12", 1, "компл.", 1),
    "N13": ("Оснащення", None, 1, "компл.", 1),
    "N14": ("Кріплення загального призначення", None, 1, "компл.", 1),
}
ROWS_OF_NODE = {
    "N01": [1, 2, 3, 4, 12, 15, 5, 6, 7, 13, 14, 21],
    "N01a": [8, 9, 10, 11, 16, 17, 18, 19, 20],
    "N01b": [22, 23, 25, 26, 27, 28],
    "N02": [31, 24, 46, 47, 50],
    "N03": [29, 30, 43, 45, 41, "42a", 49, 55, 59, 60, 52],
    "N03a": [54, 48],
    "N04": [44, 51],
    "N05": [63, "42b"],
    "N05a": [66, 67, 68, 69, 70, 71, 72, 73],
    "N06": [64, 62, 65],
    "N07": [74, 77, 58, 75, 76, 32, 78, 56, 57, 61],
    "N08": [80, 81, 82, 83, 88, 92, 89, 90, 91, 104, 105, 106, 101, 103, 102, 95, 96, 97, 98, 99, 100, 111],
    "N08a": [79, 84, 85, 86, 87],
    "N08b": [165, 112, 93, 94],
    "N09": [113, 107, 152, 108, 109, 110],
    "N10": [130, 131, 132, 133, 134, 135, 142, 136, 139, 137, 138, 145, 148, 146, 147],
    "N11": [114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 53],
    "N12": [167, 144, 141, 140, 143],
    "N12a": [],                     # строка 168 «Комплект септика» стоит на строке-узле, детали наследуют позицию
    "N13": [149, 150, 151, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 166],
    "N14": [33, 34, 35, 36, 37, 38, 39, 40],
}
# спорные решения прораба: строка -> (узел, почему)
DISPUTED = [
    (14, "N01", "ПВЛ 406 (просічно-витяжний лист) — настил сходинок/майданчика входу, приварюється до каркаса; окремого вузла «ганок» у СС немає"),
    (21, "N01", "«Завіса 16мм» — картка «Петля виточена d=20 мм (для воріт)», приварна петля (етап СС «Каркас»): ставлять на каркас (двері техніші), а не з дверними блоками"),
    (31, "N02", "Дошка 20×100 (0,08 м³ ≈ 40 м) — лаги під ЦСП підлоги 18 м²; могла бути обрешіткою покрівлі — у СС етап «Каркас», не розрізнити"),
    (24, "N02", "Пінополістирол 100 мм 40 шт = 20 м² — утеплення підлоги (18 м² + запас), хоча етап СС «Каркас»"),
    (29, "N03", "Брус 50×50 0,863 м³ (≈ 345 м) — каркас під обшивку стін і перегородок; частина могла піти на стелю"),
    (43, "N03", "Пінопласт 50 мм 88 шт = 44 м² ≈ вітробар'єр 42 м² — зовнішнє утеплення стін"),
    (45, "N03", "Мінвата 50 мм 14 пач ≈ 171 м² — стіни й перегородки (двошарово); 100 мм — у стелю (N04)"),
    (44, "N04", "Мінвата 100 мм 9 пач ≈ 55 м² — утеплення стелі/покрівлі (стеля 18 м² × 2 шари + запас); прораб кладе в «Стелю»"),
    ("42a", "N03", "«Паро-гідро бар'єр» 1,6 рул — одна позиція СС, два матеріали: пароізоляція (стіни/стеля) і гідроізоляція (під профнастил покрівлі). Поділ 0,8/0,8 — рішення прораба"),
    ("42b", "N05", "друга половина позиції 42 — гідроізоляційна плівка підпокрівельна"),
    (49, "N03", "Клей для лінолеуму — і на стіни (80,1 м²), і на підлогу (18 м²); поставлено в «Стіни», бо 82 % площі"),
    (55, "N03", "Грунт під клей — разом із клеєм лінолеуму"),
    (52, "N03", "«Меблевий кутик» 84 м — за цехом це пластиковий кутик оздоблення стін («Плінтус на стільницю 3м»), не меблі"),
    (53, "N11", "Дверцята ревізійні — доступ до сифонів/каналізації, ставлять разом із каналізацією"),
    (65, "N06", "Гладкий лист білий 23 м² — відливи, відкоси і обрамлення фасаду"),
    (62, "N06", "ПК 8 (S-8) 16,5 м² — фасадні/цокольні вставки; не покрівля"),
    (58, "N07", "Теплий відкіс (сендвіч-панель 10 мм) — відкоси вікон"),
    (32, "N07", "Замок навісний — на вхідні двері / ворота техніші"),
    (57, "N07", "Клей монтажний білий — монтаж підвіконь і відкосів"),
    (61, "N07", "Піна-клей — монтаж вікон/дверей і панелей; частина могла піти на стіни"),
    (165, "N08b", "Генератор 7 кВт + АВР + розетка 32 А IP44 + вилка 32 А = вузол резервного живлення (підключення генератора)"),
    (93, "N08b", "розетка стаціонарна 32 А IP44 — ввід від генератора, а не звичайна розетка"),
    (111, "N08", "Ізоляційна стрічка — витратний матеріал електромонтажу"),
    (107, "N09", "Тепла підлога (кабель) — опалення; покупний комплект, не вузол"),
    (142, "N10", "Хомут PP-R д20 → кліпса ППР 20: кріплення водопроводу"),
    (141, "N12", "Труба ПЕ д25 50 м — зовнішній водопровід від ємності; з'єднувачі компресійні 25 — там само"),
    (140, "N12", "з'єднувачі компресійні 25 мм — на трубі ПЕ д25 зовнішнього водопроводу"),
    (128, "N11", "«Комплект кріплень 104-112 мм» — хомути каналізаційної труби д110"),
    (129, "N11", "«Комплект кріплення труб 31-37мм» — кріплення каналізації д32; фактично 47–52 мм (на трубу 50)"),
    (139, "N10", "Зворотний клапан — на вводі води (насосна станція), але тут — у сантехніці будинку"),
    (154, "N13", "Сигналізація диму — оснащення (автономні сповіщувачі), не електрика"),
    ("33-40", "N14", "«Гвинтові кріплення» — одна група на будинок: за фактом 12 видів саморізів, місце встановлення не відновити (навіть покрівельні 4,8×50); тому окремий вузол «Кріплення», а не всередині покрівлі/стін"),
]
NODE_OF_ROW = {}
for nid, rows in ROWS_OF_NODE.items():
    for r in rows:
        NODE_OF_ROW[r] = nid
NODE_OF_ROW[SEPTIC_N] = "N12a"

# ───────────────────────── спецификации в памяти
SPECS = {}          # id -> {"name","kind","out_q","out_u","lines":[...], "base": id|None, "houses":[]}


def kkey(*parts):
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "z3/" + "/".join(str(p) for p in parts)))


def new_line(kind, n=None, card=None, q=Z, unit="", k=Decimal(1), ksrc="", node=None, src="", note="", sub=""):
    l = LINE.get(n) if isinstance(n, int) else None
    return {"key": kkey(kind, n, sub, card, node), "kind": kind, "n": n, "pos": l["pos"] if l else "", "pos_key": l["key"] if l else "",
            "etap": l["etap"] if l else "", "card": card, "on": (CARD.get(card, {}).get("on") if card else (l["on"] if l else "")),
            "q": q, "unit": unit, "k": k, "ksrc": ksrc, "node": node, "src": src, "checked": False, "price": None, "note": note}


def build_etalon():
    for nid, (name, parent, q_in_parent, unit_out, out_q) in NODES.items():
        SPECS[nid] = {"id": nid, "name": name, "kind": "Вузол", "out_q": Decimal(out_q), "out_u": unit_out, "lines": [], "base": None,
                      "parent_hint": parent}
    SPECS["E"] = {"id": "E", "name": "МД IRС 2026 15 м — еталон", "kind": "Будинок", "out_q": Decimal(1), "out_u": "будинок",
                  "lines": [], "base": None}
    area = {nid: (NODES[nid][2] if NODES[nid][3] == "м²" else None) for nid in NODES}
    # строки-узлы
    for nid, (name, parent, q_in_parent, unit_out, out_q) in NODES.items():
        host = SPECS[parent] if parent else SPECS["E"]
        ln = new_line("Вузол", n=SEPTIC_N if nid == "N12a" else None, q=Dd(q_in_parent), unit=unit_out, node=nid,
                      src="Вузол прораба" if nid != "N12a" else "Позиція-комплект СС → вузол", sub=nid)
        ln["checked"] = True
        host["lines"].append(ln)
    # строки-листья из позиций СС
    for nid, rows in ROWS_OF_NODE.items():
        for r in rows:
            sub = ""
            if isinstance(r, str) and r.startswith("42"):
                n, sub = 42, r
            else:
                n = r
            l = LINE[n]
            a = AUTO[n]
            per = area[nid]                               # площадной узел: количество на 1 м²
            if a["cat"] == "Група":
                q = l["q"]
                ln = new_line("Група за загальною назвою", n=n, q=q, unit=l["unit"].strip(), src="Група (список «групою»)", sub=sub)
                ln["on"] = l["on"]; ln["checked"] = True
                SPECS[nid]["lines"].append(ln); continue
            choices = MANUAL.get(n) or [(a["card"], Decimal(1), Decimal(1), None, "")]
            if sub == "42a":
                choices = [MANUAL[42][0]]
            elif sub == "42b":
                choices = [MANUAL[42][1]]
            for i, (code, share, k, ksrc, note) in enumerate(choices):
                q = l["q"] * share
                if per:
                    q = (q / Dd(per)).quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP)      # Число(15,6)
                if code and not ksrc and n in KFIX:
                    k, ksrc = KFIX[n]
                if code and not ksrc:
                    if ukey_eq(l["ukey"], UKEY_CARD.get(code, "")):
                        k, ksrc = Decimal(1), "ЄдиниціЗбігаються"
                    else:
                        k, ksrc = Decimal(1), "НеЗнайдено"
                if n in MANUAL:
                    src = ("Вручну: жовту відхилено" if a.get("yellow_rejected") else "Вручну (симуляція бухгалтера)")
                elif a["cat"].startswith("зелена"):
                    src = "ЄдинаКарткаОН"
                elif a["cat"].startswith("жовта"):
                    src = "СписуваласьНаБудинкиТипу"
                else:
                    src = "без картки"
                ln = new_line("Деталь", n=n, card=code, q=q, unit=l["unit"].strip(), k=k, ksrc=ksrc or "", src=src, note=note,
                              sub=sub or str(i))
                ln["checked"] = bool(code)
                SPECS[nid]["lines"].append(ln)
    # септик: «Склад з документа комплектації» — ТЧ «Комплектующие» КН ІБ00-000174 (30 м №1), без деления на выпуск 0,5 (M5 §3.2)
    comp = collections.OrderedDict()
    for r in RAW["kn_septic"]:
        if r["Номер"] == "ІБ00-000174":
            comp[r["Код"]] = comp.get(r["Код"], Z) + Dd(r["Кол"])
    for code, q in comp.items():
        ln = new_line("Деталь", n=None, card=code, q=q, unit=CARD[code]["unit"], ksrc="ЄдиниціЗбігаються",
                      src="З документа комплектації ІБ00-000174", sub="sept")
        ln["checked"] = True
        SPECS["N12a"]["lines"].append(ln)
    return comp


SEPTIC_COMP = build_etalon()


# ───────────────────────── разузлование §4.2–4.4
def md5(s):
    return hashlib.md5(s.encode("utf-8")).hexdigest()


def explode(root):
    leaves, inst = [], []

    def walk(sid, mult, chain, kinst, etap, pos, pos_key, lvl, path):
        if lvl >= 10:
            raise RuntimeError("Глибина вкладеності > 10: " + " → ".join(path))
        S = SPECS[sid]
        share = mult / S["out_q"]
        for ln in S["lines"]:
            qq = ln["q"] * share
            ch = chain + ln["key"]
            key = md5(ch)
            e = ln["etap"] or etap; p = ln["pos"] or pos; pk = ln["pos_key"] or pos_key
            if ln["kind"] == "Вузол":
                if ln["node"] in path:
                    raise RuntimeError("Цикл: " + " → ".join(path + [ln["node"]]))
                inst.append({"key": key, "parent": kinst, "node": ln["node"], "name": SPECS[ln["node"]]["name"], "K": qq,
                             "unit": SPECS[ln["node"]]["out_u"], "level": lvl + 1, "path": "/".join(path[1:] + [ln["node"]]),
                             "pos": p, "pos_key": pk})
                walk(ln["node"], qq, ch + "/", key, e, p, pk, lvl + 1, path + [ln["node"]])
            else:
                qc = qq / (ln["k"] or Decimal(1))
                leaves.append({"key": key, "line_key": ln["key"], "inst": kinst, "spec": sid, "kind": ln["kind"], "card": ln["card"],
                               "on": ln["on"], "n": ln["n"], "etap": e, "pos": p, "pos_key": pk, "unit_line": ln["unit"], "q_line": qq,
                               "q": qc, "price": ln["price"], "level": lvl, "path": "/".join(path[1:]), "src": ln["src"], "line": ln})

    walk(root, Decimal(1), "", "", "", "", "", 0, [root])
    for lf in leaves:
        lf["sum_n"] = (lf["q"] * (lf["price"] or Z)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if lf["price"] is not None else None
    for it in inst:
        it["V"] = sum(((lf["q"] * (lf["price"] or Z)) for lf in leaves if lf["inst"] == it["key"] or _under(lf, it, inst)), Z)
    return leaves, inst


def _under(lf, it, inst):
    k = lf["inst"]
    par = {x["key"]: x["parent"] for x in inst}
    while k:
        if k == it["key"]:
            return True
        k = par.get(k)
    return False


# ───────────────────────── §4.5 «Розрахувати ціни норми» (в эталоне; веса — цена посл. закупки × кол-во на дату команды)
def calc_norm_prices():
    leaves, _ = explode("E")
    byp = collections.defaultdict(list)
    for lf in leaves:
        n = lf["n"] if lf["n"] else SEPTIC_N          # детали септика наследуют позицию 168
        byp[n].append(lf)
    protocol = []
    for n, lst in byp.items():
        P = LINE[n]["sum"]
        ws = []
        equal = False
        for lf in lst:
            if lf["kind"] != "Деталь" or not lf["card"]:
                equal = True; ws.append(None); continue
            p, basis, _ = price_last(lf["card"])
            if p is None:
                equal = True; ws.append(None); continue
            ws.append(lf["q"] * p * VAT)
        if equal:
            ws = [Decimal(1)] * len(lst)
            if len(lst) > 1:
                protocol.append(f"поз. {n}: ваги рівні (не в усіх листків є закупівля/картка)")
        tw = sum(ws)
        for lf, w in zip(lst, ws):
            qn = lf["q"] if lf["kind"] == "Деталь" and lf["card"] else lf["q_line"]
            price = (P * w / tw / qn).quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP) if qn else Z
            lf["line"]["price"] = price
            lf["line"]["price_basis"] = "за карткою" if lf["kind"] == "Деталь" and lf["card"] else "за од. рядка"
    return protocol


PRICE_PROTOCOL = calc_norm_prices()
ETALON_LEAVES, ETALON_INST = explode("E")
ETALON_SUM = sum(lf["sum_n"] for lf in ETALON_LEAVES)


# ───────────────────────── варианты домов §3.7: полная копия эталона + копия узла с цепочкой родителей
def find_paths(root, pred):
    """Все пути [(spec_id, idx), ...] до строк, удовлетворяющих pred."""
    res = []

    def walk(sid, path):
        for i, ln in enumerate(SPECS[sid]["lines"]):
            if pred(ln):
                res.append(path + [(sid, i)])
            if ln["kind"] == "Вузол":
                walk(ln["node"], path + [(sid, i)])

    walk(root, [])
    return res


def own_spec(v, path, house):
    """Copy-on-write: делает узлы по пути собственными для варианта v. Возвращает (spec_id, idx) последней строки."""
    sid, idx = path[0]
    assert sid == v
    for depth in range(1, len(path)):
        nsid, nidx = path[depth]
        parent_sid, parent_idx = path[depth - 1]
        if nsid.endswith("@" + house):
            continue
        new_id = nsid + "@" + house
        if new_id not in SPECS:
            base = SPECS[nsid]
            SPECS[new_id] = {"id": new_id, "name": f"{base['name']} ({HNAME[house]})", "kind": "Вузол", "out_q": base["out_q"],
                             "out_u": base["out_u"], "lines": copy.deepcopy(base["lines"]), "base": nsid,
                             "base_lines": copy.deepcopy(base["lines"]), "house": house}
            VARIANTS[house]["copied_nodes"].append({"node": nsid, "copy": new_id, "name": SPECS[new_id]["name"]})
        SPECS[parent_sid]["lines"][parent_idx]["node"] = new_id
        path[depth] = (new_id, nidx)
    return path[-1]


VARIANTS = {}


def make_variant(h):
    vid = "V" + h
    if vid in SPECS:
        return vid
    SPECS[vid] = {"id": vid, "name": f"МД IRС 2026 15 м — {HNAME[h]}", "kind": "Будинок", "out_q": Decimal(1), "out_u": "будинок",
                  "lines": copy.deepcopy(SPECS["E"]["lines"]), "base": "E", "base_lines": copy.deepcopy(SPECS["E"]["lines"]),
                  "houses": [h]}
    VARIANTS[h] = {"spec": vid, "edits": [], "copied_nodes": []}
    return vid


def card_name(c):
    return CARD.get(c, {}).get("name", c or "—") if c else "—"


def edit_set_card(h, old_card, new_card, why, rule, keep_qty=True):
    vid = make_variant(h)
    paths = find_paths(vid, lambda ln: ln["kind"] == "Деталь" and ln["card"] == old_card)
    assert paths, (h, old_card)
    sid, i = own_spec(vid, paths[0], h)
    ln = SPECS[sid]["lines"][i]
    ln["card"] = new_card; ln["on"] = CARD.get(new_card, {}).get("on", ln["on"]); ln["checked"] = False
    if not ukey_eq(UKEY_CARD.get(new_card, ""), UKEY_CARD.get(old_card, "")):
        ln["ksrc"] = "НеЗнайдено"
    ln["price"] = ETALON_PRICE_BY_CARD.get(new_card) or _manual_price(new_card)
    ln["src"] = "Вручну (варіант)"
    VARIANTS[h]["edits"].append({"rule": rule, "where": SPECS[sid]["name"], "pos": ln["pos"] or f"(позиція вузла) {SPECS[sid]['name']}",
                                 "was": f"{card_name(old_card)} × {r6(ln['q'])}", "now": f"{card_name(new_card)} × {r6(ln['q'])}",
                                 "why": why})
    return ln


def edit_set_qty(h, card, q_new, why, rule):
    vid = make_variant(h)
    paths = find_paths(vid, lambda ln: ln["kind"] == "Деталь" and ln["card"] == card)
    assert paths, (h, card)
    sid, i = own_spec(vid, paths[0], h)
    ln = SPECS[sid]["lines"][i]
    was = ln["q"]
    ln["q"] = Dd(q_new); ln["checked"] = False
    VARIANTS[h]["edits"].append({"rule": rule, "where": SPECS[sid]["name"], "pos": ln["pos"], "was": f"{card_name(card)} × {r6(was)}",
                                 "now": f"{card_name(card)} × {r6(ln['q'])}" if q_new else "рядок видалено", "why": why})
    if not q_new:
        del SPECS[sid]["lines"][i]
    return ln


def edit_delete_node(h, node_id, why, rule):
    vid = make_variant(h)
    paths = find_paths(vid, lambda ln: ln["kind"] == "Вузол" and ln["node"].split("@")[0] == node_id)
    assert paths, (h, node_id)
    sid, i = own_spec(vid, paths[0][:-1] + [paths[0][-1]], h) if len(paths[0]) > 1 else paths[0][0]
    ln = SPECS[sid]["lines"][i]
    VARIANTS[h]["edits"].append({"rule": rule, "where": SPECS[sid]["name"], "pos": ln["pos"], "was": f"Вузол «{SPECS[ln['node']]['name']}» × {r6(ln['q'])}",
                                 "now": "рядок-вузол видалено", "why": why})
    del SPECS[sid]["lines"][i]


def edit_add(h, node_id, card, q, why, rule, pos_n=None):
    vid = make_variant(h)
    if node_id == vid:
        sid = vid
    else:
        paths = find_paths(vid, lambda ln: ln["kind"] == "Вузол" and ln["node"].split("@")[0] == node_id)
        assert paths, (h, node_id)
        p = paths[0]
        own_spec(vid, p, h)
        sid = SPECS[p[-1][0]]["lines"][p[-1][1]]["node"]
        if not sid.endswith("@" + h):                    # узел-цель тоже копируем
            sid2 = sid + "@" + h
            if sid2 not in SPECS:
                base = SPECS[sid]
                SPECS[sid2] = {"id": sid2, "name": f"{base['name']} ({HNAME[h]})", "kind": "Вузол", "out_q": base["out_q"], "out_u": base["out_u"],
                               "lines": copy.deepcopy(base["lines"]), "base": sid, "base_lines": copy.deepcopy(base["lines"]), "house": h}
                VARIANTS[h]["copied_nodes"].append({"node": sid, "copy": sid2, "name": SPECS[sid2]["name"]})
            SPECS[p[-1][0]]["lines"][p[-1][1]]["node"] = sid2
            sid = sid2
    l = LINE.get(pos_n) if pos_n else None
    ln = new_line("Деталь", n=pos_n, card=card, q=Dd(q), unit=CARD[card]["unit"], ksrc="ЄдиниціЗбігаються", src="Додано в будинку",
                  note=why, sub="add" + h)
    if l is None:
        ln["pos"], ln["pos_key"] = "", ""
    ln["price"] = ETALON_PRICE_BY_CARD.get(card) or _manual_price(card)
    SPECS[sid]["lines"].append(ln)
    VARIANTS[h]["edits"].append({"rule": rule, "where": SPECS[sid]["name"], "pos": (l["pos"] + (" (нульовий рядок СС)" if l and not l["q"] else "")) if l else ("успадковує поз. 168 «Комплект септика»" if sid.split("@")[0] == "N12a" else "без позиції СС"),
                                 "was": "—", "now": f"{card_name(card)} × {r6(ln['q'])}", "why": why})
    return ln


def _manual_price(card):
    """Цена новой карточки варианта (M5 §4.5: бухгалтер вводит цену) — симуляция: последняя закупка з ПДВ."""
    p, basis, _ = price_last(card)
    return (p * VAT).quantize(Decimal("0.000001")) if p is not None else Z


ETALON_PRICE_BY_CARD = {}
for lf in ETALON_LEAVES:
    if lf["kind"] == "Деталь" and lf["card"] and lf["card"] not in ETALON_PRICE_BY_CARD:
        ETALON_PRICE_BY_CARD[lf["card"]] = lf["line"]["price"]
ETALON_CARDS = {lf["card"] for lf in ETALON_LEAVES if lf["kind"] == "Деталь" and lf["card"]}
GROUP_ONS = {K(lf["on"]) for lf in ETALON_LEAVES if lf["kind"].startswith("Група")}


def build_variants():
    # V1. Конструктивные варианты узла «Септик» (R2 §2.3; задача пользователя)
    edit_set_card("15-1", "00-00010526", "00-00005945", "на №1 септик з кільцем з днищем КЦД-15-9 замість днища ПН 1,5", "V1 конструктивний варіант вузла")
    vid = "V15-1"
    paths = find_paths(vid, lambda ln: ln["kind"] == "Деталь" and ln["card"] == "00000017651")
    sid, i = own_spec(vid, paths[0], "15-1")
    SPECS[sid]["lines"][i]["q"] = Decimal(1)
    VARIANTS["15-1"]["edits"].append({"rule": "V1 конструктивний варіант вузла", "where": SPECS[sid]["name"], "pos": "(позиція вузла) Септик",
                                      "was": "Кільце  КС 15-9 × 2", "now": "Кільце  КС 15-9 × 1", "why": "друге кільце — інша картка"})
    edit_add("15-1", "N12a", "00-00005944", 1, "друге кільце — «Кільце бетонне КС-15-9»", "V1 конструктивний варіант вузла")
    edit_add("15-2", "N12a", "00-00017426", 1, "на №2 у септику додатково кришка ПП 15-15 (уточнити: кришка понад плиту чи замість)", "V1 конструктивний варіант вузла")
    edit_delete_node("15-5", "N12a", "на №5 септик не списувався — будинок без септика (уточнити: якщо поставлять пізніше, варіант не потрібен)", "V1 конструктивний варіант вузла")
    # V3. «Різна комплектація будинків» за ПричиниЦеху (позиція є не на всіх будинках / різна кількість)
    for h in ("15-1", "15-3", "15-4", "15-5"):
        edit_set_qty(h, "00000021201", 2, "цех: «на №1 і №3-5 по 2 клапани при нормі 1 — різна комплектація»", "V3 різна комплектація (цех)")
    for h in ("15-2", "15-6"):
        edit_set_qty(h, "00000021201", 0, "цех: «на №2 і №6 потреби не було»", "V3 різна комплектація (цех)")
    for h in ("15-3", "15-4", "15-5", "15-6"):
        edit_add(h, "N10", "00-00019463", 1, "цех: змішувач лише на №3-6", "V3 різна комплектація (цех)", pos_n=186)
    for h, q in (("15-4", 2), ("15-5", 1), ("15-6", 1)):
        edit_add(h, "N10", "00-00015706", q, "цех: шланг 0,4 м лише на №4-6", "V3 різна комплектація (цех)", pos_n=137)
    for h in ("15-5", "15-6"):
        edit_add(h, "N05", "00-00004975", 1, "цех: мембрана дахова лише на №5-6", "V3 різна комплектація (цех)")
    for h, q in (("15-1", 1), ("15-2", 1), ("15-4", 2)):
        edit_add(h, "N11", "00-00005350", q, "цех: труба каналізаційна 50 мм лише на №1, №2, №4", "V3 різна комплектація (цех)")
    # V2. Заміна аналогом 1:1: карточка нормы на дом не списывалась, списана другая карточка той же ОН (нет в спецификации)
    auto = []
    for h in H15:
        vid = make_variant(h) if h in VARIANTS else None
        root = vid or "E"
        leaves, _ = explode(root)
        spec_cards = {lf["card"] for lf in leaves if lf["card"]}
        for lf in leaves:
            c = lf["card"]
            if lf["kind"] != "Деталь" or not c or FQ[c][h] > 0 or LQ[c][h] > 0:
                continue
            on = CARD.get(c, {}).get("on_code")
            if K(CARD.get(c, {}).get("on")) == "мшп":
                continue                                   # ОН «мшп» — сборная, аналогов по ней не ищем
            cands = [b for b in USED15 if CARD.get(b, {}).get("on_code") == on and b not in spec_cards and b not in ETALON_CARDS
                     and (FQ[b][h] + LQ[b][h]) > 0 and ukey_eq(UKEY_CARD.get(b), UKEY_CARD.get(c))]
            if not cands:
                continue
            b = max(cands, key=lambda x: FQ[x][h] + LQ[x][h])
            auto.append((h, c, b))
    global V2_REJECTED
    V2_REJECTED = [(h, c, b, REJECT_V2[(c, b)]) for h, c, b in auto if (c, b) in REJECT_V2]
    auto = [(h, c, b) for h, c, b in auto if (c, b) not in REJECT_V2]
    for h, c, b in auto:
        edit_set_card(h, c, b, f"на будинок списано «{card_name(b)}» ({float(FQ[b][h] + LQ[b][h]):g}), картку норми — ні", "V2 заміна аналогом 1:1")
    return auto


V2_AUTO = build_variants()
