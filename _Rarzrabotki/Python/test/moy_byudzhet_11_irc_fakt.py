# -*- coding: utf-8 -*-
"""Снимок №11 «факты IRC» для сценария 00009 v3 (материалы IRC методом потоков M05 §3.5). ТОЛЬКО ЧТЕНИЕ.

Базы: ERP (закупки, Казна-зеркало, остатки ERP), BuhBud (движения 20/22/28, остатки, приходы), Казна (поздние оплаты 30.09).
Срезы (промт v3 §1): товарный — момент снимка, все документы (периоды по ДАТАВРЕМЯ(2026,12,31,23,59,59), как f31/f32
исследования); денежный — Казна по 30.09.2026 23:59:59.

Логика перенесена из исследования `_Rarzrabotki/Python/test/ekonomiya_ss_irc/` (не написана заново):
  запросы f20 (закупки ERP), f31 (движения BuhBud), f33 (приходы BuhBud), f50 (Казна); vat_ratio — f40 (НДС партий);
  дефицит 30 м №3–№7 — f41 (Е)/(Ж) + FIFO по домам №4…№7; классы корзин — f32/f40.
Кэш исследования (out/core.json, analysis.json) используется только после сверки живых выборок с кэшем по регистраторам.

Выход: data/irc_fakt.json (stamp, srez, suppliers, kniga, kz, av, stocks, consumed, deficit, mshp_h, sens, controls).
Запуск (из папки скрипта): python moy_byudzhet_11_irc_fakt.py
"""
import os, sys, re, json, math, collections, datetime
from decimal import Decimal, ROUND_HALF_UP
from moy_byudzhet_common import *

RES = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\ekonomiya_ss_irc"
RES_OUT = os.path.join(RES, "out")
CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
CONN_KAZNA = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'
Z = Decimal(0)
T_MONEY = "2026-09-30 23:59:59"
HOUSES = ["30-3", "30-4", "30-5", "30-6", "30-7"]          # незавершённые 30 м (FIFO по номеру = порядок ship_real v2)
SCOPE8 = [f"15-{i}" for i in range(1, 7)] + ["30-1", "30-2"]
PROJECT_W = {"ІБ _МД IRC 2026", "ІС _МД IRC 2026", "МД IRС 2026", "Цех МД IRS 2026", "МД IRS 2026"}
NONTRADE = ("нова пошта", "автосвіфт", "(без контрагента)")


def Dd(x):
    return Decimal(str(x)) if x not in (None, "") else Z


def r2(x):
    return float(Dd(x).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def dv(a, b):
    return a / b if b else Z


def F(x):
    return fmt(r2(x))


def dts(d):
    """Дата 1С -> 'YYYY-MM-DD HH:MM:SS' (строкой самой 1С, без сдвига COM)."""
    a, b = S(d).split(" ")
    h, mi, se = b.split(":")
    return f"{a[6:10]}-{a[3:5]}-{a[0:2]} {int(h):02d}:{mi}:{se}"


def house(name):
    m = re.search(r"(15|30) м №\s*(\d+)", name or "")
    return f"{m.group(1)}-{m.group(2)}" if (m and "IR" in (name or "")) else None


def wcls(w):
    """Класс склада (как f40): P — склад проекта, H:15-1… — склад дома, MSP, X — чужой."""
    if not w:
        return "—"
    if w in PROJECT_W:
        return "P"
    h = house(w)
    if h:
        return "H:" + h
    if "МШП" in w:
        return "MSP"
    if "IR" in w:
        return "IRC?"
    return "X"


is_mat = lambda a: (a or "")[:2] in ("20", "22", "28")
K = lambda s: " ".join(str(s or "").replace("\xa0", " ").split()).casefold()


def nk(s):
    """Ключ контрагента (как f51: без кавычек, пробелы, регистр)."""
    return " ".join(str(s or "").replace("\xa0", " ").replace('"', "").split()).casefold()


LEGAL = {"тов", "фоп", "пнвп", "фірма", "компанія", "торгівельно-виробнича", "фо", "ооо", "азс"}


def toks(s):
    s = str(s or "").replace("\xa0", " ").replace('"', " ").replace("I", "І").replace("i", "і").casefold()
    return {t for t in re.split(r"[\s,]+", s) if t and t not in LEGAL}


# ════════════════════════════════════════════ живые выборки
_v8 = None


def v8():
    global _v8
    if _v8 is None:
        import win32com.client
        _v8 = win32com.client.Dispatch("V83.COMConnector")
    return _v8


def bq(B, text):
    z = B.NewObject("Запрос")
    z.Text = text
    try:
        return z.Execute().Выгрузить()
    except Exception as e:
        raise RuntimeError(f"запрос BuhBud/Казна упал: {e.excepinfo[2] if getattr(e, 'excepinfo', None) else e}\n{text}")


def sv(B, v):
    if v is None:
        return ""
    if isinstance(v, (int, float, str, bool)):
        return v
    try:
        return B.String(v)
    except Exception:
        return str(v)


ROOT_ERP = """ВЫБРАТЬ РАЗЛИЧНЫЕ СС.ПодразделениеЗакупки КАК Корень
ПОМЕСТИТЬ ВТ_Корень
ИЗ Справочник.А_СтруктураСебестоимости КАК СС
ГДЕ СС.Код В ("000000008", "000000009")
;
"""


def live_erp():
    out = {}
    t = query("""ВЫБРАТЬ С.Код КАК Код, С.КоличествоДомов КАК Домов, ЕСТЬNULL(СУММА(Ком.Сумма), 0) КАК Сум, КОЛИЧЕСТВО(Ком.НомерСтроки) КАК Строк
        ИЗ Справочник.А_СтруктураСебестоимости КАК С
            ЛЕВОЕ СОЕДИНЕНИЕ Справочник.А_СтруктураСебестоимости.Комплектующие КАК Ком ПО Ком.Ссылка = С.Ссылка
        ГДЕ С.Код В ("000000008", "000000009")
        СГРУППИРОВАТЬ ПО С.Код, С.КоличествоДомов""")
    out["cards"] = {S(t.Получить(i).Код).strip(): {"houses": int(t.Получить(i).Домов), "sum": float(t.Получить(i).Сум),
                                                   "rows": int(t.Получить(i).Строк)} for i in range(t.Количество())}
    # f20: закупки IRC (СебестоимостьТоваров, приход ПриобретениеТоваровУслуг), период — товарный срез
    t = query(ROOT_ERP + """ВЫБРАТЬ
        ВЫРАЗИТЬ(СТ.Регистратор КАК Документ.ПриобретениеТоваровУслуг).Номер КАК Номер, СТ.Период КАК Период,
        ВЫРАЗИТЬ(СТ.Регистратор КАК Документ.ПриобретениеТоваровУслуг).Контрагент КАК Контрагент,
        ВЫРАЗИТЬ(СТ.Регистратор КАК Документ.ПриобретениеТоваровУслуг).Контрагент.КодПоЕДРПОУ КАК ЕДРПОУ,
        ВЫРАЗИТЬ(СТ.Регистратор КАК Документ.ПриобретениеТоваровУслуг).Договор КАК Договор,
        ВЫРАЗИТЬ(СТ.Регистратор КАК Документ.ПриобретениеТоваровУслуг).ЗаказПоставщику КАК Заказ,
        СТ.АналитикаУчетаНоменклатуры.МестоХранения КАК Склад,
        СТ.АналитикаУчетаНоменклатуры.Номенклатура.Наименование КАК Ном,
        СТ.АналитикаУчетаНоменклатуры.Номенклатура.Код КАК НомКод,
        СТ.АналитикаУчетаНоменклатуры.Номенклатура.А_ОбщееНазваниеНоменклатуры.Наименование КАК ОН,
        СТ.Количество КАК Кол, СТ.Стоимость КАК Стоимость, СТ.СтоимостьБезНДС КАК БезНДС
    ИЗ РегистрНакопления.СебестоимостьТоваров КАК СТ
    ГДЕ СТ.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        И СТ.РазделУчета = ЗНАЧЕНИЕ(Перечисление.РазделыУчетаСебестоимостиТоваров.ТоварыНаСкладах)
        И СТ.Регистратор ССЫЛКА Документ.ПриобретениеТоваровУслуг
        И ВЫРАЗИТЬ(СТ.Регистратор КАК Документ.ПриобретениеТоваровУслуг).Подразделение В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Корень.Корень ИЗ ВТ_Корень КАК ВТ_Корень)
        И СТ.Период МЕЖДУ ДАТАВРЕМЯ(2000, 1, 1) И ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59)""")
    pu = []
    for i in range(t.Количество()):
        r = t.Получить(i)
        pu.append({"Номер": S(r.Номер), "Период": S(r.Период), "date": ymd(r.Период), "Контрагент": S(r.Контрагент),
                   "ЕДРПОУ": S(r.ЕДРПОУ).strip(), "Договор": S(r.Договор), "Заказ": S(r.Заказ), "Склад": S(r.Склад),
                   "Ном": S(r.Ном), "НомКод": S(r.НомКод), "ОН": S(r.ОН), "Кол": float(r.Кол),
                   "Стоимость": float(r.Стоимость), "БезНДС": float(r.БезНДС)})
    out["purch"] = pu
    # f50: Казна-зеркало IRC, статьи материалов/МШП и возвратов
    t = query(ROOT_ERP + """ВЫБРАТЬ Д.Период КАК Период, Д.СтатьяДвиженияДенежныхСредств.Код КАК Ст, Д.Сумма КАК Сумма,
        Д.ДокументДвиженияКазны КАК Док,
        ЕСТЬNULL(ВЫРАЗИТЬ(Д.ДокументДвиженияКазны КАК Документ.СписаниеБезналичныхДенежныхСредств).Контрагент,
            ЕСТЬNULL(ВЫРАЗИТЬ(Д.ДокументДвиженияКазны КАК Документ.РасходныйКассовыйОрдер).Контрагент,
                ВЫРАЗИТЬ(Д.ДокументДвиженияКазны КАК Документ.ПоступлениеБезналичныхДенежныхСредств).Контрагент)) КАК Контрагент,
        ВЫРАЗИТЬ(Д.ДокументДвиженияКазны КАК Документ.РасходныйКассовыйОрдер).ПодотчетноеЛицо КАК Подотчет,
        ВЫРАЗИТЬ(ВЫРАЗИТЬ(Д.ДокументДвиженияКазны КАК Документ.СписаниеБезналичныхДенежныхСредств).НазначениеПлатежа КАК СТРОКА(300)) КАК Назн,
        ЕСТЬNULL(ВЫРАЗИТЬ(Д.ДокументДвиженияКазны КАК Документ.СписаниеБезналичныхДенежныхСредств).Контрагент.КодПоЕДРПОУ,
            ВЫРАЗИТЬ(Д.ДокументДвиженияКазны КАК Документ.РасходныйКассовыйОрдер).Контрагент.КодПоЕДРПОУ) КАК ЕДРПОУ
    ИЗ РегистрНакопления.А_ДвиженияДенегИзКазны КАК Д
    ГДЕ (Д.Подразделение В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Корень.Корень ИЗ ВТ_Корень КАК ВТ_Корень)
            ИЛИ Д.ПодразделениеКазна В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Корень.Корень ИЗ ВТ_Корень КАК ВТ_Корень))
        И Д.СтатьяДвиженияДенежныхСредств.Код В ("УТ-001067", "УТ-001065", "УТ-001147", "УТ-001204")
        И Д.Период МЕЖДУ ДАТАВРЕМЯ(2025, 1, 1) И ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59)""")
    mo = []
    for i in range(t.Количество()):
        r = t.Получить(i)
        mo.append({"Период": S(r.Период), "dt": dts(r.Период), "Ст": S(r.Ст).strip(),
                   "Сумма": float(r.Сумма), "Док": S(r.Док), "Контрагент": S(r.Контрагент), "Подотчет": S(r.Подотчет),
                   "Назн": S(r.Назн), "ЕДРПОУ": S(r.ЕДРПОУ).strip()})
    out["money"] = mo
    t = query("ВЫБРАТЬ МАКСИМУМ(Д.Период) КАК МП ИЗ РегистрНакопления.А_ДвиженияДенегИзКазны КАК Д")
    mp = t.Получить(0).МП
    out["mirror_last"] = dts(mp)
    # I14 (в): остатки ERP (Стоимость с НДС) по местам хранения IRC на товарном срезе
    t = query("""ВЫБРАТЬ О.АналитикаУчетаНоменклатуры.МестоХранения КАК МХ, СУММА(О.СтоимостьОстаток) КАК С, СУММА(О.СтоимостьБезНДСОстаток) КАК СБ
        ИЗ РегистрНакопления.СебестоимостьТоваров.Остатки(ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59),
            РазделУчета = ЗНАЧЕНИЕ(Перечисление.РазделыУчетаСебестоимостиТоваров.ТоварыНаСкладах)) КАК О
        СГРУППИРОВАТЬ ПО О.АналитикаУчетаНоменклатуры.МестоХранения""")
    out["erp_stock"] = {}
    for i in range(t.Количество()):
        n = S(t.Получить(i).МХ)
        if "IR" in n:
            out["erp_stock"][n] = [float(t.Получить(i).С), float(t.Получить(i).СБ)]
    return out


def live_buh():
    B = v8().Connect(CONN_BUH)
    out = {}
    # f31: все движения Хозрасчетный 20/22/26/28 по складам IRC и партиям IRC, 2026 (товарный срез)
    text = open(os.path.join(RES, "f31_buh_moves.py"), encoding="utf-8").read().split('TEXT = """', 1)[1].split('"""', 1)[0]
    vt = bq(B, text)
    mv = []
    for i in range(vt.Количество()):
        r = vt.Получить(i)
        mv.append({"Период": str(r.Период)[:19], "Рег": sv(B, r.Регистратор), "ТипРег": sv(B, r.ТипРег), "СчДт": r.СчетДт, "СчКт": r.СчетКт,
                   "НомДт": sv(B, r.НомДт), "ПартДт": sv(B, r.ПартияДт), "СклДт": sv(B, r.СкладДт), "НомКт": sv(B, r.НомКт),
                   "ПартКт": sv(B, r.ПартияКт), "СклКт": sv(B, r.СкладКт), "Сумма": r.Сумма, "КолДт": r.КолДт, "КолКт": r.КолКт,
                   "Орг": sv(B, r.Орг), "КодДт": sv(B, r.КодДт), "ОНДт": sv(B, r.ОНДт), "КодКт": sv(B, r.КодКт), "ОНКт": sv(B, r.ОНКт),
                   "ЕдДт": sv(B, r.ЕдДт), "ЕдКт": sv(B, r.ЕдКт)})
    out["moves"] = mv
    # f33: приходы BuhBud (НДС партий, контрагент)
    text = open(os.path.join(RES, "f33_buh_receipts.py"), encoding="utf-8").read().split('TEXT = """', 1)[1].split('"""', 1)[0]
    vt = bq(B, text)
    rc = []
    for i in range(vt.Количество()):
        r = vt.Получить(i)
        net = Dd(r.Сумма) - (Dd(r.НДС) if r.ВклНДС else Z)
        rc.append({"Док": sv(B, r.Док), "Дата": str(r.Дата)[:19], "Контрагент": sv(B, r.Контрагент), "Код": sv(B, r.Код),
                   "СтавкаНом": sv(B, r.СтавкаНом), "Проведен": bool(r.Проведен), "Нетто": str(net), "НДС": str(Dd(r.НДС))})
    out["receipts"] = rc
    # остатки 20/22/28 по складам на товарный срез (как v02_buh)
    vt = bq(B, """ВЫБРАТЬ П.Ссылка КАК Счет ПОМЕСТИТЬ ВТ_Счета ИЗ ПланСчетов.Хозрасчетный КАК П ГДЕ П.Код В ("20", "22", "28")
        ;
        ВЫБРАТЬ О.Субконто3 КАК Склад, СУММА(О.СуммаОстаток) КАК Сумма
        ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59),
            Счет В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Счета.Счет ИЗ ВТ_Счета КАК ВТ_Счета), , ) КАК О
        СГРУППИРОВАТЬ ПО О.Субконто3""")
    out["bal"] = {}
    for i in range(vt.Количество()):
        n = sv(B, vt.Получить(i).Склад)
        if "IR" in str(n) or "МШП" in str(n):
            out["bal"][n] = float(vt.Получить(i).Сумма)
    return out


def live_kazna(mirror_last):
    """Поздние оплаты 30.09 после последней записи зеркала (база kazna, РН.БДДС, только чтение)."""
    try:
        Kz = v8().Connect(CONN_KAZNA)
    except Exception as e:
        return {"status": f"недоступна: {e}", "late": None}
    y, m, d = mirror_last[:10].split("-")
    hh, mi, ss = mirror_last[11:19].split(":")
    lit = f"ДАТАВРЕМЯ({int(y)}, {int(m)}, {int(d)}, {int(hh)}, {int(mi)}, {int(ss)})"
    vt = bq(Kz, f"""ВЫБРАТЬ П.Ссылка КАК Подр ПОМЕСТИТЬ ВТ_П ИЗ Справочник.Подразделения КАК П
        ГДЕ П.Наименование ПОДОБНО "%IRC 2026%" ИЛИ П.Наименование ПОДОБНО "%IRС 2026%"
        ;
        ВЫБРАТЬ Б.СтатьяДвиженияДенежныхСредств.Код КАК Ст,
            ВЫБОР КОГДА Б.Период <= {lit} ТОГДА "1_до_зеркала" КОГДА Б.Период <= ДАТАВРЕМЯ(2026, 9, 30, 23, 59, 59) ТОГДА "2_поздние_30_09"
                ИНАЧЕ "3_после_30_09" КОНЕЦ КАК Окно,
            СУММА(Б.СуммаРасход - Б.СуммаПриход) КАК Сум, КОЛИЧЕСТВО(*) КАК N
        ИЗ РегистрНакопления.БДДС КАК Б
        ГДЕ НЕ Б.ПланФакт И Б.Подразделение В ИЕРАРХИИ (ВЫБРАТЬ ВТ_П.Подр ИЗ ВТ_П КАК ВТ_П)
            И Б.СтатьяДвиженияДенежныхСредств.Код В ("УТ-001067", "УТ-001065")
            И Б.Период МЕЖДУ ДАТАВРЕМЯ(2026, 1, 1) И ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59)
        СГРУППИРОВАТЬ ПО Б.СтатьяДвиженияДенежныхСредств.Код,
            ВЫБОР КОГДА Б.Период <= {lit} ТОГДА "1_до_зеркала" КОГДА Б.Период <= ДАТАВРЕМЯ(2026, 9, 30, 23, 59, 59) ТОГДА "2_поздние_30_09"
                ИНАЧЕ "3_после_30_09" КОНЕЦ""")
    res = collections.defaultdict(float)
    for i in range(vt.Количество()):
        r = vt.Получить(i)
        res[f"{sv(Kz, r.Ст).strip()}|{r.Окно}"] += float(r.Сум)
    return {"status": "доступна (kazna, РН.БДДС, факт)", "windows": dict(res),
            "late": {a: res.get(f"{a}|2_поздние_30_09", 0.0) for a in ("УТ-001067", "УТ-001065")}}


# ════════════════════════════════════════════ сверка с кэшем исследования (по регистраторам)
def cache_diff(live, cache, keyf):
    a = collections.Counter(keyf(r) for r in live)
    b = collections.Counter(keyf(r) for r in cache)
    return list((a - b).elements()), list((b - a).elements())


def rc_load(name):
    return json.load(open(os.path.join(RES_OUT, name), encoding="utf-8"))


# ════════════════════════════════════════════ НДС партий (перенос f40_core.vat_ratio)
class Vat:
    def __init__(self, receipts, moves):
        self.rc, self.rcdoc, self.rate_nom = collections.defaultdict(lambda: [Z, Z]), collections.defaultdict(lambda: [Z, Z]), {}
        for r in receipts:
            self.rc[(r["Док"], r["Код"])][0] += Dd(r["Нетто"]); self.rc[(r["Док"], r["Код"])][1] += Dd(r["НДС"])
            self.rcdoc[r["Док"]][0] += Dd(r["Нетто"]); self.rcdoc[r["Док"]][1] += Dd(r["НДС"])
            self.rate_nom[r["Код"]] = r["СтавкаНом"]
        self.src = collections.defaultdict(list)
        for r in moves:
            if is_mat(r["СчДт"]) and is_mat(r["СчКт"]):
                self.src[(r["ПартДт"], r["КодДт"], r["СклДт"])].append((r["ПартКт"], r["КодКт"], r["СклКт"], Dd(r["Сумма"])))
        self.memo, self.unresolved, self.fallback, self.taint = {}, collections.defaultdict(lambda: [0, Z]), set(), {}

    def ratio(self, part, code, skl, depth=0, amount=Z):
        key = (part, code, skl)
        if key in self.memo:
            return self.memo[key]
        res, taint = None, False
        if (part, code) in self.rc and self.rc[(part, code)][0] != 0:
            res = self.rc[(part, code)][1] / self.rc[(part, code)][0]
        elif part in self.rcdoc and self.rcdoc[part][0] != 0:
            res = self.rcdoc[part][1] / self.rcdoc[part][0]
        elif part.startswith("Авансовий звіт"):
            res = Z
        elif key in self.src and depth < 12:
            tot = sum(s for *_, s in self.src[key])
            if tot:
                res = sum(s * self.ratio(p, c, w, depth + 1) for p, c, w, s in self.src[key]) / tot
                taint = any(self.taint.get((p, c, w), False) for p, c, w, _ in self.src[key])
        if res is None:
            rt = self.rate_nom.get(code)
            res = Decimal("0.2") if rt in (None, "20%") else (Decimal("0.07") if rt == "7%" else Z)
            self.unresolved[part.split(" ")[0]][0] += 1
            self.fallback.add(key)
            taint = True
        self.taint[key] = taint
        self.memo[key] = res
        return res


# ════════════════════════════════════════════ main
if __name__ == "__main__":
    stamp = snapshot_stamp()
    print(f"Снимок №11 «факты IRC», {stamp}. Только чтение.")
    L = live_erp()
    print(f"ERP: закупок {len(L['purch'])}, движений Казны {len(L['money'])}, зеркало Казны — последняя запись {L['mirror_last']}")
    LB = live_buh()
    print(f"BuhBud: движений {len(LB['moves'])}, строк приходов {len(LB['receipts'])}")
    KZN = live_kazna(L["mirror_last"])
    print(f"Казна: {KZN['status']}; поздние оплаты 30.09 (после зеркала): {KZN.get('late')}")

    # ---------- сверка живых выборок с кэшем исследования (по регистратору/строке, не по дате)
    C_PU, C_MV, C_MO = rc_load("purch_erp.json"), rc_load("buh_moves.json"), rc_load("money_kazna.json")
    pu_new, pu_gone = cache_diff(L["purch"], C_PU, lambda r: (r["Номер"], r["НомКод"], round(r["Кол"], 3), round(r["Стоимость"], 2)))
    mv_key = lambda r: (r["Период"], r["Рег"], r["СчДт"], r["СчКт"], r["КодДт"], r["КодКт"], r["СклДт"], r["СклКт"], r["ПартДт"], r["ПартКт"],
                        round(float(r["Сумма"] or 0), 2))
    mv_new, mv_gone = cache_diff(LB["moves"], C_MV, mv_key)
    mo_live = [r for r in L["money"] if r["Ст"] in ("УТ-001067", "УТ-001065") and r["dt"] >= "2026-01-01"]
    mo_cache = [r for r in C_MO if r["СтКод"] in ("УТ-001067", "УТ-001065")]
    mo_new, mo_gone = cache_diff(mo_live, [{"Док": r["Док"], "Сумма": r["Сумма"], "Ст": r["СтКод"]} for r in mo_cache],
                                 lambda r: (r["Док"], r["Ст"], round(float(r["Сумма"]), 2)))
    cache_ok = not (pu_new or pu_gone or mv_new or mv_gone or mo_new or mo_gone)
    print(f"Сверка с кэшем исследования: закупки +{len(pu_new)}/−{len(pu_gone)}, движения BuhBud +{len(mv_new)}/−{len(mv_gone)}, "
          f"Казна +{len(mo_new)}/−{len(mo_gone)} → {'кэш актуален' if cache_ok else 'КЭШ УСТАРЕЛ'}")
    if not cache_ok:
        for x in (pu_new + pu_gone)[:10] + (mv_new + mv_gone)[:20] + (mo_new + mo_gone)[:10]:
            print("   ", x)
        sys.exit("СТОП: живая база разошлась с кэшем исследования — перезапустить конвейер f20 → f30/f31 → f32 → f40 → f41 (промт §3.1)")

    CORE = rc_load("core.json")
    AN = rc_load("analysis.json")
    ON = {k: collections.defaultdict(Decimal, {m: Dd(v) for m, v in d.items()}) for k, d in CORE["on"].items()}
    NAME = CORE["names"]
    SS = rc_load("ss_erp.json")
    MV = LB["moves"]
    VAT = Vat(LB["receipts"], MV)
    controls = []

    def ctl(n, name, research, mine, src, ok, note=""):
        controls.append({"n": n, "name": name, "research": research, "mine": mine, "src": src, "ok": bool(ok), "note": note})

    # ---------- 1. нормы СС
    c8, c9 = L["cards"]["000000008"], L["cards"]["000000009"]
    norm5 = r2(Dd(c9["sum"]) * 5)
    ctl(1, "Норма СС 008 / 009 / 30 м №3–№7", "704 612,01 × 6 (190 строк); 1 009 864,74 × 7 (170); 5 049 323,70",
        f"{F(c8['sum'])} × {c8['houses']} ({c8['rows']}); {F(c9['sum'])} × {c9['houses']} ({c9['rows']}); {F(norm5)}",
        "ERP Справочник.А_СтруктураСебестоимости.Комплектующие",
        abs(c8["sum"] - 704612.01) < 0.005 and abs(c9["sum"] - 1009864.74) < 0.005 and c8["rows"] == 190 and c9["rows"] == 170
        and abs(norm5 - 5049323.70) < 0.005)

    # ---------- 2. закупки
    PU = L["purch"]
    pu_le = [r for r in PU if r["date"] <= "2026-09-30"]
    pu_after = [r for r in PU if r["date"] > "2026-09-30"]
    by_m = collections.defaultdict(Decimal)
    for r in pu_le:
        by_m[r["date"][:7]] += Dd(r["Стоимость"])
    pu_g, pu_n = sum(Dd(r["Стоимость"]) for r in PU), sum(Dd(r["БезНДС"]) for r in PU)
    nsup = len({nk(r["Контрагент"]) for r in PU})
    ctl(2, "Закупки IRC ERP (брутто / нетто, строк, поставщиков, по месяцам, после 30.09)",
        "9 053 180,08 / 7 634 660,50; 493; 47; 07 4 603 160,77, 08 3 640 995,44, 09 809 023,87; после 30.09 — 0",
        f"{F(pu_g)} / {F(pu_n)}; {len(PU)}; {nsup}; " + ", ".join(f"{m[5:]} {F(v)}" for m, v in sorted(by_m.items()))
        + f"; после 30.09 — {len(pu_after)} строк {F(sum(Dd(r['Стоимость']) for r in pu_after))}",
        "ERP РН.СебестоимостьТоваров (f20), период по ДАТАВРЕМЯ(2026,12,31,23,59,59)",
        abs(pu_g - Decimal("9053180.08")) < Decimal("0.005") and len(PU) == 493 and nsup == 47)

    # ---------- 3. оплаты Казны
    MO = [r for r in L["money"] if r["Ст"] in ("УТ-001067", "УТ-001065")]
    pay = collections.defaultdict(Decimal)
    for r in MO:
        if r["dt"] <= T_MONEY:
            pay[(r["Ст"], r["dt"][:7])] += -Dd(r["Сумма"])
    p067 = sum(v for (a, m), v in pay.items() if a == "УТ-001067")
    p065 = sum(v for (a, m), v in pay.items() if a == "УТ-001065")
    late = KZN.get("late") or {}
    ctl(3, "Оплаты Казны IRC 067 / 065 по 30.09 23:59:59; последняя запись зеркала; поздние 30.09",
        "067 7 555 928,55 (3 590 421,86 / 1 193 223,28 / 2 772 283,41); 065 42 445,53 (14 549,08 / 27 896,45); зеркало 30.09 15:36",
        f"067 {F(p067)} (" + " / ".join(F(pay[('УТ-001067', m)]) for m in ("2026-07", "2026-08", "2026-09")) + f"); 065 {F(p065)} ("
        + " / ".join(F(pay[('УТ-001065', m)]) for m in ("2026-07", "2026-08")) + f"); зеркало {L['mirror_last']}; Казна: {KZN['status']}, "
        f"поздние 30.09: 067 {F(late.get('УТ-001067', 0))}, 065 {F(late.get('УТ-001065', 0))}",
        "ERP РН.А_ДвиженияДенегИзКазны + kazna РН.БДДС (факт)",
        abs(p067 - Decimal("7555928.55")) < Decimal("0.005") and abs(p065 - Decimal("42445.53")) < Decimal("0.005"))

    # ---------- 5–7. остатки и списание BuhBud (живые, нетто)
    bal = collections.defaultdict(Decimal)
    for w, s in LB["bal"].items():
        c = wcls(w)
        bal["P" if c == "P" else (c[2:] if c.startswith("H:") else c)] += Dd(s)
    cons_n = collections.defaultdict(Decimal)
    for r in MV:
        if is_mat(r["СчКт"]) and not is_mat(r["СчДт"]):
            c = wcls(r["СклКт"])
            if c.startswith("H:"):
                cons_n[c[2:]] += Dd(r["Сумма"])
    P_net, P_gt = bal["P"], sum(o["stk|P|gt"] for o in ON.values())
    ctl(5, "Склад проекта «ІБ _МД IRC 2026» нетто / брутто", "3 044 261,07 / 3 631 218,04",
        f"{F(P_net)} / {F(P_gt)}", "BuhBud Хозрасчетный.Остатки 20/22/28 на товарный срез; брутто — НДС партий (кэш f40)",
        abs(P_net - Decimal("3044261.07")) < Decimal("0.005") and abs(P_gt - Decimal("3631218.04")) < Decimal("0.005") * 3631218)
    built_h = {h: bal[h] for h in ("15-4", "30-2", "30-1")}
    built_gt = sum(o[f"stkh|{h}|gt"] for o in ON.values() for h in SCOPE8)
    ctl(6, "Склады домов нетто: 30 м №3; построенные (15 м №4 / 30 м №2 / 30 м №1); брутто построенных",
        "164 734,64; 114 166,64 (30 768,30 / 83 051,38 / 346,96); 136 988,17",
        f"{F(bal['30-3'])}; {F(sum(bal[h] for h in SCOPE8))} ({' / '.join(F(v) for v in built_h.values())}); {F(built_gt)}",
        "BuhBud Хозрасчетный.Остатки по складам домов",
        abs(bal["30-3"] - Decimal("164734.64")) < Decimal("0.005") and abs(sum(bal[h] for h in SCOPE8) - Decimal("114166.64")) < Decimal("0.005"))
    cons8_n = sum(cons_n[h] for h in SCOPE8)
    cons8_gt = sum(o[f"c|{h}|gt"] for o in ON.values() for h in SCOPE8)
    c3_gt = sum(o["c|30-3|gt"] for o in ON.values())
    ratios = {h: float(dv(sum(o[f"c|{h}|gt"] for o in ON.values()), cons_n[h])) for h in SCOPE8}
    ctl(7, "Списано 8 домов нетто / брутто (% нормы); 30 м №3; брутто/нетто по домам 1,17–1,19",
        "4 232 193,58 / 5 001 128,05 (80,05 %); 100 669,28 / 120 727,54",
        f"{F(cons8_n)} / {F(cons8_gt)} ({float(cons8_gt / Decimal('6247401.54') * 100):.2f} %); {F(cons_n['30-3'])} / {F(c3_gt)}; "
        f"k по домам {min(ratios.values()):.4f}…{max(ratios.values()):.4f}",
        "BuhBud ДвиженияССубконто Кт 20/22/28 → Дт 23x/26 по складам домов (вкл. 01.10)",
        abs(cons8_n - Decimal("4232193.58")) < Decimal("0.005") and abs(cons_n["30-3"] - Decimal("100669.28")) < Decimal("0.005"))

    # ---------- 9. дефицит из кэша (f41 Е/Ж) + FIFO №4…№7
    def deficit_cache():
        out = collections.defaultdict(Decimal)
        own_used = P_used = Z
        for k, o in ON.items():
            nq, ns = o["nr|30-1|ПланКол"], o["nr|30-1|ПланГрн"]
            pss = dv(ns, nq)
            own_q, own_g, P_q, P_g = o["stkh|30-3|q"], o["stkh|30-3|gt"], o["stk|P|q"], o["stk|P|gt"]
            if nq <= 0 and ns > 0:
                out["30-3"] += max(Z, ns - o["c|30-3|gt"])
                for h in HOUSES[1:]:
                    out[h] += ns
                continue
            if nq <= 0:
                continue
            need = max(Z, nq - o["c|30-3|q"])
            f_own = min(need, max(Z, own_q)); own_used += f_own * dv(own_g, own_q)
            need -= f_own
            f_P = min(need, max(Z, P_q)); P_used += f_P * dv(P_g, P_q)
            out["30-3"] += (need - f_P) * pss
            rem = max(Z, P_q - f_P)
            for h in HOUSES[1:]:
                take = min(nq, rem); rem -= take; P_used += take * dv(P_g, P_q)
                out[h] += (nq - take) * pss
        return out, own_used, P_used
    DC, own_used_c, P_used_c = deficit_cache()
    G_cache = sum(DC.values())
    an_tot = Dd(AN["tot"]["e303_def"]) + Dd(AN["tot"]["e47_def"])
    ref_h = {"30-3": "63328.74", "30-4": "138288.86", "30-5": "262283.02", "30-6": "369608.35", "30-7": "460372.81"}
    ok9a = abs(G_cache - Decimal("1293881.78")) < Decimal("0.005") and abs(G_cache - an_tot) < Decimal("0.005") \
        and all(abs(DC[h] - Dd(v)) < Decimal("0.005") for h, v in ref_h.items())

    # грубая независимая оценка (как v3_deficit прошлой сессии): ОН карточки 009 × 5 против остатка «ІБ _МД IRC 2026» +
    # склад 30 м №3 + списано на 30 м №3 (кол-во, ОН номенклатуры BuhBud), без алиасов
    norm9 = collections.defaultdict(lambda: [Z, Z])
    for r in SS["kom"]:
        if "30 м" in r["СС"]:
            nn = norm9[K(r["ОН"]) or ("номсс|" + K(r["НомСС"]))]
            nn[0] += Dd(r["Кол"]); nn[1] += Dd(r["Сумма"])
    avail = collections.defaultdict(lambda: [Z, Z])
    for r in MV:
        dt_, kt_ = r["СчДт"] or "", r["СчКт"] or ""
        s = Dd(r["Сумма"])
        for side, sign in (("Дт", 1), ("Кт", -1)):
            acc = dt_ if side == "Дт" else kt_
            w = r["Скл" + side]
            if is_mat(acc) and (wcls(w) == "P" or wcls(w) == "H:30-3"):
                a = avail[K(r["ОН" + side])]
                a[0] += sign * Dd(r["Кол" + side] or 0); a[1] += sign * s
        if is_mat(kt_) and not is_mat(dt_) and wcls(r["СклКт"]) == "H:30-3":
            a = avail[K(r["ОНКт"])]
            a[0] += Dd(r["КолКт"] or 0); a[1] += s
    rough = Z
    for k, (nq, ns) in norm9.items():
        aq, an_ = avail.get(k, [Z, Z])
        if aq <= 0:
            rough += ns * 5
        elif nq > 0:
            rough += max(Z, nq * 5 - aq) * (ns / nq)
        else:
            rough += max(Z, ns * 5 - an_ * Decimal("1.2"))
    rough_dev = float((rough - G_cache) / G_cache * 100)
    ctl(9, "Дефицит к норме 30 м №3–№7 из кэша (f41 Е/Ж) и по домам FIFO; грубая оценка ±10 %",
        "1 293 881,78 (№3 63 328,74; №4 138 288,86; №5 262 283,02; №6 369 608,35; №7 460 372,81); грубая 1 373 091,78 (+6,1 %)",
        f"{F(G_cache)} (" + "; ".join(f"№{h[3:]} {F(DC[h])}" for h in HOUSES) + f"); Σ домов − итог = {F(sum(DC.values()) - an_tot)}; "
        f"грубая {F(rough)} ({rough_dev:+.1f} %)",
        "кэш out/core.json логикой f41 (Е)/(Ж) + FIFO; грубая — живые BuhBud движения по ОН × норма 009",
        ok9a and abs(rough_dev) <= 10)

    # ---------- 10. тождество §5.2
    surplus_c = Dd(AN["tot"]["P_left_after"])
    norm5_d = Dd(c9["sum"]) * 5
    cheap = norm5_d - c3_gt - own_used_c - P_used_c - G_cache
    under_norm = P_used_c
    ident = norm5_d - (c3_gt + own_used_c + under_norm + G_cache + cheap)
    ctl(10, "Тождество §5.2: норма = списано №3 + склад №3 + со склада проекта + дефицит + удешевление; склад проекта = под норму + излишек",
        "120 727,54 + 196 958,00 + 2 979 585,39 + 1 293 881,78 + 458 170,99 = 5 049 323,70; 2 979 585,39 + 651 632,65 = 3 631 218,04",
        f"{F(c3_gt)} + {F(own_used_c)} + {F(under_norm)} + {F(G_cache)} + {F(cheap)} = {F(norm5_d)}; "
        f"{F(under_norm)} + {F(surplus_c)} = {F(under_norm + surplus_c)} (склад {F(P_gt)}, невязка {F(P_gt - under_norm - surplus_c)})",
        "кэш core/analysis", abs(cheap - Decimal("458170.99")) <= Decimal("0.05") and abs(P_gt - under_norm - surplus_c) <= Decimal("0.05"))

    # ════════════════════════════════════ ДЕФИЦИТ v3: свёртка алиасов и состава (промт §8), неттинг по позициям
    # остатки и списание по номенклатуре (живые движения = кэш)
    nom = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z]))   # nom[(ключ ОН, номенклатура)][класс] = [кол, нетто]
    vat_g = collections.defaultdict(Decimal)                                        # (ключ, ном, класс) -> брутто остатка
    for r in MV:
        dt_, kt_ = r["СчДт"] or "", r["СчКт"] or ""
        s = Dd(r["Сумма"])
        if is_mat(dt_):
            k_, c_ = (K(r["ОНДт"]) or ("ном|" + K(r["НомДт"])), r["НомДт"]), wcls(r["СклДт"])
            nom[k_][c_][0] += Dd(r["КолДт"] or 0); nom[k_][c_][1] += s
            vat_g[(k_, c_)] += s * (1 + VAT.ratio(r["ПартДт"], r["КодДт"], r["СклДт"]))
        if is_mat(kt_):
            k_, c_ = (K(r["ОНКт"]) or ("ном|" + K(r["НомКт"])), r["НомКт"]), wcls(r["СклКт"])
            nom[k_][c_][0] -= Dd(r["КолКт"] or 0); nom[k_][c_][1] -= s
            vat_g[(k_, c_)] -= s * (1 + VAT.ratio(r["ПартКт"], r["КодКт"], r["СклКт"]))
            if not is_mat(dt_):
                nom[k_]["c:" + c_][0] += Dd(r["КолКт"] or 0); nom[k_]["c:" + c_][1] += s

    def stock(key, cls, pred=lambda n: True):
        q = sum(v[cls][0] for (k_, n), v in nom.items() if k_ == key and pred(n))
        g = sum(vat_g[((k_, n), cls)] for (k_, n) in nom if k_ == key and pred(n))
        return q, g

    def cons(key, house_, pred=lambda n: True):
        return sum(v["c:H:" + house_][0] for (k_, n), v in nom.items() if k_ == key and pred(n))

    def net(nq, pss, c3q, own, pools):
        """Неттинг одной позиции: №3 — свой склад, затем пулы склада проекта; №4…№7 — FIFO из остатка пулов.
        pools: [(кол, брутто, метка)]. Возврат: дефицит (кол) по домам, использовано из своего склада/пулов (брутто), остаток пулов."""
        pools = [[max(Z, q), dv(g, q) if q > 0 else Z, lab] for q, g, lab in pools]
        d, used = {}, collections.defaultdict(Decimal)
        need = max(Z, nq - c3q)
        f_own = min(need, max(Z, own[0])); used["own"] += f_own * dv(own[1], own[0]); need -= f_own
        for h in HOUSES:
            if h != "30-3":
                need = nq
            for p in pools:
                take = min(need, p[0]); p[0] -= take; need -= take; used[p[2]] += take * p[1]
            d[h] = need
        return d, used, pools

    items, adjust = [], []
    G = collections.defaultdict(Decimal)
    used_tot = collections.defaultdict(Decimal)
    left_pools = Z
    SPECIAL = {"гайка", "шпилька", "комплект септика", "кріплення ринви", "з'єднувач компресійний"}
    cat30 = {k: r.get("d30_cat") for k, r in AN["rows"].items()}

    def add_item(name, key, nq, pss, d, flags):
        it = {"name": name, "key": key, "nq": float(nq), "price_ss": r2(pss), "def_q": {h: float(v) for h, v in d.items()},
              "def": {h: r2(v * pss) for h, v in d.items()}, **flags}
        it["sum"] = r2(sum(Dd(v) for v in it["def"].values()))
        for h, v in d.items():
            G[h] += v * pss
        if it["sum"] > 0:
            items.append(it)
        return it

    for k, o in ON.items():
        nq, ns = o["nr|30-1|ПланКол"], o["nr|30-1|ПланГрн"]
        if k in SPECIAL or (nq <= 0 and ns <= 0):
            continue
        pss = dv(ns, nq)
        flags = {"vne_smety": False, "sostav_zero": cat30.get(k) == "не использовано", "mshp": False, "alias_of": None,
                 "card": "000000009", "n": 2, "Z": "2/5", "consumed3_q": float(o["c|30-3|q"]), "own3_q": float(o["stkh|30-3|q"]),
                 "stockP_q": float(o["stk|P|q"])}
        if nq <= 0:       # план суммой без количества (f41)
            d = {"30-3": max(Z, ns - o["c|30-3|gt"]), **{h: ns for h in HOUSES[1:]}}
            it = add_item(NAME[k], k, Decimal(1), Decimal(1), d, flags)
            continue
        d, used, pools = net(nq, pss, o["c|30-3|q"], (o["stkh|30-3|q"], o["stkh|30-3|gt"]), [(o["stk|P|q"], o["stk|P|gt"], "P")])
        for kk, v in used.items():
            used_tot[kk] += v
        add_item(NAME[k], k, nq, pss, d, flags)

    # --- специальные позиции (промт §8)
    lines9 = [r for r in SS["kom"] if "30 м" in r["СС"]]
    # Гайка: норма по размерам; запас неттится только того же размера (М48 — только явный М48)
    size = lambda s: (re.search(r"[МмMm]\s?(\d+)", s or "") or [None, "?"])[1]
    g_lines = collections.defaultdict(lambda: [Z, Z])
    for r in lines9:
        if K(r["ОН"]) == "гайка":
            g_lines[size(r["НомСС"])][0] += Dd(r["Кол"]); g_lines[size(r["НомСС"])][1] += Dd(r["Сумма"])
    g_sum = Z
    for sz, (q, s) in sorted(g_lines.items()):
        pred = lambda n, sz=sz: size(n) == sz
        own = stock("гайка", "H:30-3", pred); Pp = stock("гайка", "P", pred)
        d, used, pools = net(q, s / q, cons("гайка", "30-3", pred), own, [(Pp[0], Pp[1], "P")])
        for kk, v in used.items():
            used_tot[kk] += v
        it = add_item(f"Гайка М{sz}", "гайка", q, s / q, d, {"vne_smety": False, "sostav_zero": False, "mshp": False, "alias_of": None,
                                                            "card": "000000009", "n": 2, "Z": "2/5", "rule": "розмір М — лише такий самий розмір"})
        g_sum += Dd(it["sum"])
    # Шпилька: норма М48; запас — только явно М48 (Стріжень М48 DIN 975), «Шпилька 10х1000» не подходит (M05 §3.3 усл. 4)
    o = ON["шпилька"]
    pred48 = lambda n: "48" in (n or "")
    own = stock("шпилька", "H:30-3", pred48); Pp = stock("шпилька", "P", pred48)
    d, used, pools = net(o["nr|30-1|ПланКол"], dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"]), cons("шпилька", "30-3", pred48), own,
                         [(Pp[0], Pp[1], "P")])
    for kk, v in used.items():
        used_tot[kk] += v
    sh_it = add_item("Шпилька М48", "шпилька", o["nr|30-1|ПланКол"], dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"]), d,
                     {"vne_smety": False, "sostav_zero": False, "mshp": False, "alias_of": None, "card": "000000009", "n": 2, "Z": "2/5",
                      "rule": "лише явно М48"})
    # Комплект септика: состав из факта (15 м №2–№4, №6; 30 м №2): днище 1, кільце КС 15-9 2, люк 1, плита перекриття 1
    COMP = {"днище": 1, "кільце": 2, "люк": 1, "плита перекриття": 1}
    comp_of = lambda n: next((c for c in COMP if c in K(n)), None)
    evid = {}
    for h in ("15-1", "15-2", "15-3", "15-4", "15-5", "15-6", "30-1", "30-2"):
        evid[h] = {c: float(cons("комплект септика", h, lambda n, c=c: comp_of(n) == c)) for c in COMP}
    confirmed = [h for h, e in evid.items() if all(e[c] == COMP[c] for c in COMP)]
    sets = lambda cls: min((stock("комплект септика", cls, lambda n, c=c: comp_of(n) == c)[0] / COMP[c]) for c in COMP)
    P_sets = Decimal(math.floor(sets("P"))) if len(confirmed) >= 3 else Z
    P_sets_g = sum(stock("комплект септика", "P", lambda n, c=c: comp_of(n) == c)[1] / max(Decimal(1), stock("комплект септика", "P", lambda n, c=c: comp_of(n) == c)[0]) * COMP[c] for c in COMP) * P_sets
    own_sets = Decimal(math.floor(sets("H:30-3"))) if len(confirmed) >= 3 else Z
    o = ON["комплект септика"]
    s_pss = dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"])
    d, used, pools = net(o["nr|30-1|ПланКол"], s_pss, Z, (own_sets, Z), [(P_sets, P_sets_g, "P")])
    for kk, v in used.items():
        used_tot[kk] += v
    sept_left_g = stock("комплект септика", "P")[1] - P_sets_g
    add_item("Комплект септика (комплектами за складом з факту)", "комплект септика", o["nr|30-1|ПланКол"], s_pss, d,
             {"vne_smety": False, "sostav_zero": False, "mshp": False, "alias_of": None, "card": "000000009", "n": 2, "Z": "2/5",
              "rule": f"склад: {int(P_sets)} компл. (днище 1, кільце 2, люк 1, плита 1 — підтверджено фактом {', '.join(confirmed)}); "
                      f"залишок деталей {F(sept_left_g)} — реєстр надлишків"})
    # кріплення ринви ← держак ринви (пласт.) со склада проекта
    o = ON["кріплення ринви"]
    pr_r = lambda n: "ринви" in K(n)
    al = stock("держак ринви", "P", pr_r)
    d, used, pools = net(o["nr|30-1|ПланКол"], dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"]), o["c|30-3|q"],
                         (o["stkh|30-3|q"], o["stkh|30-3|gt"]), [(o["stk|P|q"], o["stk|P|gt"], "P"), (al[0], al[1], "P:alias")])
    for kk, v in used.items():
        used_tot[kk] += v
    add_item("кріплення ринви", "кріплення ринви", o["nr|30-1|ПланКол"], dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"]), d,
             {"vne_smety": False, "sostav_zero": False, "mshp": False, "alias_of": "Держак ринви Profil пласт.", "card": "000000009", "n": 2, "Z": "2/5"})
    alias_used = {"держак ринви": used.get("P:alias", Z)}
    # з'єднувач компресійний ← муфта ПЕ 25 (компресійна, д.25-3/4)
    o = ON["з'єднувач компресійний"]
    pr_m = lambda n: "25" in (n or "") and ("3/4" in (n or "") or "пе" in K(n))
    al = stock("муфта", "P", pr_m)
    d, used, pools = net(o["nr|30-1|ПланКол"], dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"]), o["c|30-3|q"],
                         (o["stkh|30-3|q"], o["stkh|30-3|gt"]), [(o["stk|P|q"], o["stk|P|gt"], "P"), (al[0], al[1], "P:alias")])
    for kk, v in used.items():
        used_tot[kk] += v
    add_item("З'єднувач компресійний", "з'єднувач компресійний", o["nr|30-1|ПланКол"], dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"]), d,
             {"vne_smety": False, "sostav_zero": False, "mshp": False, "alias_of": "Муфта ПЕ 25×3/4", "card": "000000009", "n": 2, "Z": "2/5"})
    alias_used["муфта"] = used.get("P:alias", Z)

    G_tot = sum(G.values())
    # корректировки кэш → свёртка (по позициям)
    def cache_pos(k):
        o = ON[k]
        nq, ns = o["nr|30-1|ПланКол"], o["nr|30-1|ПланГрн"]
        pss = dv(ns, nq)
        need = max(Z, nq - o["c|30-3|q"]); f_own = min(need, max(Z, o["stkh|30-3|q"])); need -= f_own
        f_P = min(need, max(Z, o["stk|P|q"])); tot = (need - f_P) * pss; rem = max(Z, o["stk|P|q"] - f_P)
        for _ in range(4):
            take = min(nq, rem); rem -= take; tot += (nq - take) * pss
        return tot
    for k, lab in (("гайка", "Гайка: неттинг по розмірах (М48/М24/М12), а не змішаною групою"),
                   ("шпилька", "Шпилька М48: склад лише явно М48 (без «Шпилька 10х1000»)"),
                   ("комплект септика", "Комплект септика: склад — комплектами за складом з факту, деталі понад — у надлишки"),
                   ("кріплення ринви", "кріплення ринви ← Держак ринви пласт. зі складу проєкту (аліас)"),
                   ("з'єднувач компресійний", "З'єднувач компресійний ← Муфта ПЕ 25×3/4 (аліас)")):
        new = sum(Dd(it["sum"]) for it in items if it["key"] == k)
        adjust.append({"what": lab, "cache": r2(cache_pos(k)), "v3": r2(new), "delta": r2(new - cache_pos(k))})
    adj_sum = sum(Dd(a["delta"]) for a in adjust)
    assert abs(G_cache + adj_sum - G_tot) < Decimal("0.02"), (G_cache, adj_sum, G_tot)
    print(f"Дефицит: кэш {F(G_cache)} → после свёртки алиасов и состава {F(G_tot)} (Δ {F(adj_sum)}); по домам "
          + "; ".join(f"№{h[3:]} {F(G[h])}" for h in HOUSES))

    # ════════════════════════════════════ МШП_h (по факту выдачи на 30 м №1, №2; к дому — по цене закупки)
    # Одна потребность — в одном члене (промт §4.5): в МШП_h только группа «мшп» без нормы в карточке 000000009. Выдачи со сч. 22
    # позиций нормы (дріт, електроди, обрізні диски) и алиасы нормы в группе «мшп» (ANALIZ стр. 423: клей монтажний Lacrysil →
    # «клей», круг відрізний 230 мм → «Обрізні диски», 7 з 16 шт.) уже сидят в дефиците G_h — сюда не идут.
    G_KEYS = {it["key"] for it in items}
    NORM_KEYS = {k for k, o in ON.items() if o["nr|30-1|ПланКол"] > 0 or o["nr|30-1|ПланГрн"] > 0}
    MSHP_ALIAS = ((lambda n: "клей" in K(n) and "монтаж" in K(n), "клей"),
                  (lambda n: ("круг відрізний" in K(n) or "відрізний диск" in K(n) or "диск відрізний" in K(n)) and "230" in K(n), "обрізні диски"))

    def mshp_fold(name, on):
        """Свёрнутый ключ потребности для выдачи со сч. 22 / группы «мшп»: алиас нормы → её ключ; иначе общее название."""
        return next((t for f, t in MSHP_ALIAS if f(name)), K(on) or ("ном|" + K(name)))
    mshp_h = {}
    mshp_items = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z, Z]))   # [кол, брутто, нетто]
    mshp_excl = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z, Z]))
    mshp_why = {}
    for r in MV:
        if is_mat(r["СчКт"]) and not is_mat(r["СчДт"]):
            h = house(r["СклКт"])
            if h in ("30-1", "30-2", "30-3") and ((r["СчКт"] or "").startswith("22") or K(r["ОНКт"]) == "мшп"):
                g = Dd(r["Сумма"]) * (1 + VAT.ratio(r["ПартКт"], r["КодКт"], r["СклКт"]))
                fk = mshp_fold(r["НомКт"], r["ОНКт"])
                if K(r["ОНКт"]) == "мшп" and fk == "мшп":
                    tgt = mshp_items[r["НомКт"]]
                else:
                    tgt = mshp_excl[r["НомКт"]]
                    mshp_why[r["НомКт"]] = (fk, ("аліас норми «" + fk + "» — у G_h") if fk in NORM_KEYS and K(r["ОНКт"]) == "мшп" else
                                            ("позиція норми СС «" + fk + "» (видача з рах. 22) — у G_h") if fk in NORM_KEYS else
                                            "рах. 22 без норми, не група «мшп» — у ВС_h (чутливість)")
                tgt[h][0] += Dd(r["КолКт"] or 0); tgt[h][1] += g; tgt[h][2] += Dd(r["Сумма"])
    for h in ("30-1", "30-2", "30-3"):
        mshp_h[h] = sum(v[h][1] for v in mshp_items.values())
    mshp_rate = min(mshp_h["30-1"], mshp_h["30-2"])
    MSHP = {h: (max(Z, mshp_rate - mshp_h["30-3"]) if h == "30-3" else mshp_rate) for h in HOUSES}
    mshp_fk = {"мшп"} if mshp_items else set()          # свёрнутые ключи потребностей, попавших в МШП_h
    mshp_inter_G = sorted(mshp_fk & G_KEYS)
    mshp_excl_tab = [{"name": n, "key": mshp_why[n][0], "why": mshp_why[n][1],
                      "net": {h: r2(v[h][2]) for h in ("30-1", "30-2", "30-3")}, "gross": {h: r2(v[h][1]) for h in ("30-1", "30-2", "30-3")},
                      "in_G": mshp_why[n][0] in G_KEYS} for n, v in sorted(mshp_excl.items(), key=lambda x: -sum(y[1] for y in x[1].values()))]
    print("МШП_h: вилучено (норма СС / аліаси норми — вже в G_h): "
          + "; ".join(f"{e['name'][:34]} → {e['key']} (№1 {F(e['net']['30-1'])}, №2 {F(e['net']['30-2'])}, №3 {F(e['net']['30-3'])} нетто)"
                      for e in mshp_excl_tab))
    print(f"МШП_h ∩ G_h після згортання аліасів: {mshp_inter_G or '∅'}")
    # кто платит МШП домов IRC: Казна 065 IRC по контрагентам и перемещения ІБ_МШП → склады домов IRC
    pay065 = collections.defaultdict(Decimal)
    for r in MO:
        if r["Ст"] == "УТ-001065" and r["dt"] <= T_MONEY:
            pay065[r["Контрагент"] or r["Подотчет"] or "(без контрагента)"] += -Dd(r["Сумма"])
    msp_to_irc = sum(Dd(r["Сумма"]) for r in MV if is_mat(r["СчДт"]) and is_mat(r["СчКт"]) and wcls(r["СклКт"]) == "MSP"
                     and (wcls(r["СклДт"]).startswith("H:") or wcls(r["СклДт"]) == "P"))
    pu_msp = sum(Dd(r["БезНДС"]) for r in PU if "МШП" in r["Склад"])
    pu_msp_g = sum(Dd(r["Стоимость"]) for r in PU if "МШП" in r["Склад"])
    mshp_conf = p065 > 0 and pu_msp > 0
    # закуплено IRC на ІБ_МШП и ещё не передано на склады IRC (не больше остатка ІБ_МШП) — риск двойной оплаты против МШП_h
    msp_bal = sum(Dd(s) for w, s in LB["bal"].items() if wcls(w) == "MSP")
    irc_msp_left = max(Z, min(pu_msp - msp_to_irc, msp_bal))
    irc_msp_left_g = irc_msp_left * dv(pu_msp_g, pu_msp)
    print(f"МШП: 30 м №1 {F(mshp_h['30-1'])}, №2 {F(mshp_h['30-2'])}, №3 уже {F(mshp_h['30-3'])} → на дом {F(mshp_rate)}; "
          f"Σ №3–№7 {F(sum(MSHP.values()))}; IRC платил 065 {F(p065)}, закупил МШП на ІБ_МШП {F(pu_msp)} нетто, "
          f"ІБ_МШП → склади IRC {F(msp_to_irc)} нетто, залишок закупленого IRC на ІБ_МШП ≈ {F(irc_msp_left)} нетто / {F(irc_msp_left_g)} брутто "
          f"(залишок ІБ_МШП {F(msp_bal)}) → {'IRC платить' if mshp_conf else 'НЕ підтверджено'}")

    # ════════════════════════════════════ ПОСТАВЩИКИ, «Долги Коррект», КЗ и авансы
    DOLGI = load("dolgi.json")
    sup = collections.defaultdict(lambda: {"names": set(), "edrpou": set(), "pu_g": Z, "pu_n": Z, "pu_g_after": Z, "pu_msp_g": Z,
                                           "p067": Z, "p065": Z, "p_after": Z, "refund": Z, "pays": [], "purch": []})
    for r in PU:
        s_ = sup[nk(r["Контрагент"])]
        s_["names"].add(r["Контрагент"]); s_["edrpou"].add(r["ЕДРПОУ"])
        s_["pu_g"] += Dd(r["Стоимость"]); s_["pu_n"] += Dd(r["БезНДС"])
        if r["date"] > "2026-09-30":
            s_["pu_g_after"] += Dd(r["Стоимость"])
        if "МШП" in r["Склад"] or K(r["ОН"]) == "мшп":
            s_["pu_msp_g"] += Dd(r["Стоимость"])
        s_["purch"].append((r["date"], r["ОН"], r["Ном"][:40], r["Кол"], r["Стоимость"]))
    for r in L["money"]:
        who = r["Контрагент"] or r["Подотчет"] or "(без контрагента)"
        s_ = sup[nk(who)]
        s_["names"].add(who)
        if r["ЕДРПОУ"]:
            s_["edrpou"].add(r["ЕДРПОУ"])
        if r["Ст"] in ("УТ-001147", "УТ-001204"):
            s_["refund"] += Dd(r["Сумма"])
            continue
        if r["dt"] <= T_MONEY:
            s_["p067" if r["Ст"] == "УТ-001067" else "p065"] += -Dd(r["Сумма"])
            s_["pays"].append((r["dt"][:10], r["Ст"], -r["Сумма"], r["Назн"][:160]))
        else:
            s_["p_after"] += -Dd(r["Сумма"])
    # BuhBud: поступления по поставщику (нетто), доп. расходы BuhBud — для I14 (б) и классов расхождений
    rcdoc = {r["Док"]: r["Контрагент"] for r in LB["receipts"]}
    buh_in = collections.defaultdict(lambda: collections.defaultdict(Decimal))
    for r in MV:
        if is_mat(r["СчДт"]) and not is_mat(r["СчКт"]):
            p = r["ПартДт"]
            kind = "АО" if p.startswith("Авансовий") else ("доп" if "додатков" in r["Рег"].lower() else "ПТУ")
            buh_in[nk(rcdoc.get(p, "(" + p.split(" ")[0] + ")"))][kind] += Dd(r["Сумма"])
    # книга «Долги Коррект»: IRC, материалы и МШП
    kn_rows = [x for x in DOLGI["kniga"] if " ".join(x["cfo"].replace("\xa0", " ").split()) == "МД IRC 2026"
               and x["article"] in ("Строительные материалы", "МШП")]
    keys = list(sup.keys())

    def match(kontr):
        tb = toks(kontr)
        best = [k for k in keys if tb and (tb <= toks(" ".join(sup[k]["names"])) or toks(" ".join(sup[k]["names"])) <= tb)
                and toks(" ".join(sup[k]["names"]))]
        return best[0] if len(best) == 1 else (best and sorted(best, key=lambda k: -len(toks(" ".join(sup[k]["names"]) ) & tb))[0])
    erp_tp = collections.defaultdict(lambda: [0.0, 0.0])
    for x in DOLGI["erp"]:
        if x["projname"] == "МД IRC 2026" and x["kind"] == "поставщики":
            erp_tp[x["partner"]][0] += x["debt"]; erp_tp[x["partner"]][1] += x["advance"]
    book = collections.defaultdict(list)
    unmatched_book = []
    for x in kn_rows:
        if x["debt"] <= 0:
            continue
        m = match(x["kontr"])
        if not m:
            m = nk(x["kontr"])
            sup[m]["names"].add(x["kontr"])
            unmatched_book.append(x["row"])
        book[m].append(x)
    # авансовые отчёты (ERP «000Ц-…» = закупка через подотчётное лицо): по дате отчёта — сумма всех закупок дня; возмещение —
    # выдача Казны IRC с текстом «…авансов…» (по 30.09), не раньше отчёта, сумма до 0,01. Невозмещённое — долг работнику, остаётся в КЗ.
    ao_grp = collections.defaultdict(Decimal)
    for r in PU:
        if r["Номер"].startswith("000Ц"):
            ao_grp[r["date"]] += Dd(r["Стоимость"])
    reimb = sorted((r["dt"], -Dd(r["Сумма"]), r["Док"]) for r in MO
                   if r["dt"] <= T_MONEY and not r["Контрагент"] and "відшкод" in K(r["Назн"]).replace("i", "і")
                   and "авансов" in K(r["Назн"]).replace("i", "і"))
    used_rb, AO_GROUPS = set(), []
    for d_, s in sorted(ao_grp.items()):
        hit = next((i for i, (rd, rs, _) in enumerate(reimb) if i not in used_rb and rd[:10] >= d_ and abs(rs - s) <= Decimal("0.01")), None)
        if hit is not None:
            used_rb.add(hit)
        AO_GROUPS.append({"date": d_, "sum": r2(s), "reimbursed": None if hit is None else
                          {"date": reimb[hit][0][:10], "sum": r2(reimb[hit][1]), "doc": reimb[hit][2]}})
    ao_open = {g["date"] for g in AO_GROUPS if not g["reimbursed"]}
    SUP = []
    for k_, s_ in sup.items():
        paid = s_["p067"] + s_["p065"]
        bal_ = s_["pu_g"] - paid
        erp_kz, erp_av = max(Z, bal_), max(Z, -bal_)
        b067 = sum(Dd(x["debt"]) for x in book[k_] if x["article"] == "Строительные материалы")
        b065 = sum(Dd(x["debt"]) for x in book[k_] if x["article"] == "МШП")
        bt = b067 + b065
        if erp_kz == 0 and erp_av == 0 and bt == 0:
            continue
        name = sorted(s_["names"], key=len)[-1]
        nm = nk(name)
        nontrade = any(t in nm for t in NONTRADE) or s_["pu_g"] == 0 and any(t in nm for t in NONTRADE)
        bi = buh_in.get(k_, {})
        dop = bi.get("доп", Z)
        dlt = bt - erp_kz
        tp = next((v for p_, v in erp_tp.items() if toks(p_) and (toks(p_) <= toks(name) or toks(name) <= toks(p_))), None)
        # класс расхождения (X1 §3.5) и решение (X1 §3.4: КЗ — большая сторона, АВ — меньшая);
        # третий источник — ERP РасчетыСПоставщикамиПоСрокам на 01.10 по партнёру проекта (dolgi.json → erp), только для I2
        tpd = Dd(tp[0]) if tp else None
        svc = max(Z, bi.get("ПТУ", Z) - sum(Dd(r["БезНДС"]) for r in PU if nk(r["Контрагент"]) == k_ and not r["Номер"].startswith("000Ц")))
        ao_only = s_["pu_g"] > 0 and all(r["Номер"].startswith("000Ц") for r in PU if nk(r["Контрагент"]) == k_) and paid == 0
        upside, resolved = Z, Z
        ao_unr = sum(Dd(r["Стоимость"]) for r in PU if nk(r["Контрагент"]) == k_ and r["Номер"].startswith("000Ц") and r["date"] in ao_open)
        erp_kz_used = min(erp_kz, ao_unr) if (ao_only and bt == 0) else erp_kz
        if ao_only and bt == 0 and ao_unr > 0:
            cls, why = "A", ("закупівля за авансовим звітом працівника від " + ", ".join(sorted({r["date"] for r in PU if nk(r["Контрагент"]) == k_
                                                                                               and r["date"] in ao_open}))
                             + " (ERP «000Ц-…»): відшкодування в Казні IRC (067/065, сума ± 0,01, текст «відшкод… авансов…») не знайдено; "
                               "ручна перевірка 03.10: серед виплат підзвітній особі в усіх підрозділах 06–10.2026 такої суми немає, "
                               "переказ 15.07 10 260,00 — «Перечисление ДС на другой счет» (не аванс) → борг працівнику, у КЗ (більша сторона)")
        elif ao_only and bt == 0:
            cls, why = "A", ("закупівля за авансовим звітом працівника (ERP «000Ц-…» = BuhBud «Авансовий звіт»): оплачено працівником, "
                             "відшкодовано Казною («Вiдшкодування авансових звiтiв (ІРС тмц)» на ту саму суму) — постачальнику боргу немає")
        elif abs(dlt) <= 1:
            cls, why = "R", "книга = ERP"
        elif dlt > 0:
            third = tpd if tpd is not None else Z
            resolved = max(Z, min(bt, third) - erp_kz)
            upside = max(Z, bt - max(erp_kz, third))
            conf = f"; ERP-розрахунки (РасчетыСПоставщикамиПоСрокам) — борг {F(third)}" if tpd is not None else "; у ERP-розрахунках партнера немає"
            if upside <= 1 and erp_av > 0:
                cls, why = "C", (f"книга — борг {F(bt)}, у закупівлях ERP — «аванс» {F(erp_av)}{conf}: борг підтверджено третім джерелом, "
                                 f"«аванс» — артефакт закупівель без послуг/рядків поза СебестоимостьТоваров")
            elif upside <= 1:
                cls, why = "M", (f"книга {F(bt)} > закупівлі ERP {F(erp_kz)}{conf}: різниця — послуги/доставка в накладних "
                                 f"(у BuhBud у собівартості {F(svc)} нетто), яких немає в СебестоимостьТоваров — підтверджено")
            elif erp_av > 0:
                cls, why = "C", f"книга — борг {F(bt)}, у закупівлях ERP — аванс {F(erp_av)}{conf}; не розв'язано {F(upside)} → більша сторона"
            elif s_["refund"] > 0:
                cls, why = "U", (f"повернення {F(s_['refund'])} надійшло на статтю повернень (гроші Казни) — розв'язано {F(resolved)}{conf}; "
                                 f"решта {F(upside)} (рядок книги без накладної ERP) → більша сторона")
            else:
                cls, why = "X", f"книга {F(bt)} > ERP {F(erp_kz)}{conf}; причину не знайдено, {F(upside)} → більша сторона"
        else:
            conf = f"; ERP-розрахунки — борг {F(tpd)}" if tpd is not None else ""
            cls, why = ("M", f"у книзі рядка немає, накладна ERP є{conf}") if bt == 0 else ("X", f"книга {F(bt)} < ERP {F(erp_kz)}{conf} → ERP")
        kz = max(erp_kz_used, bt)
        av = erp_av if bt == 0 else Z
        # аванс: сверка с ERP-расчётами (третий источник), класс X1 §3.5 — для I2 (в базе зачёт только по правилу X1 §3.4)
        av_cls, av_why = "", ""
        if av > 0:
            tpa = Dd(tp[1]) - Dd(tp[0]) if tp else None
            if tpa is not None and abs(tpa - av) <= 1:
                av_cls, av_why = "R", "аванс = ERP-розрахунки"
            elif nontrade and "(без" in name:
                rb_ = sum(Dd(p[2]) for p in s_["pays"] if "авансов" in K(p[3]).replace("i", "і"))
                av_cls, av_why = "A", (f"відшкодування авансових звітів працівникам {F(rb_)} — товари у закупівлях ERP від АЗС/магазинів (клас A у КЗ); "
                                       f"решта {F(av - rb_)} — переказ на власний рахунок компанії («Перечисление ДС на другой счет», "
                                       f"логістика), нетоварний")
            elif nontrade:
                av_cls, av_why = "C", "послуги (пошта/транспорт) на статті матеріалів — не товарний аванс"
            elif svc > 0:
                av_cls, av_why = "M", (f"частина «авансу» — доставка/послуги в накладних BuhBud ({F(svc)} нетто), яких немає в "
                                       f"СебестоимостьТоваров; ERP-розрахунки: сальдо {F(tpa if tpa is not None else 0)} (аванс +)")
            else:
                av_cls, av_why = "A", (f"ERP-розрахунки: аванс {F(tpa if tpa is not None else 0)} — оплати віднесено до іншого об'єкта "
                                       f"розрахунків/договору (напр. рахунок 2025 р.)")
        # 067 / 065: ряды книги по их статье; остаток ERP сверх книги — по статье оплат Казны, при смешанной — по группе закупок
        rest = kz - bt
        if rest > 0:
            if s_["p065"] > 0 and s_["p067"] == 0:
                sh = Decimal(1)
            elif s_["p067"] > 0 and s_["p065"] == 0:
                sh = Z
            else:
                sh = dv(s_["pu_msp_g"], s_["pu_g"])
            k065, k067 = b065 + rest * sh, b067 + rest * (1 - sh)
        else:
            k065, k067 = b065, b067
        SUP.append({"key": k_, "name": name, "edrpou": sorted(e for e in s_["edrpou"] if e), "pu_g": r2(s_["pu_g"]), "pu_n": r2(s_["pu_n"]),
                    "pu_g_after30": r2(s_["pu_g_after"]), "pu_mshp_g": r2(s_["pu_msp_g"]), "paid067": r2(s_["p067"]), "paid065": r2(s_["p065"]),
                    "paid_after30": r2(s_["p_after"]), "refund": r2(s_["refund"]), "erp_kz": r2(erp_kz), "erp_av": r2(erp_av),
                    "book067": r2(b067), "book065": r2(b065), "book_rows": [x["row"] for x in book[k_]],
                    "book_rows065": [x["row"] for x in book[k_] if x["article"] == "МШП"], "erp_tp": tp,
                    "buh_in": {kk: r2(v) for kk, v in bi.items()}, "buh_services_net": r2(svc), "ao_only": bool(ao_only),
                    "delta": r2(dlt), "cls": cls, "why": why, "upside": r2(upside), "resolved_by_third": r2(resolved),
                    "av_cls": av_cls, "av_why": av_why,
                    "ao_unreimbursed": r2(erp_kz_used) if ao_only and bt == 0 else 0.0,
                    "decision": ("КЗ — авансовий звіт без відшкодування (борг працівнику)" if ao_only and bt == 0 and erp_kz_used > 0 else
                                 "0 — оплачено через авансовий звіт, відшкодовано Казною" if ao_only and bt == 0 else
                                 "книга (більша сторона)" if bt > erp_kz + 1 else ("ERP" if erp_kz > bt + 1 else "ERP = книга")),
                    "kz": r2(kz), "kz067": r2(k067), "kz065": r2(k065), "av": r2(av), "nontrade": bool(nontrade),
                    "pays": s_["pays"][-6:], "purch_last": sorted(s_["purch"])[-4:]})
    SUP.sort(key=lambda x: -(x["kz"] + x["av"]))
    UPS = sum(Dd(x["upside"]) for x in SUP)
    AO_KZ = sum(Dd(x["erp_kz"]) - Dd(x["kz"]) for x in SUP if x["ao_only"])          # возмещено Казной — не долг
    AO_UNR = sum(Dd(x["ao_unreimbursed"]) for x in SUP if x["ao_only"])                # отчёты без возмещения — долг работнику (в КЗ)
    KZ067 = sum(Dd(x["kz067"]) for x in SUP)
    KZ065 = sum(Dd(x["kz065"]) for x in SUP)
    ERP_KZ = sum(Dd(x["erp_kz"]) for x in SUP)
    ERP_AV = sum(Dd(x["erp_av"]) for x in SUP)
    print(f"КЗ: ERP брутто {F(ERP_KZ)}, авансы {F(ERP_AV)}; по большей стороне 067 {F(KZ067)} + 065 {F(KZ065)} = {F(KZ067 + KZ065)}")
    ctl(4, "Поставщик за поставщиком: долг / авансы / нетто = закуплено − оплачено", "1 668 344,64 / 213 538,64 / 1 454 806,00",
        f"{F(ERP_KZ)} / {F(ERP_AV)} / {F(ERP_KZ - ERP_AV)} (= {F(pu_g)} − {F(p067 + p065)} = {F(pu_g - p067 - p065)})",
        "закупки (п. 2) − оплаты 067+065 (п. 3) по контрагенту", abs(ERP_KZ - Decimal("1668344.64")) < Decimal("0.005")
        and abs(ERP_AV - Decimal("213538.64")) < Decimal("0.005"))
    # адреса рядов книги
    addr = []
    for x in kn_rows:
        if x["debt"] <= 0:
            addr.append({"row": x["row"], "kontr": x["kontr"], "article": x["article"], "debt": x["debt"], "to": "пропущен (долг 0, план)"})
            continue
        addr.append({"row": x["row"], "kontr": x["kontr"], "article": x["article"], "debt": x["debt"],
                     "to": "065" if x["article"] == "МШП" else "067"})
    b_mat = sum(Dd(x["debt"]) for x in kn_rows if x["article"] == "Строительные материалы")
    b_msp = sum(Dd(x["debt"]) for x in kn_rows if x["article"] == "МШП")
    msp_rows = [x for x in kn_rows if x["article"] == "МШП" and x["debt"] > 0]
    ctl(14, "«Долги Коррект» IRC: материалы (строк) / МШП (ряды)", "1 604 157,13 (42); МШП 32 420,14 — ряды 47, 57, 74, 79, 85",
        f"{F(b_mat)} ({sum(1 for x in kn_rows if x['article'] == 'Строительные материалы')}); МШП {F(b_msp)} — ряды "
        + ", ".join(str(x["row"]) for x in msp_rows), "data/dolgi.json → kniga",
        abs(b_mat - Decimal("1604157.13")) < Decimal("0.005") and abs(b_msp - Decimal("32420.14")) < Decimal("0.005"))

    # ---------- авансы: каждый аванс отдельно (X1 §3.4, K13 M04)
    AVL = []
    zachet = collections.defaultdict(Decimal)
    gq = {(it["key"], it["name"]): it for it in items}
    for x in SUP:
        if x["av"] <= 0:
            continue
        s_ = sup[x["key"]]
        pays = sorted(s_["pays"])
        last = pays[-1] if pays else None
        age = (datetime.date(2026, 9, 30) - datetime.date.fromisoformat(last[0])).days if last else None
        zak = sorted({r["Заказ"] for r in PU if nk(r["Контрагент"]) == x["key"] and r["Заказ"].strip()})
        rec = {"name": x["name"], "av": x["av"], "last_pay": last[0] if last else None, "last_pay_text": last[3] if last else "",
               "age_30_09": age, "term_src": ("заказ ERP: " + ", ".join(zak)) if zak else "заказу ERP немає; рахунок без строку поставки",
               "goods": "", "goods_list": [], "closes_deficit": False, "decision": "", "zachet": 0.0}
        if x["nontrade"] or not x["pu_g"] and "(без" in x["name"]:
            rec["decision"] = "0 — нетоварний"
        elif age is None or age >= 30:
            rec["decision"] = "0 — старий (≥ 30 днів), строк/статус поставки не підтверджено → ризики"
        else:
            # свежий: товар из назначения платежей последнего дня, не поступивший в ERP после оплаты
            day = [p for p in pays if p[0] == last[0]]
            got = {K(r["ОН"]) for r in PU if nk(r["Контрагент"]) == x["key"] and r["date"] >= last[0]}
            goods = []
            def cp(a, b):
                n = 0
                while n < min(len(a), len(b)) and a[n] == b[n]:
                    n += 1
                return n
            for p in day:
                words = [w for w in re.split(r"[^\w']+", K(p[3])) if len(w) >= 4]
                for it in items:
                    w0 = K(it["name"]).split(" ")[0]
                    if it["key"] not in got and any(cp(w, w0) >= max(4, min(len(w), len(w0)) - 3) for w in words):
                        goods.append((p, it))
            if goods:
                tot = Z
                for p, it in goods:
                    # количество — по цене прошлой поставки этого товара этим поставщиком (в назначении платежа его нет)
                    prev = sorted((r for r in PU if nk(r["Контрагент"]) == x["key"] and K(r["ОН"]) == it["key"] and r["Кол"] > 0),
                                  key=lambda r: r["date"])
                    unit = Dd(prev[-1]["Стоимость"]) / Dd(prev[-1]["Кол"]) if prev else Z
                    qty = Decimal(int(Dd(p[2]) / unit)) if unit else Z
                    left = qty
                    for h in HOUSES:
                        tk = min(left, Dd(it["def_q"][h]))
                        if tk > 0:
                            zachet[h] += tk * Dd(it["price_ss"]); tot += tk * Dd(it["price_ss"]); left -= tk
                    if qty <= 0:
                        continue
                    rec["goods"] += f"{it['name']} ~{int(qty)} шт (оплата {p[0]} {F(p[2])}, ціна попередньої поставки {F(unit)}); "
                    rec["goods_list"].append({"item": it["name"], "qty": int(qty), "date": p[0], "pay": float(p[2]), "unit": r2(unit)})
                rec["closes_deficit"] = tot > 0
                rec["zachet"] = r2(min(tot, Dd(x["av"])))
                rec["decision"] = f"зачёт {F(rec['zachet'])} — свіжий (< 30 днів), товар у позиціях дефіциту (за ціною СС, кількістю)"
            else:
                rec["decision"] = "0 — свіжий, але товар не в позиціях дефіциту (або вже надійшов) → ризики"
        AVL.append(rec)
    AV_Z = sum(zachet.values())
    print(f"Авансы: товарные {F(sum(Dd(a['av']) for a in AVL if 'нетовар' not in a['decision']))}, зачтено {F(AV_Z)}")

    # ════════════════════════════════════ реестр рисков (в базу не идёт; суммы со знаком: − upside, + downside)
    RISKS = []
    g_by_key = collections.defaultdict(Decimal)
    for it in items:
        g_by_key[it["key"]] += Dd(it["sum"])
    for x in SUP:
        # поставщик: в ERP аванс (оплачено > получено), в книге долг (взят большей стороной) и всё, что он поставлял, — позиции дефицита G_h:
        # оплачено + долг по книге «покрывают» больше штук, чем пришло, — КЗ и G_h могут платить за одни и те же штуки
        its = {K(r["ОН"]) for r in PU if nk(r["Контрагент"]) == x["key"]}
        if x["erp_av"] > 0 and x["book067"] + x["book065"] > 0 and x["kz"] > 0 and its and its <= set(g_by_key):
            q_in = sum(r["Кол"] for r in PU if nk(r["Контрагент"]) == x["key"])
            gi = sum(g_by_key[k] for k in its)
            inv = sorted({m.group(1) for p in sup[x["key"]]["pays"] for m in [re.search(r"Рахунок\s*№\s*(\S+)", p[3])] if m})
            RISKS.append({"code": "kz_and_G_same_goods", "low": -x["kz"], "high": -r2(gi),
                          "what": f"{x['name']}: КЗ за книгою {F(x['kz'])} (ряди {', '.join(str(r) for r in x['book_rows'])}) і дефіцит G_h "
                                  f"«{', '.join(sorted(its))}» {F(gi)} — ймовірно ті самі штуки: оплачено {F(x['paid067'] + x['paid065'])} "
                                  f"(рахунки № {', '.join(inv)}), отримано в ERP {q_in:g} шт на {F(x['pu_g'])}, аванс ERP {F(x['erp_av'])}; "
                                  f"оплачено + борг за книгою ≈ {F(Dd(x['paid067']) + Dd(x['paid065']) + Dd(x['kz']))}"
                                  + (f"; перекривається з upside КЗ «книга > ERP» {F(x['upside'])} по цьому постачальнику — вилкою, не доданком"
                                     if x["upside"] > 0 else ""),
                          "owner": "постачання / економіст: скільки штук за рахунками виставлено й поставлено",
                          "if_confirmed": "поставку підтверджено, але не проведено — зняти КЗ або відповідні штуки з G_h"})
    if irc_msp_left_g > 0 and sum(MSHP.values()) > 0:
        RISKS.append({"code": "irc_mshp_on_ib_mshp", "low": 0.0, "high": -r2(min(sum(MSHP.values()), irc_msp_left_g)),
                      "what": f"МШП, закуплені IRC на ІБ_МШП ({F(pu_msp)} нетто) і ще не передані на склади IRC ({F(msp_to_irc)} нетто передано): "
                              f"залишок ≈ {F(irc_msp_left)} нетто / {F(irc_msp_left_g)} брутто — якщо склад видасть їх будинкам IRC, МШП_h "
                              f"не потребує нових грошей (не більше Σ МШП_h {F(sum(MSHP.values()))}); вилкою з «МШП_h платить ЦО»",
                      "owner": "кладовщик ІБ_МШП / постачання"})
    if AO_UNR > 0:
        RISKS.append({"code": "ao_unreimbursed", "low": 0.0, "high": -r2(AO_UNR),
                      "what": f"авансові звіти IRC без відшкодування в Казні ({', '.join(g['date'] + ' ' + F(g['sum']) for g in AO_GROUPS if not g['reimbursed'])}) "
                              f"— у КЗ; якщо відшкодовано поза Казною IRC (готівка, інший підрозділ), борг 0",
                      "owner": "бухгалтер / підзвітна особа"})
    for rk in RISKS:
        print(f"Ризик: {rk['what']} → {F(rk['low'])} … {F(rk['high'])}")

    # ════════════════════════════════════ чувствительность (в строки не идёт)
    sens = {}
    # ВС: позиции без нормы 009, выданные на оба построенных 30 м (кроме МШП и алиасов), per house = min(№1, №2)
    ALIAS_KEYS = {"держак ринви", "кутики (металопрокат)", "муфта", "мшп", "дозатор рідкого мила", "дозатор рушників"}
    vs = collections.defaultdict(Decimal)
    vs_items, vs_keys = [], set()
    for k, o in ON.items():
        if o["nr|30-1|ПланГрн"] > 0 or o["nr|30-1|ПланКол"] > 0 or k in ALIAS_KEYS:
            continue
        q1, q2 = o["c|30-1|q"], o["c|30-2|q"]
        if q1 <= 0 or q2 <= 0:
            continue
        rate = min(q1, q2)
        price = dv(o["c|30-1|gt"] + o["c|30-2|gt"], q1 + q2)
        d, used, pools = net(rate, price, o["c|30-3|q"], (o["stkh|30-3|q"], o["stkh|30-3|gt"]), [(o["stk|P|q"], o["stk|P|gt"], "P")])
        s = sum(v * price for v in d.values())
        for h, v in d.items():
            vs[h] += v * price
        if s > 0:
            vs_items.append((NAME[k], r2(s))); vs_keys.add(k)
    inter = {it["key"] for it in items} & vs_keys
    sens["vs"] = {"sum": r2(sum(vs.values())), "by_house": {h: r2(v) for h, v in vs.items()}, "items": sorted(vs_items, key=lambda x: -x[1]),
                  "intersect_G": sorted(inter), "intersect_MSHP": sorted(vs_keys & mshp_fk),
                  "note": "у v2 вже є УТ-001069 «Непередбачувані матеріали» 2 000 × 5 = 10 000"}
    # СК: позиции 009, не списанные ни на одном построенном 30 м (после свёртки алиасов), их дефицит в G
    sk_all = [it for it in items if it["sostav_zero"]]
    NOTBOUGHT = ("бойлер 82", "інформаційний банер", "підвіконик пвх", "труба гофра", "сигналізація виявлення диму", "підключення емкості")
    sk_nb = [it for it in sk_all if it["key"] in NOTBOUGHT]
    sens["sk"] = {"all": r2(sum(Dd(i["sum"]) for i in sk_all)), "not_bought": r2(sum(Dd(i["sum"]) for i in sk_nb)),
                  "items": [(i["name"], i["sum"]) for i in sorted(sk_all, key=lambda i: -i["sum"])],
                  "Z": "2/5", "z_weighted_not_bought": r2(-Decimal(2) / 5 * sum(Dd(i["sum"]) for i in sk_nb)),
                  "z_weighted_all": r2(-Decimal(2) / 5 * sum(Dd(i["sum"]) for i in sk_all))}
    # вариант 2 «норматив цеха»: сопоставимые позиции — расход min(№1, №2) вместо нормы; + ВС 100 %
    v2d = Z
    for k, o in ON.items():
        if k in SPECIAL or cat30.get(k) != "сопоставимо" or o["nr|30-1|ПланКол"] <= 0:
            continue
        nq, pss = o["nr|30-1|ПланКол"], dv(o["nr|30-1|ПланГрн"], o["nr|30-1|ПланКол"])
        if o["c|30-1|q"] <= 0 or o["c|30-2|q"] <= 0:
            continue            # «частичный ноль» — не экономия объёма (ANALIZ §7)
        rate = (o["c|30-1|q"] + o["c|30-2|q"]) / 2
        d0, _, _ = net(nq, pss, o["c|30-3|q"], (o["stkh|30-3|q"], o["stkh|30-3|gt"]), [(o["stk|P|q"], o["stk|P|gt"], "P")])
        d1, _, _ = net(rate, pss, o["c|30-3|q"], (o["stkh|30-3|q"], o["stkh|30-3|gt"]), [(o["stk|P|q"], o["stk|P|gt"], "P")])
        v2d += sum(d1.values()) * pss - sum(d0.values()) * pss
    sens["var2"] = {"volume": r2(v2d), "vs": sens["vs"]["sum"], "total": r2(v2d + Dd(sens["vs"]["sum"]))}
    # цена замещения (M05 §3.6): склад проекта → расход 30 м → расход 15 м → СС; вне [0,5; 2] × СС — СС
    rep = Z
    for it in items:
        o = ON.get(it["key"])
        if o is None:
            continue
        ps = Dd(it["price_ss"])
        cand = [dv(o["stk|P|gt"], o["stk|P|q"]), dv(o["c|30-1|gt"] + o["c|30-2|gt"], o["c|30-1|q"] + o["c|30-2|q"]),
                dv(sum(o[f"c|15-{i}|gt"] for i in range(1, 7)), sum(o[f"c|15-{i}|q"] for i in range(1, 7)))]
        p = next((c for c in cand if c > 0), ps)
        if ps > 0 and not (Decimal("0.5") <= p / ps <= 2) or it["key"] in ("шпилька", "гайка", "шайба"):
            p = ps
        rep += sum(Dd(q) for q in it["def_q"].values()) * (p - ps)
    sens["rep_price"] = r2(rep)
    # дисконт A-позиций склада проекта (инвентаризации нет): недостача 5 % / 10 % идёт в дефицит
    pv = sorted(((k, o["stk|P|gt"]) for k, o in ON.items() if o["stk|P|gt"] > 0), key=lambda x: -x[1])
    acc, A = Z, []
    for k, v in pv:
        if acc < P_gt * Decimal("0.8"):
            A.append(k); acc += v
    under_norm_A = Z
    for k in A:
        o = ON[k]
        nq = o["nr|30-1|ПланКол"]
        if nq > 0:
            needP = max(Z, nq - o["c|30-3|q"] - max(Z, o["stkh|30-3|q"])) + nq * 4
            under_norm_A += min(needP, max(Z, o["stk|P|q"])) * dv(o["stk|P|gt"], o["stk|P|q"])
    sens["discount"] = {"A_positions": len(A), "A_value": r2(acc), "A_under_norm": r2(under_norm_A),
                        "5%": r2(under_norm_A * Decimal("0.05")), "10%": r2(under_norm_A * Decimal("0.10"))}

    # ════════════════════════════════════ исключено из используемого остатка (M05 §3.3, условия 1, 4, 6)
    excl = []
    for (k_, n), v in nom.items():
        q, s = v["P"]
        if q < 0 or (q == 0 and abs(s) > Decimal("0.005")):
            excl.append({"cond": 1, "name": n, "q": float(q), "net": r2(s), "why": "від'ємна кількість або сума при кількості 0"})
    no_norm_g = sum(o["stk|P|gt"] for k, o in ON.items() if o["nr|30-1|ПланКол"] <= 0 and o["nr|30-1|ПланГрн"] <= 0)
    alias_g = sum(alias_used.values())
    first_in = min((r["Период"] for r in MV if is_mat(r["СчДт"]) and wcls(r["СклДт"]) == "P"), default="")
    stock_left = P_gt - used_tot["P"] - used_tot["P:alias"]
    excl_sum = {"1": r2(sum(Dd(e["net"]) for e in excl)),
                "4_no_norm_30m": r2(no_norm_g - alias_g),
                "4_over_norm_left": r2(stock_left - (no_norm_g - alias_g)),
                "6_illiquid": 0.0, "6_note": f"перший прихід на склад проєкту {first_in[:10]} — < 180 днів, неліквідів немає"}

    # ════════════════════════════════════ I13 / I14 / I6 (факты)
    unres = {k: v[0] for k, v in VAT.unresolved.items()}
    k_cons = dv(cons8_gt + c3_gt, cons8_n + cons_n["30-3"])
    k_stock = dv(P_gt, P_net)
    erp_P = sum(Dd(v[0]) for n, v in L["erp_stock"].items() if wcls(n) == "P")
    i14v = float((erp_P - P_gt) / P_gt * 100)
    # I14 (б): приход ERP = поступления BuhBud по поставщику (нетто); АО и доп. расходы — отдельно
    ao_erp = sum(Dd(r["БезНДС"]) for r in PU if r["Номер"].startswith("000Ц"))
    i14b = []
    for k_ in set(nk(r["Контрагент"]) for r in PU) | set(buh_in):
        e_ = sum(Dd(r["БезНДС"]) for r in PU if nk(r["Контрагент"]) == k_ and not r["Номер"].startswith("000Ц"))
        b_ = buh_in.get(k_, {}).get("ПТУ", Z)
        if abs(e_ - b_) > Decimal("0.01"):
            i14b.append({"supplier": k_, "erp_net": r2(e_), "buh_ptu_net": r2(b_), "delta": r2(e_ - b_)})
    ao_buh = sum(v.get("АО", Z) for v in buh_in.values())
    dop_buh = sum(v.get("доп", Z) for v in buh_in.values())
    # S_in / S_out (f32: измеренные перемещения, нетто → брутто по НДС партий)
    s_in = s_out = s_in_g = s_out_g = Z
    for r in MV:
        if is_mat(r["СчДт"]) and is_mat(r["СчКт"]):
            a, b = wcls(r["СклКт"]), wcls(r["СклДт"])
            g = Dd(r["Сумма"]) * (1 + VAT.ratio(r["ПартКт"], r["КодКт"], r["СклКт"]))
            if a in ("X", "MSP") and (b == "P" or b.startswith("H:")):
                s_in += Dd(r["Сумма"]); s_in_g += g
            if a == "P" and b in ("X", "MSP"):
                s_out += Dd(r["Сумма"]); s_out_g += g
    fb_n = fb_g = Z
    for r in MV:
        if is_mat(r["СчКт"]) and not is_mat(r["СчДт"]) and wcls(r["СклКт"]).startswith("H:") \
                and VAT.taint.get((r["ПартКт"], r["КодКт"], r["СклКт"])):
            fb_n += Dd(r["Сумма"]); fb_g += Dd(r["Сумма"]) * (1 + VAT.ratio(r["ПартКт"], r["КодКт"], r["СклКт"]))
    fbs_n = fbs_g = Z
    for r in MV:
        for side, sign in (("Дт", 1), ("Кт", -1)):
            acc = r["СчДт"] if side == "Дт" else r["СчКт"]
            if is_mat(acc) and wcls(r["Скл" + side]) == "P" and VAT.taint.get((r["Парт" + side], r["Код" + side], r["Скл" + side])):
                fbs_n += sign * Dd(r["Сумма"]); fbs_g += sign * Dd(r["Сумма"]) * (1 + VAT.ratio(r["Парт" + side], r["Код" + side], r["Скл" + side]))
    for b in i14b:
        b["why"] = ("послуги/доставка в накладній BuhBud (у собівартості), у СебестоимостьТоваров ERP їх немає" if b["delta"] < 0
                    else "розбіжність не пояснено")
    i14_unexpl = sum(abs(Dd(b["delta"])) for b in i14b if b["delta"] > 0)
    inv = {"I13": {"k_consumed": float(k_cons), "k_stock": float(k_stock), "unresolved_parts": unres,
                   "fallback_consumed_net": r2(fb_n), "fallback_consumed_gross": r2(fb_g),
                   "fallback_stockP_net": r2(fbs_n), "fallback_stockP_gross": r2(fbs_g),
                   "share_consumed_pct": float(dv(fb_g, cons8_gt + c3_gt) * 100), "share_stock_pct": float(dv(fbs_g, P_gt) * 100)},
           "I14": {"srez_tovar": f"момент снимка {stamp}, ERP и BuhBud — все документы (ДАТАВРЕМЯ(2026,12,31,23,59,59))",
                   "srez_money": T_MONEY, "erp_P_gross": r2(erp_P), "buh_P_gross": r2(P_gt), "dev_pct": i14v,
                   "baskets": i14b, "ao_erp": r2(ao_erp), "ao_buh": r2(ao_buh), "dop_buh": r2(dop_buh),
                   "services_buh": r2(-sum(Dd(b["delta"]) for b in i14b if b["delta"] < 0)), "unexplained": r2(i14_unexpl),
                   "erp_net": r2(pu_n), "buh_in_net": r2(sum(sum(v.values()) for v in buh_in.values()))},
           "I6": {"paid": r2(p067 + p065), "consumed_g": r2(cons8_gt + c3_gt), "s_book_g": r2(P_gt + built_gt + sum(o["stkh|30-3|gt"] for o in ON.values())),
                  "prepaid": r2(ERP_AV), "kz": r2(ERP_KZ), "s_in_n": r2(s_in), "s_out_n": r2(s_out), "s_in_g": r2(s_in_g), "s_out_g": r2(s_out_g)}}
    ctl(11, "I14: ERP СебестоимостьТоваров склад проекта (брутто) против BuhBud брутто; корзины по поставщикам",
        "ERP на 01.10 — 3 621 133,98 (−0,28 %); корзины: 7 634 660,50 = ПТУ 7 543 375,28 + авансовые 10 562,32 + МШП/чужие 80 722,90; доп. 20 169,17",
        f"ERP {F(erp_P)} против {F(P_gt)} ({i14v:+.2f} %); АО ERP {F(ao_erp)} / BuhBud {F(ao_buh)}; доп. расходы BuhBud {F(dop_buh)}; "
        f"поставщиков с |Δ| > 0,01 (без АО): {len(i14b)} — послуги в накладних BuhBud {F(-sum(Dd(b['delta']) for b in i14b if b['delta'] < 0))}, "
        f"не пояснено {F(i14_unexpl)}", "ERP РН.СебестоимостьТоваров.Остатки; BuhBud приходы f33 + движения",
        abs(i14v) <= 0.5 and i14_unexpl <= Decimal("0.01") and abs(ao_erp - ao_buh) <= Decimal("0.01"))
    # 12, 13 — вне сметы / не списано (только чувствительность)
    ctl(12, "Вне сметы на 30 м №1–№2 → дефицит к 30 м №3–№7 (без МШП, без алиасов)", "77 600,93 (окт 46 351,75 / ноя 31 249,19); МШП 32 141,06",
        f"{F(sens['vs']['sum'])}; МШП {F(sum(MSHP.values()))}", "кэш core.json + живые движения", True,
        f"в базе ВС = 0 (n = 2 < 3); МШП_h — лишь группа «мшп» без норм СС и их алиасов (выдачи со сч. 22 позиций нормы — "
        f"{', '.join(sorted({e['key'] for e in mshp_excl_tab}))} — уже в G_h), поэтому меньше ориентира исследования")
    ctl(13, "Позиции СС 009 без списания на 30 м №1–№2 в дефиците", "113 030,00 (не закуплено и не списано)",
        f"не закуплено {F(sens['sk']['not_bought'])}; все «не використано» {F(sens['sk']['all'])}", "кэш analysis.json d30_cat", True,
        "в базе СК = 0 (n = 2 < 3)")
    controls.sort(key=lambda c: c["n"])

    out = {"stamp": stamp,
           "srez": {"tovar": f"момент снимка {stamp}: все документы ERP и BuhBud (периоды по ДАТАВРЕМЯ(2026,12,31,23,59,59))",
                    "money": T_MONEY, "kazna_mirror_last": L["mirror_last"], "kazna_status": KZN["status"],
                    "kazna_late_30_09": KZN.get("late"), "kazna_windows": KZN.get("windows")},
           "cache_ok": cache_ok,
           "controls": controls,
           "suppliers": SUP,
           "kniga": {"rows": addr, "unmatched_rows": unmatched_book},
           "kz": {"067": r2(KZ067), "065": r2(KZ065), "erp_gross": r2(ERP_KZ), "erp_av": r2(ERP_AV), "upside": r2(UPS), "ao_excluded": r2(AO_KZ),
                  "ao_unreimbursed": r2(AO_UNR), "ao_groups": AO_GROUPS,
                  "ao_unreimbursed_067_065": [r2(sum(Dd(x["kz067"]) for x in SUP if x["ao_only"])),
                                              r2(sum(Dd(x["kz065"]) for x in SUP if x["ao_only"]))],
                  "book_excess": r2(KZ067 + KZ065 - ERP_KZ),
                  "mshp_suppliers": [{"name": x["name"], "kz065": x["kz065"], "rows": x["book_rows065"],
                                      "erp_rest": x["kz065"] > x["book065"] + 0.005 and not x["ao_only"],
                                      "ao": bool(x["ao_only"])} for x in SUP if x["kz065"] > 0]},
           "av": {"items": AVL, "zachet": r2(AV_Z), "zachet_by_house": {h: r2(v) for h, v in zachet.items()},
                  "trade": r2(sum(Dd(a["av"]) for a in AVL if "нетовар" not in a["decision"]))},
           "stocks": {"P_net": r2(P_net), "P_gross": r2(P_gt), "houses_net": {h: r2(bal[h]) for h in SCOPE8 + ["30-3"]},
                      "built_gross": r2(built_gt), "own3_gross": r2(sum(o["stkh|30-3|gt"] for o in ON.values())),
                      "under_norm": r2(used_tot["P"] + used_tot["P:alias"]), "own3_used": r2(used_tot["own"]),
                      "surplus_after": r2(stock_left), "surplus_cache": r2(surplus_c), "septic_left_components": r2(sept_left_g),
                      "excluded": excl, "excluded_sum": excl_sum, "erp_stock": L["erp_stock"]},
           "consumed": {"net": {h: r2(cons_n[h]) for h in SCOPE8 + ["30-3"]}, "gross8": r2(cons8_gt), "gross_30_3": r2(c3_gt)},
           "deficit": {"cache": {"total": r2(G_cache), "by_house": {h: r2(DC[h]) for h in HOUSES}},
                       "folded": {"total": r2(G_tot), "by_house": {h: r2(G[h]) for h in HOUSES}}, "adjust": adjust,
                       "items": sorted(items, key=lambda i: -i["sum"]),
                       "identity": {"norm5": r2(norm5_d), "cons3_g": r2(c3_gt), "own_used": r2(own_used_c), "under_norm": r2(under_norm),
                                    "cheap": r2(cheap), "surplus": r2(surplus_c)}},
           "mshp_h": {"per_house": r2(mshp_rate), "by_house": {h: r2(v) for h, v in MSHP.items()},
                      "houses_built": {h: r2(v) for h, v in mshp_h.items()}, "max_per_house": r2(max(mshp_h["30-1"], mshp_h["30-2"])),
                      "confirmed": mshp_conf, "pay065": {k: r2(v) for k, v in pay065.items()}, "pu_mshp_net": r2(pu_msp),
                      "msp_to_irc_net": r2(msp_to_irc), "msp_balance_net": r2(msp_bal), "irc_left_on_ib_mshp_net": r2(irc_msp_left),
                      "irc_left_on_ib_mshp_gross": r2(irc_msp_left_g),
                      "built_net": {h: r2(sum(v[h][2] for v in mshp_items.values())) for h in ("30-1", "30-2", "30-3")},
                      "items": [{"name": n, "net": {h: r2(v[h][2]) for h in ("30-1", "30-2", "30-3")},
                                 "gross": {h: r2(v[h][1]) for h in ("30-1", "30-2", "30-3")}}
                                for n, v in sorted(mshp_items.items(), key=lambda x: -sum(y[1] for y in x[1].values()))],
                      "excluded": mshp_excl_tab, "intersect_G": mshp_inter_G,
                      "rule": "МШП на будинок = менший з видач на 30 м №1 і №2 (за сумою з ПДВ партій), №3 — за вирахуванням уже виданого; "
                              "лише група «мшп» без норми в картці 000000009: позиції норми з рах. 22 (дріт, електроди, обрізні диски) "
                              "і аліаси норми (клей монтажний Lacrysil → «клей», круг відрізний 230 → «Обрізні диски») — у G_h"},
           "risks": RISKS,
           "sens": sens,
           "inv": inv}
    print("\nКонтрольные итоги §3.2:")
    for c in controls:
        print(f"  {c['n']:>2}. {'OK ' if c['ok'] else 'НЕТ'} {c['name']}\n       исследование: {c['research']}\n       моё: {c['mine']}"
              + (f"\n       ({c['note']})" if c['note'] else ""))
    save("irc_fakt.json", out)
    bad = [c["n"] for c in controls if not c["ok"] and c["n"] <= 10]
    if bad:
        sys.exit(f"СТОП: не сошлись контрольные строки {bad} (промт §0)")
