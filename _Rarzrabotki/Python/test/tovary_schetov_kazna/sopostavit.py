# -*- coding: utf-8 -*-
"""Сопоставление счетов Казны с заказами ЕРП и строк счетов с существующей номенклатурой (новых карточек не создаём).

Заказ: заявка ЕРП с А_НомерСчетаПоДаннымПоставщикаКазна = номер счёта и тем же ЕДРПОУ контрагента →
РасшифровкаПлатежа.ОбъектРасчетов.Объект (ЗаказПоставщику). Несколько заказов — берётся тот, чья заявка ближе по
сумме к сумме документа Казны; ничья — пусто + предупреждение.
Номенклатура: (1) позиции поступлений этого заказа, (2) поступлений этого контрагента, (3) точное совпадение
нормализованного наименования в справочнике. Сходство — доля общих значимых токенов ≥ 0,8 (для 1–2) / равенство (3).
"""
import re
from difflib import SequenceMatcher

QZ = """ВЫБРАТЬ
	Р.Ссылка.А_НомерСчетаПоДаннымПоставщикаКазна КАК Номер,
	Р.Ссылка.Контрагент.КодПоЕДРПОУ КАК ЕДРПОУ,
	Р.Ссылка.СуммаДокумента КАК СуммаЗаявки,
	Р.ОбъектРасчетов.Объект КАК Заказ
ИЗ Документ.ЗаявкаНаРасходованиеДенежныхСредств.РасшифровкаПлатежа КАК Р
ГДЕ НЕ Р.Ссылка.ПометкаУдаления И Р.Ссылка.Дата >= ДАТАВРЕМЯ(2026, 1, 1)
	И Р.Ссылка.А_НомерСчетаПоДаннымПоставщикаКазна <> ""
	И ТИПЗНАЧЕНИЯ(Р.ОбъектРасчетов.Объект) = ТИП(Документ.ЗаказПоставщику)"""

QN = """ВЫБРАТЬ РАЗЛИЧНЫЕ
	Т.Ссылка.ЗаказПоставщику КАК Заказ,
	Т.Ссылка.Контрагент.КодПоЕДРПОУ КАК ЕДРПОУ,
	Т.Номенклатура КАК Ном,
	Т.Номенклатура.Наименование КАК Наим
ИЗ Документ.ПриобретениеТоваровУслуг.Товары КАК Т
ГДЕ Т.Ссылка.Проведен И Т.Ссылка.Дата >= ДАТАВРЕМЯ(2026, 1, 1)
	И Т.Номенклатура <> ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка)"""

QS = """ВЫБРАТЬ Н.Ссылка КАК Ном, Н.Наименование КАК Наим
ИЗ Справочник.Номенклатура КАК Н
ГДЕ НЕ Н.ПометкаУдаления И НЕ Н.ЭтоГруппа"""

STOP = {"шт", "м", "мм", "см", "для", "та", "і", "и", "з", "із", "с", "в", "у", "на", "по", "до", "від", "тов", "ral"}


def norm(s):
    s = (s or "").lower().replace("ё", "е").replace("’", "'").replace("`", "'")
    s = re.sub(r"[«»\"“”&]", " ", s)
    s = re.sub(r"\b\d{8,10}\b", " ", s)          # коды УКТЗЕД в конце строки
    s = re.sub(r"[^\w\s,.x×*/-]", " ", s)
    return re.sub(r"\s+", " ", s).strip(" .,-")


def tokens(s):
    return {t for t in re.split(r"[\s,./-]+", norm(s)) if t and t not in STOP and len(t) > 1}


def similar(a, b):
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return 0.0
    jac = len(ta & tb) / len(ta | tb)
    seq = SequenceMatcher(None, norm(a), norm(b)).ratio()
    return max(jac, seq)


class Sopostavitel:
    def __init__(self, erp):
        self.erp = erp
        S = erp.String
        self.S = S
        z = erp.NewObject("Запрос")
        z.Text = QZ
        r = z.Execute().Выгрузить()
        self.zakazy = {}
        for i in range(r.Количество()):
            x = r.Получить(i)
            k = (S(x.Получить(0)).strip(), S(x.Получить(1)).strip())
            self.zakazy.setdefault(k, {})[S(x.Получить(3).УникальныйИдентификатор())] = (x.Получить(3), float(x.Получить(2) or 0))
        z.Text = QN
        r = z.Execute().Выгрузить()
        self.po_zakazu, self.po_kontr = {}, {}
        for i in range(r.Количество()):
            x = r.Получить(i)
            zak = x.Получить(0)
            nom, naim = x.Получить(2), S(x.Получить(3))
            if S(zak):
                self.po_zakazu.setdefault(S(zak.УникальныйИдентификатор()), {})[naim] = nom
            self.po_kontr.setdefault(S(x.Получить(1)).strip(), {})[naim] = nom
        z.Text = QS
        r = z.Execute().Выгрузить()
        self.spr = {}
        for i in range(r.Количество()):
            x = r.Получить(i)
            self.spr.setdefault(norm(S(x.Получить(1))), x.Получить(0))

    def zakaz(self, schet):
        """→ (ссылка | None, uid | "", предупреждение)."""
        cand = self.zakazy.get((schet["nomer"], schet["edrpou"]), {})
        if not cand:
            return None, "", "заказ ЕРП не найден"
        if len(cand) == 1:
            uid, (ref, _) = next(iter(cand.items()))
            return ref, uid, ""
        best = sorted(cand.items(), key=lambda kv: abs(kv[1][1] - schet["summa_doc"]))
        if abs(best[0][1][1] - schet["summa_doc"]) < abs(best[1][1][1] - schet["summa_doc"]):
            return best[0][1][0], best[0][0], f"заказов {len(cand)}, взят по сумме заявки"
        return None, "", f"заказов {len(cand)}, однозначно не выбран"

    def nomenklatura(self, nazv, zakaz_uid, edrpou):
        """→ (ссылка | None, источник)."""
        for src, pool in (("заказ", self.po_zakazu.get(zakaz_uid, {})), ("контрагент", self.po_kontr.get(edrpou, {}))):
            best, score = None, 0.0
            for naim, ref in pool.items():
                sc = similar(nazv, naim)
                if sc > score:
                    best, score = ref, sc
            if best is not None and score >= 0.8:
                return best, src
        ref = self.spr.get(norm(nazv))
        if ref is not None:
            return ref, "справочник"
        return None, ""
