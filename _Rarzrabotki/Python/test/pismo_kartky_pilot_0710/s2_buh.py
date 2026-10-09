# -*- coding: utf-8 -*-
"""Шаг 2: BuhBud — листья простых эталонов (РазвернутьДоКарточек, Истина) + позиции СС. Только чтение."""
import sys, json, time
from conn import buh, uid

t0 = time.time()
c = buh()
print("connect %.1fs" % (time.time() - t0))
S = c.String

def ref_ok(x):
    return x is not None and not isinstance(x, (str, int, float, bool))

out = {}
for code in ("000000101", "000000005", "000000099"):
    t1 = time.time()
    sp = c.Справочники.А_СпецификацииДомов.НайтиПоКоду(code)
    cc = sp.СтруктураСебестоимости
    info = {"code": code, "name": sp.Наименование, "status": S(sp.Статус), "vid": S(sp.ВидСпецификации),
            "cc_code": cc.Код if ref_ok(cc) else "", "cc_name": cc.Наименование if ref_ok(cc) else "",
            "cc_uuid": uid(c, cc) if ref_ok(cc) else "",
            "cc_domov": float(cc.КоличествоДомов) if ref_ok(cc) else 0}
    res = c.А_СпецификацииДомовСервер.РазвернутьДоКарточек(sp, True)
    L = res.Листья
    leaves = []
    for i in range(L.Количество()):
        s = L.Получить(i)
        n = s.Номенклатура
        on = s.ОбщееНазвание
        et = s.ЭтапСС
        leaves.append({
            "key": s.КлючЛиста, "vid": S(s.ВидСтроки),
            "nom_uuid": uid(c, n), "nom_name": n.Наименование if ref_ok(n) else "", "nom_code": n.Код if ref_ok(n) else "",
            "nom_ed": S(n.БазоваяЕдиницаИзмерения) if ref_ok(n) else "",
            "nom_on": S(n.ОбщееНазваниеНоменклатуры) if ref_ok(n) else "",
            "on": on.Наименование if ref_ok(on) else "", "on_uuid": uid(c, on),
            "etap": et.Наименование if ref_ok(et) else "", "etap_uuid": uid(c, et), "etap_code": et.Код if ref_ok(et) else "",
            "pos": s.ПозицияСС, "ed_str": S(s.ЕдиницаСтроки), "kol_ed_str": float(s.КоличествоЕдСтроки),
            "kol": float(s.Количество), "price": float(s.ЦенаНормы), "sum": float(s.СуммаНормы),
            "path": s.ПутьУзла, "lvl": int(s.Уровень), "np": bool(s.НеПоставляется), "k": float(s.Коэффициент),
        })
    # позиции СС BuhBud
    poss = []
    if ref_ok(cc):
        q = c.NewObject("Запрос")
        q.Text = """
ВЫБРАТЬ
	К.НомерСтроки КАК НомерСтроки,
	К.Этап КАК Этап,
	К.Этап.Наименование КАК ЭтапИмя,
	К.НоменклатураСС КАК НоменклатураСС,
	К.ОбщееНазвание.Наименование КАК ОНИмя,
	К.Количество КАК Количество,
	К.Сумма КАК Сумма,
	К.ЕдиницаСС КАК ЕдиницаСС
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК К
ГДЕ
	К.Ссылка = &СС
УПОРЯДОЧИТЬ ПО
	К.НомерСтроки"""
        q.SetParameter("СС", cc)
        r = q.Execute().Выгрузить()
        for i in range(r.Количество()):
            s = r.Получить(i)
            poss.append({"n": int(s.НомерСтроки), "etap": s.ЭтапИмя or "", "etap_uuid": uid(c, s.Этап),
                         "pos": s.НоменклатураСС, "on": s.ОНИмя or "", "kol": float(s.Количество),
                         "sum": float(s.Сумма), "ed": S(s.ЕдиницаСС)})
    info["leaves"] = leaves
    info["cc_pos"] = poss
    out[code] = info
    print(code, info["name"], "|", info["status"], "| СС", info["cc_code"], info["cc_name"], "домов", info["cc_domov"],
          "| листьев", len(leaves), "Σнорм %.2f" % sum(x["sum"] for x in leaves),
          "НП %.2f" % sum(x["sum"] for x in leaves if x["np"]),
          "| позиций СС", len(poss), "ΣСС %.2f" % sum(x["sum"] for x in poss), "| %.1fs" % (time.time() - t1))
json.dump(out, open("buh_leaves.json", "w", encoding="utf-8"), ensure_ascii=False, indent=0)
print("всего %.1fs" % (time.time() - t0))
