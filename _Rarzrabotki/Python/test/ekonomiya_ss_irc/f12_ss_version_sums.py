# -*- coding: utf-8 -*-
"""Ф0. История СС в BuhBud: по каждой версии (ВерсииОбъектов) Σ Комплектующие.Количество/Сумма и число строк.
Печатает только версии, где итоги изменились. Только чтение."""
import os, sys, base64, re
import xml.etree.ElementTree as ET
sys.path.insert(0, os.path.dirname(__file__))
from common import *

Bh = buh()
vt = q(Bh, """
ВЫБРАТЬ С.Наименование КАК СС, Вер.НомерВерсии КАК N, Вер.ДатаВерсии КАК Дата, Вер.АвторВерсии КАК Автор, Вер.ВерсияОбъекта КАК Хран
ИЗ РегистрСведений.ВерсииОбъектов КАК Вер
  ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.СтруктураСебестоимости КАК С ПО Вер.Объект = С.Ссылка
ГДЕ С.Наименование ПОДОБНО "%IR%"
УПОРЯДОЧИТЬ ПО СС, N
""")
res = []
prev = {}
for i in range(vt.Количество()):
    r = vt.Получить(i)
    ss = r.СС; n = r.N; dt = r.Дата; author = Bh.String(r.Автор)
    try:
        bd = r.Хран.Получить()
        raw = base64.b64decode(Bh.Base64Строка(bd)) if bd is not None else b""
    except Exception as e:
        raw = b""
    txt = raw.decode("utf-8-sig", errors="ignore")
    # ТЧ Комплектующие: блок <Комплектующие> ... <Row>…<Количество>x</Количество>…<Сумма>y</Сумма>
    kol = D(0); summ = D(0); cnt = 0; domov = None
    m = re.search(r"<КоличествоДомов>([^<]*)</КоличествоДомов>", txt)
    if m: domov = m.group(1)
    blk = re.search(r"<Комплектующие>(.*?)</Комплектующие>", txt, re.S)
    if blk:
        for row in re.findall(r"<Row>(.*?)</Row>", blk.group(1), re.S):
            k = re.search(r"<Количество>([^<]*)</Количество>", row)
            s = re.search(r"<Сумма>([^<]*)</Сумма>", row)
            kol += D(k.group(1) if k else 0); summ += D(s.group(1) if s else 0); cnt += 1
    key = (cnt, kol, summ, domov)
    rec = {"СС": ss, "N": n, "Дата": str(dt)[:19], "Автор": author, "строк": cnt, "Σкол": str(kol), "Σсумма": str(summ),
           "Домов": domov, "len": len(txt)}
    res.append(rec)
    if prev.get(ss) != key:
        print(f"{ss[:14]} v{n:>3} {str(dt)[:19]} {author:<12} строк={cnt:<4} Σкол={kol:<14} Σсумма={fmt(summ):>15} Домов={domov} xml={len(txt)}")
        prev[ss] = key
dump_json(res, "ss_versions_buh.json")
