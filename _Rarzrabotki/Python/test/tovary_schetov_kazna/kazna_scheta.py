# -*- coding: utf-8 -*-
"""Счета Казны (СчетНаОплатуПоставщика) 2026 по «МД IRC 2026» (с дочерними) и «МД ООН Админ» + их файлы.

Только чтение Казны. Файлы копируются с тома (ПолныйПутьWindows + ПутьКФайлу) в RAB/<uid>/.
Результат — RAB/scheta.json.

    python kazna_scheta.py
"""
import json
import os
import shutil
import sys

import win32com.client

sys.stdout.reconfigure(encoding='utf-8')
RAB = os.environ.get("TOVARY_RAB", r"C:\Razrabotki\tovary_schetov_kazna")
KAZNA = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'

Q = """ВЫБРАТЬ
	С.Ссылка КАК Счет, С.НомерВходящегоДокумента КАК Номер, С.ДатаВходящегоДокумента КАК ДатаВх, С.Дата КАК Дата,
	С.Контрагент.Наименование КАК КонтрНаим, С.Контрагент.ЕГРПОУ КАК ЕГРПОУ, С.Подразделение.Наименование КАК Подр,
	С.СуммаСчета КАК СуммаСчета, С.СуммаДокумента КАК СуммаДок
ИЗ Документ.СчетНаОплатуПоставщика КАК С
ГДЕ НЕ С.ПометкаУдаления И С.Дата МЕЖДУ ДАТАВРЕМЯ(2026, 1, 1) И ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59)
	И (С.Подразделение В ИЕРАРХИИ (&IRC) ИЛИ С.Подразделение = &Адмін)
УПОРЯДОЧИТЬ ПО С.Дата"""

QF = """ВЫБРАТЬ Ф.ВладелецФайла КАК Счет, Ф.Наименование КАК Имя, Ф.ТекущаяВерсияРасширение КАК Расш,
	Ф.ТекущаяВерсияТом.ПолныйПутьWindows КАК Том, Ф.ТекущаяВерсияПутьКФайлу КАК Путь, Ф.ТекущаяВерсияРазмер КАК Размер
ИЗ Справочник.Файлы КАК Ф
ГДЕ Ф.ВладелецФайла В (&Счета) И НЕ Ф.ПометкаУдаления"""


def main():
    k = win32com.client.Dispatch("V83.COMConnector").Connect(KAZNA)
    S = k.String
    irc = k.Справочники.Подразделения.НайтиПоНаименованию("МД IRC 2026", True)
    adm = k.Справочники.Подразделения.НайтиПоНаименованию("МД ООН Админ", True)
    z = k.NewObject("Запрос")
    z.Text = Q
    z.SetParameter("IRC", irc)
    z.SetParameter("Адмін", adm)
    r = z.Execute().Выгрузить()
    scheta, refs = {}, k.NewObject("Массив")
    for i in range(r.Количество()):
        x = r.Получить(i)
        uid = S(x.Получить(0).УникальныйИдентификатор())
        refs.Добавить(x.Получить(0))
        scheta[uid] = dict(uid=uid, nomer=S(x.Получить(1)).strip(), data_vh=S(x.Получить(2))[:10], data=S(x.Получить(3))[:10],
                           kontragent=S(x.Получить(4)), edrpou=S(x.Получить(5)).strip(), podr=S(x.Получить(6)),
                           summa_scheta=float(x.Получить(7) or 0), summa_doc=float(x.Получить(8) or 0), files=[])
    z.Text = QF
    z.SetParameter("Счета", refs)
    f = z.Execute().Выгрузить()
    copied = missing = 0
    for i in range(f.Количество()):
        x = f.Получить(i)
        uid = S(x.Получить(0).УникальныйИдентификатор())
        name, ext = S(x.Получить(1)), S(x.Получить(2)).lower().strip(".")
        src = os.path.join(S(x.Получить(3)), S(x.Получить(4)))
        dst_dir = os.path.join(RAB, uid)
        os.makedirs(dst_dir, exist_ok=True)
        safe = "".join(ch for ch in name if ch not in '\\/:*?"<>|')[:80] or "file"
        dst = os.path.join(dst_dir, f"{len(scheta[uid]['files']) + 1:02d}_{safe}.{ext}")
        if os.path.exists(src):
            if not os.path.exists(dst):
                shutil.copy2(src, dst)
            copied += 1
            scheta[uid]["files"].append(dict(name=name, ext=ext, path=dst, size=int(x.Получить(5) or 0)))
        else:
            missing += 1
            scheta[uid]["files"].append(dict(name=name, ext=ext, path=None, src=src))
    os.makedirs(RAB, exist_ok=True)
    with open(os.path.join(RAB, "scheta.json"), "w", encoding="utf-8") as fh:
        json.dump(list(scheta.values()), fh, ensure_ascii=False, indent=1)
    with_files = sum(1 for s in scheta.values() if any(x["path"] for x in s["files"]))
    print(f"счетов {len(scheta)}, с файлами {with_files}, файлов скопировано {copied}, не найдено на томе {missing}")
    print("сумма счетов:", round(sum(s["summa_scheta"] for s in scheta.values()), 2))


if __name__ == "__main__":
    main()
