# -*- coding: utf-8 -*-
"""Инвариант ETC: нет закупок -> План на факт = План, Прогноз = План, % = 100.

Требование бухгалтера «не играться с копейками»: при отсутствии закупок остаток плана
берётся РОВНО суммой плана, без round-trip через среднюю цену.
"""
import os
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\ПланФактССБухгалтерія.erf"
CONN = 'Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'


def main():
    assert os.path.exists(ERF), f"Не собран .erf: {ERF}"

    v8 = win32com.client.Dispatch("V83.COMConnector")
    b = v8.Connect(CONN)
    отчёт = b.ВнешниеОтчеты.Создать(ERF)

    z = b.NewObject("Запрос")
    z.Текст = ("ВЫБРАТЬ ДАТАВРЕМЯ(2020, 1, 1) КАК Н, "
               "КОНЕЦПЕРИОДА(ДАТАВРЕМЯ(2020, 12, 31), ДЕНЬ) КАК К")
    гр = z.Выполнить().Выгрузить()[0]

    период = b.NewObject("СтандартныйПериод")
    период.ДатаНачала = гр.Н
    период.ДатаОкончания = гр.К

    тз = отчёт.ПолучитьДанные(период)

    план = факт = etc = прогноз = 0.0
    for i in range(тз.Количество()):
        r = тз[i]
        план += r.ПланГрн
        факт += r.ФактГрн
        etc += r.ПланНаФактГрн
        прогноз += r.ПрогнозГрн

    assert abs(факт) < 0.01, f"В 2020 году не должно быть закупок, а факт = {факт:,.2f}"
    assert abs(etc - план) < 0.01, f"ETC {etc:,.2f} != План {план:,.2f}"
    assert abs(прогноз - план) < 0.01, f"Прогноз {прогноз:,.2f} != План {план:,.2f}"

    print(f"OK: План={план:,.2f} Факт={факт:,.2f} ПланНаФакт={etc:,.2f} Прогноз={прогноз:,.2f}")
    print("ETC INVARIANT PASSED")


if __name__ == "__main__":
    main()
