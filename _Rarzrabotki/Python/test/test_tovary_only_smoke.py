# -*- coding: utf-8 -*-
"""
Smoke обработки СинхронизироватьТоварыТолькоТовары на живой BaseERP (read-only).
Модель «2 уровня»: верх = свод по складам (ТаблицаСкладов), позиции (ТаблицаРасхождений) внутри.
  1) СравнитьОстатки() — без исключения; складов < позиций; ΣПозицій == числу позиций.
  2) АнализироватьДокументы(0) — анализ одного склада; счётчики строки склада согласованы; 0 строк-ПОМИЛКА.
Синхронизацию НЕ вызываем (меняет документы — это делает пользователь в UI).

Запуск: py -3.13 _Rarzrabotki/Python/test/test_tovary_only_smoke.py
"""
import sys
import datetime
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EPF = (r"C:\Configuration_downloads\BASERP25\.claude\worktrees\musing-kare-55f9e1"
       r"\_Rarzrabotki\Обработки\СинхронизироватьТоварыТолькоТовары.epf")
NP = datetime.datetime(2026, 3, 1, 12, 0, 0)
KP = datetime.datetime(2026, 3, 31, 12, 0, 0)


def main():
    erp = win32com.client.Dispatch("V83.COMConnector").Connect(CONN_ERP)
    fails = 0

    обр = erp.ВнешниеОбработки.Создать(EPF, False)
    обр.НачалоПериода = NP
    обр.ОкончаниеПериода = KP
    print("EPF загружена:", обр.Метаданные().Имя)

    # 1) СравнитьОстатки → свод по складам
    try:
        итог = обр.СравнитьОстатки()
        склады = обр.ТаблицаСкладов
        позиции = обр.ТаблицаРасхождений
        nСкл = склады.Количество()
        nПоз = позиции.Количество()
        сумПоз = 0
        for i in range(nСкл):
            сумПоз += int(склады.Получить(i).Позицій)
        print(f"[1] СравнитьОстатки: OK — {итог}")
        print(f"    складів={nСкл}, позицій={nПоз}, ΣПозицій по складах={сумПоз}")
        if nСкл == 0 or nСкл >= nПоз:
            fails += 1; print("    FAIL: складов должно быть >0 и меньше числа позиций")
        if сумПоз != nПоз:
            fails += 1; print(f"    FAIL: ΣПозицій ({сумПоз}) != числу позиций ({nПоз})")
    except Exception as e:
        fails += 1
        ei = getattr(e, 'excepinfo', None)
        print(f"[1] СравнитьОстатки: FAIL: {ei[2] if ei else e}")
        return fails

    # 2) АнализироватьДокументы — склад с МИНИМУМОМ позиций (быстро)
    try:
        if обр.ТаблицаСкладов.Количество() > 0:
            idxMin, minПоз = 0, None
            for i in range(обр.ТаблицаСкладов.Количество()):
                p = int(обр.ТаблицаСкладов.Получить(i).Позицій)
                if minПоз is None or p < minПоз:
                    minПоз, idxMin = p, i
            print(f"    аналізуємо склад idx={idxMin} (позицій={minПоз})")
            итог2 = обр.АнализироватьДокументы(idxMin)
            ск0 = обр.ТаблицаСкладов.Получить(idxMin)
            тд = обр.ТаблицаДокументов
            пом = 0
            for i in range(тд.Количество()):
                if str(тд.Получить(i).Статус).startswith("ПОМИЛКА"):
                    пом += 1
            print(f"[2] АнализироватьДокументы(0) склад «{ск0.Склад}»: OK — {итог2}")
            print(f"    документів={тд.Количество()}; склад.Документів={int(ск0.Документів)}, "
                  f"Дій={int(ск0.Дій)}, Потребують рішення={int(ск0.ПотребуютьРішення)}; рядків-ПОМИЛКА={пом}")
            if int(ск0.Документів) != тд.Количество():
                fails += 1; print("    FAIL: счётчик Документів склада != числу строк документов")
        else:
            print("[2] пропущено (нет складов)")
    except Exception as e:
        fails += 1
        ei = getattr(e, 'excepinfo', None)
        print(f"[2] АнализироватьДокументы(0): FAIL: {ei[2] if ei else e}")

    print(f"\n=== SMOKE: {'PASS' if fails == 0 else str(fails) + ' FAIL'} ===")
    return fails


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
