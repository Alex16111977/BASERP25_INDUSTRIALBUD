# -*- coding: utf-8 -*-
"""
Smoke обработки СинхронизироватьТоварыТолькоТовары на живой BaseERP (read-only).
Загружаем .epf через COM, период янв-2026:
  1) СравнитьОстатки() — без исключения; в ТаблицаРасхождений только Раздел="Товари";
  2) АнализироватьДокументы(0) — per-row анализ одной позиции, без исключения.
Синхронизацию НЕ вызываем (она меняет документы — это делает пользователь в UI).

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
NP = datetime.datetime(2026, 1, 1, 12, 0, 0)
KP = datetime.datetime(2026, 1, 31, 12, 0, 0)


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN_ERP)
    fails = 0

    обр = erp.ВнешниеОбработки.Создать(EPF, False)
    обр.НачалоПериода = NP
    обр.ОкончаниеПериода = KP
    print("EPF загружена через COM:", обр.Метаданные().Имя)

    # 1) СравнитьОстатки
    try:
        итог = обр.СравнитьОстатки()
        тр = обр.ТаблицаРасхождений
        n = тр.Количество()
        # проверка: только Товари
        чужие = 0
        for i in range(n):
            if str(тр.Получить(i).Раздел).strip() != "Товари":
                чужие += 1
        print(f"[1] СравнитьОстатки: OK — {итог}; рядків={n}; не-Товари={чужие}")
        if чужие > 0:
            fails += 1
            print("    FAIL: есть строки с Раздел != Товари")
    except Exception as e:
        fails += 1
        ei = getattr(e, 'excepinfo', None)
        print(f"[1] СравнитьОстатки: FAIL: {ei[2] if ei else e}")
        return fails

    # 2) АнализироватьДокументы(-1) по первым 10 отмеченным позициям (пакетный путь Фазы 2)
    try:
        тр = обр.ТаблицаРасхождений
        n = тр.Количество()
        for i in range(n):
            тр.Получить(i).Синхронизировать = (i < 10)
        итог2 = обр.АнализироватьДокументы(-1)
        тд = обр.ТаблицаДокументов
        print(f"[2] АнализироватьДокументы(-1) перших 10 поз.: OK — {итог2}; "
              f"документів={тд.Количество()}")
    except Exception as e:
        fails += 1
        ei = getattr(e, 'excepinfo', None)
        print(f"[2] АнализироватьДокументы(-1): FAIL: {ei[2] if ei else e}")

    print(f"\n=== SMOKE: {'PASS' if fails == 0 else str(fails) + ' FAIL'} ===")
    return fails


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
