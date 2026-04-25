# -*- coding: utf-8 -*-
"""
Проверка баланса ІБ00-000597 СО ВСЕМИ статьями (включая ВложенияСобственныхСредств / ВыводСобственныхСредств).
Если без исключения статей баланс по организации сойдётся — значит, текущая обработка неправильно исключает эти статьи.
"""
import sys
import win32com.client
import datetime

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN_ERP)
print("Connected\n")


def analyze_all_articles(uuid_str, label):
    uid = conn.NewObject("УникальныйИдентификатор", uuid_str)
    ref = conn.Документы.ПриобретениеТоваровУслуг.ПолучитьСсылку(uid)

    print("=" * 70)
    print(f"{label} — ВСЕ статьи (без исключений)")
    print("=" * 70)

    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Об.Подразделение КАК Подразделение,
        Об.Статья КАК Статья,
        Об.СуммаПриход КАК Дебет,
        Об.СуммаРасход КАК Кредит,
        Об.СуммаПриход - Об.СуммаРасход КАК Контроль
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(, , Регистратор, ) КАК Об
    ГДЕ Об.Регистратор = &Регистратор
    """
    q.УстановитьПараметр("Регистратор", ref)
    table = q.Выполнить().Выгрузить()

    by_podr = {}
    by_article = {}
    for i in range(table.Количество()):
        row = table.Получить(i)
        podr = conn.String(row.Подразделение)
        st = conn.String(row.Статья)
        c = float(row.Контроль)
        by_podr[podr] = by_podr.get(podr, 0.0) + c
        by_article[st] = by_article.get(st, 0.0) + c

    print(f"\n  По подразделениям:")
    total = 0.0
    for k, v in sorted(by_podr.items(), key=lambda kv: -kv[1]):
        total += v
        if round(v, 2) != 0:
            m = "+" if v > 0 else "-"
            print(f"    {m} {k:<50s} {v:>15.2f}")

    print(f"\n  По статьям (сводно):")
    for k, v in sorted(by_article.items(), key=lambda kv: -abs(kv[1])):
        if round(v, 2) != 0:
            print(f"    {k:<50s} {v:>15.2f}")

    print(f"\n  Итого по всем подразделениям и статьям: {total:.2f}")
    return total


analyze_all_articles("20b3ae2a-f767-11f0-8105-00155dce3d04", "ІБ00-008058 от 31.12.2025")
print()
analyze_all_articles("3732c824-08ad-11f1-8106-00155dce3d04", "ІБ00-000597 от 10.02.2026")
