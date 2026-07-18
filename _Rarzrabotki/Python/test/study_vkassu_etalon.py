"""
Изучение эталонной ВКассу 000Ц-000001 от 31.01.2026 (заполнено типовым способом).
Цель: понять структуру ТЧ Зарплата + шапки, чтобы повторить в нашем алгоритме.
"""
import io
import sys
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def safe_name(ref):
    if ref is None:
        return "<None>"
    try:
        if ref.Пустая():
            return "<пусто>"
    except:
        pass
    try:
        return str(ref.Наименование) if hasattr(ref, "Наименование") else str(ref)
    except:
        try:
            return str(ref)
        except:
            return "<unprintable>"


def safe_str(v):
    if v is None:
        return "<None>"
    try:
        return str(v)
    except:
        return "<unprintable>"


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN_ERP)

    # Найти эталон ВКассу
    q = erp.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 Вед.Ссылка
ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу КАК Вед
ГДЕ Вед.Номер = "000Ц-000001"
    И НАЧАЛОПЕРИОДА(Вед.Дата, ДЕНЬ) = ДАТАВРЕМЯ(2026,1,31)
"""
    rs = q.Выполнить()
    if rs.Пустой():
        print("FAIL: эталон не найден")
        sys.exit(1)
    sel = rs.Выбрать(); sel.Следующий()
    vk_ref = sel.Ссылка
    vk = vk_ref.ПолучитьОбъект()
    print(f"=== ВКассу: {vk_ref} ===\n")

    # Шапка
    print("# ШАПКА")
    print(f"  Дата               = {vk.Дата}")
    print(f"  Номер              = {vk.Номер}")
    print(f"  ПериодРегистрации  = {vk.ПериодРегистрации}")
    print(f"  Организация        = {safe_name(vk.Организация)}")
    print(f"  Подразделение      = {safe_name(vk.Подразделение)}")
    print(f"  Касса              = {safe_name(vk.Касса)}")
    print(f"  СпособВыплаты      = {safe_name(vk.СпособВыплаты)}")
    print(f"  СтатьяФинансирования = {safe_name(vk.СтатьяФинансирования)}")
    print(f"  СтатьяРасходов     = {safe_name(vk.СтатьяРасходов)}")
    print(f"  ПроцентВыплаты     = {vk.ПроцентВыплаты}")
    print(f"  СуммаПоДокументу   = {vk.СуммаПоДокументу}")
    print(f"  ПорядокЗаполненияНалогов = {vk.ПорядокЗаполненияНалогов}")
    print(f"  ВыплатаОтраженаВБухучете = {vk.ВыплатаОтраженаВБухучете}")
    print(f"  Проведен           = {vk.Проведен}")
    print(f"  КраткийСоставДокумента = {safe_str(vk.КраткийСоставДокумента)[:100]}")

    # ТЧ Зарплата
    print(f"\n# ТЧ Зарплата ({vk.Зарплата.Количество()} строк)")
    for i in range(vk.Зарплата.Количество()):
        r = vk.Зарплата.Получить(i)
        print(f"  [{i+1}] Сотр={safe_name(r.Сотрудник)} | ФЛ={safe_name(r.ФизическоеЛицо)}")
        print(f"       Подр={safe_name(r.Подразделение)} | ПериодВзаим={r.ПериодВзаиморасчетов}")
        print(f"       СтФин={safe_name(r.СтатьяФинансирования)} | СтРасх={safe_name(r.СтатьяРасходов)}")
        print(f"       ДокОснование={safe_name(r.ДокументОснование)} ({type(r.ДокументОснование).__name__})")
        print(f"       КВыплате={r.КВыплате} | Комп={r.КомпенсацияЗаЗадержкуЗарплаты}")
        print(f"       ГруппаУч={safe_name(r.ГруппаУчетаНачислений)}")

    # ТЧ Состав
    print(f"\n# ТЧ Состав ({vk.Состав.Количество()} строк)")
    for i in range(min(5, vk.Состав.Количество())):
        r = vk.Состав.Получить(i)
        print(f"  [{i+1}] ФЛ={safe_name(r.ФизическоеЛицо)}")

    # ТЧ НДФЛ
    print(f"\n# ТЧ НДФЛ ({vk.НДФЛ.Количество()} строк)")
    for i in range(min(5, vk.НДФЛ.Количество())):
        r = vk.НДФЛ.Получить(i)
        print(f"  [{i+1}] ФЛ={safe_name(r.ФизическоеЛицо)} | Месяц={r.МесяцНалоговогоПериода}")
        print(f"       КодДох={safe_name(r.КодДохода)} | Сумма={r.Сумма} | Доход={r.Доход}")
        print(f"       ДокОснование={safe_name(r.ДокументОснование)} ({type(r.ДокументОснование).__name__})")
        print(f"       ГруппаУчУд={safe_name(r.ГруппаУчетаУдержаний)}")
        print(f"       ПериодВзаим={r.ПериодВзаиморасчетов}")

    # ТЧ ВзносыФОТ
    print(f"\n# ТЧ ВзносыФОТ ({vk.ВзносыФОТ.Количество()} строк)")
    for i in range(min(3, vk.ВзносыФОТ.Количество())):
        r = vk.ВзносыФОТ.Получить(i)
        print(f"  [{i+1}] ФЛ={safe_name(r.ФизическоеЛицо)} | Налог={safe_name(r.Налог)}")
        print(f"       Сумма={r.Сумма} | ПериодВзаим={r.ПериодВзаиморасчетов}")

    # ТЧ Основания
    print(f"\n# ТЧ Основания ({vk.Основания.Количество()} строк)")
    for i in range(min(3, vk.Основания.Количество())):
        r = vk.Основания.Получить(i)
        print(f"  [{i+1}] Документ={safe_name(r.Документ)} ({type(r.Документ).__name__})")


if __name__ == "__main__":
    main()
