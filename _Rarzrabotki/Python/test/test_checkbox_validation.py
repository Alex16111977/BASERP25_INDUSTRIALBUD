# -*- coding: utf-8 -*-
"""
Валідація: перевірка що галочка ПоказыватьТолькоРасхождения
коректно додана в XML форми та BSL модуль.
"""

import os


def check_form_xml():
    """Перевірка Form.xml"""
    path = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьТовары\Forms\Форма\Ext\Form.xml"
    with open(path, "r", encoding="utf-8-sig") as f:
        content = f.read()

    checks = []

    # 1. CheckBoxField існує
    ok = '<CheckBoxField name="ПоказыватьТолькоРасхождения" id="158">' in content
    checks.append(("CheckBoxField в формі", ok))

    # 2. DataPath прив'язаний до реквізиту об'єкта
    ok = "<DataPath>Объект.ПоказыватьТолькоРасхождения</DataPath>" in content
    checks.append(("DataPath = Объект.ПоказыватьТолькоРасхождения", ok))

    # 3. Подія OnChange зареєстрована
    ok = '<Event name="OnChange">ПоказыватьТолькоРасхожденияПриИзменении</Event>' in content
    checks.append(("OnChange подія зареєстрована", ok))

    # 4. Збереження між сеансами
    ok = "<Field>Объект.ПоказыватьТолькоРасхождения</Field>" in content
    checks.append(("Save field для збереження між сеансами", ok))

    # 5. ContextMenu та ExtendedTooltip існують (обов'язкові елементи)
    ok = 'name="ПоказыватьТолькоРасхожденияКонтекстноеМеню"' in content
    checks.append(("ContextMenu елемент", ok))
    ok = 'name="ПоказыватьТолькоРасхожденияРасширеннаяПодсказка"' in content or \
         'name="ПоказыватьТолькоРасхожденияExtendedTooltip"' in content
    checks.append(("ExtendedTooltip елемент", ok))

    # 6. ID унікальні (158, 159, 160 не дублюються)
    for id_val in ["158", "159", "160"]:
        count = content.count(f'id="{id_val}"')
        ok = count == 1
        checks.append((f"ID {id_val} унікальний (знайдено {count}x)", ok))

    # 7. RowFilter таблиці документів (має бути nil — фільтр ставиться програмно)
    ok = '<RowFilter xsi:nil="true"/>' in content
    checks.append(("RowFilter таблиці = nil (фільтр програмний)", ok))

    return checks


def check_module_bsl():
    """Перевірка Module.bsl"""
    path = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьТовары\Forms\Форма\Ext\Form\Module.bsl"
    with open(path, "r", encoding="utf-8-sig") as f:
        content = f.read()

    checks = []

    # 1. Обробник зміни галочки
    ok = "Процедура ПоказыватьТолькоРасхожденияПриИзменении(Элемент)" in content
    checks.append(("Процедура ПоказыватьТолькоРасхожденияПриИзменении", ok))

    # 2. Процедура ПрименитьФильтрДокументов
    ok = "Процедура ПрименитьФильтрДокументов()" in content
    checks.append(("Процедура ПрименитьФильтрДокументов", ok))

    # 3. Фільтр по Синхронизировать = Истина
    ok = 'ФиксированнаяСтруктура("Синхронизировать", Истина)' in content
    checks.append(("Фільтр по Синхронизировать=Истина", ok))

    # 4. Зняття фільтру (Неопределено)
    ok = "ОтборСтрок = Неопределено" in content
    checks.append(("Зняття фільтру = Неопределено", ok))

    # 5. Виклик з обробника OnChange
    ok = "ПрименитьФильтрДокументов()" in content
    checks.append(("Виклик ПрименитьФильтрДокументов з OnChange", ok))

    # 6. Виклик після аналізу документів
    # Перевіряємо що після АнализДокументовНаСервере() є виклик фільтру
    idx_analiz = content.find("АнализДокументовНаСервере()")
    if idx_analiz >= 0:
        after = content[idx_analiz:idx_analiz+200]
        ok = "ПрименитьФильтрДокументов()" in after
    else:
        ok = False
    checks.append(("ПрименитьФильтрДокументов після АнализДокументовНаСервере", ok))

    # 7. Директива &НаКлиенте перед обробниками
    ok = "&НаКлиенте" in content
    checks.append(("Директива &НаКлиенте присутня", ok))

    # 8. Обробник ПоказыватьТолькоРасхожденияПриИзменении має &НаКлиенте
    idx = content.find("Процедура ПоказыватьТолькоРасхожденияПриИзменении")
    if idx >= 0:
        before = content[max(0, idx-50):idx]
        ok = "&НаКлиенте" in before
    else:
        ok = False
    checks.append(("&НаКлиенте перед ПоказыватьТолькоРасхожденияПриИзменении", ok))

    # 9. ПрименитьФильтрДокументов має &НаКлиенте
    idx = content.find("Процедура ПрименитьФильтрДокументов()")
    if idx >= 0:
        before = content[max(0, idx-50):idx]
        ok = "&НаКлиенте" in before
    else:
        ok = False
    checks.append(("&НаКлиенте перед ПрименитьФильтрДокументов", ok))

    # 10. ТаблицаДокументовФорма — правильне ім'я елемента
    ok = "Элементы.ТаблицаДокументовФорма.ОтборСтрок" in content
    checks.append(("Використовується Элементы.ТаблицаДокументовФорма.ОтборСтрок", ok))

    return checks


def check_processor_xml():
    """Перевірка що реквізит існує в обробці"""
    path = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьТовары.xml"
    with open(path, "r", encoding="utf-8-sig") as f:
        content = f.read()

    checks = []

    # 1. Реквізит ПоказыватьТолькоРасхождения існує
    ok = "<Name>ПоказыватьТолькоРасхождения</Name>" in content
    checks.append(("Реквізит в обробці XML", ok))

    # 2. Тип Boolean
    ok = "<v8:Type>xs:boolean</v8:Type>" in content
    checks.append(("Тип xs:boolean", ok))

    return checks


def main():
    print("=" * 70)
    print("ВАЛІДАЦІЯ: Галочка ПоказыватьТолькоРасхождения")
    print("=" * 70)

    all_ok = True

    for section_name, check_func in [
        ("Обробка XML (реквізит)", check_processor_xml),
        ("Form.xml (елемент форми)", check_form_xml),
        ("Module.bsl (код обробників)", check_module_bsl),
    ]:
        print(f"\n--- {section_name} ---")
        checks = check_func()
        for desc, ok in checks:
            status = "OK" if ok else "FAIL"
            mark = "  " if ok else "!!"
            print(f"  {mark} [{status:4s}] {desc}")
            if not ok:
                all_ok = False

    print("\n" + "=" * 70)
    if all_ok:
        print("РЕЗУЛЬТАТ: ВСІ ПЕРЕВІРКИ ПРОЙШЛИ УСПІШНО!")
    else:
        print("РЕЗУЛЬТАТ: Є помилки — див. FAIL вище")
    print("=" * 70)


if __name__ == "__main__":
    main()
