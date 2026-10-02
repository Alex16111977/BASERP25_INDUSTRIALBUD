# Копия настроек «Системной учетной записи» почты Казна -> BuhBud (предопределённая запись, новую не создаём).
# Пароль НЕ переносится: его вводят вручную в форме BuhBud, затем «Проверить настройки».
# Запуск: без ключа — только показать разницу; с --write — записать и сверить.
# Запись: ОбменДанными.Загрузка + ОтключитьМеханизмРегистрацииОбъектов (справочник входит в планы обмена BuhBud).
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_KAZNA = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'
CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
PRIMITIVE = (str, bool, int, float)

Q_CHANGES = """
ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол
ИЗ Справочник.УчетныеЗаписиЭлектроннойПочты.Изменения КАК Изм
"""


def attr_names(conn):
    md = conn.Metadata.Справочники.УчетныеЗаписиЭлектроннойПочты
    return [md.Реквизиты.Get(i).Имя for i in range(md.Реквизиты.Количество())]


def changes_count(conn):
    q = conn.NewObject("Запрос")
    q.Text = Q_CHANGES
    sel = q.Execute().Выбрать()
    sel.Следующий()
    return sel.Кол


def main(write):
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kz = v8.Connect(CONN_KAZNA)
    buh = v8.Connect(CONN_BUH)

    kz_names = set(attr_names(kz))
    names = [a for a in attr_names(buh) if a in kz_names and not a.startswith("Удалить")]

    src = kz.Справочники.УчетныеЗаписиЭлектроннойПочты.СистемнаяУчетнаяЗаписьЭлектроннойПочты
    dst_ref = buh.Справочники.УчетныеЗаписиЭлектроннойПочты.СистемнаяУчетнаяЗаписьЭлектроннойПочты

    plan = {}
    for a in names:
        v = getattr(src, a)
        if not isinstance(v, PRIMITIVE):
            if kz.ЗначениеЗаполнено(v):
                raise RuntimeError(f"{a}: непримитивное заполненное значение {kz.String(v)!r} — нужен маппинг")
            continue  # пустая ссылка/перечисление — в BuhBud тоже пусто, не трогаем
        plan[a] = v

    diff = {a: v for a, v in plan.items() if getattr(dst_ref, a) != v}
    print(f"Реквизитов к сверке: {len(plan)}, отличаются: {len(diff)}")
    for a, v in diff.items():
        print(f"  {a}: {buh.String(getattr(dst_ref, a))!r} -> {kz.String(v)!r}")

    changes_before = changes_count(buh)
    if not write:
        print(f"Регистрации обмена (.Изменения): {changes_before}")
        print("Пробный запуск, запись не выполнялась (добавьте --write).")
        return

    obj = dst_ref.ПолучитьОбъект()
    for a, v in diff.items():
        setattr(obj, a, v)
    obj.ОбменДанными.Загрузка = True
    obj.ДополнительныеСвойства.Вставить("ОтключитьМеханизмРегистрацииОбъектов", True)
    obj.Записать()

    # Сверка: перечитываем ссылку заново и сравниваем с Казной
    dst_ref = buh.Справочники.УчетныеЗаписиЭлектроннойПочты.СистемнаяУчетнаяЗаписьЭлектроннойПочты
    bad = [a for a, v in plan.items() if getattr(dst_ref, a) != v]
    changes_after = changes_count(buh)
    print(f"Записано. Расхождений после записи: {len(bad)} {bad if bad else ''}")
    print(f"Регистрации обмена (.Изменения): было {changes_before}, стало {changes_after}")
    if bad or changes_after != changes_before:
        sys.exit(1)


if __name__ == "__main__":
    main("--write" in sys.argv)
