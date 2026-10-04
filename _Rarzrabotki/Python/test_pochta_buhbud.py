# Проверка системной учётной записи почты BuhBud — как кнопка «Проверить настройки» (БСП):
# подключение к SMTP по профилю учётки и отключение. Письмо НЕ отправляется, пароль не читается и не выводится.
# Код выхода: 0 — подключение успешно, 1 — ошибка подключения, 2 — пароль в хранилище не записан.
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'

Q_STORAGE = """
ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол
ИЗ РегистрСведений.БезопасноеХранилищеДанных КАК БХ
ГДЕ БХ.Владелец = &Владелец
"""


def main():
    buh = win32com.client.Dispatch("V83.COMConnector").Connect(CONN_BUH)
    buh.УстановитьПривилегированныйРежим(True)
    ref = buh.Справочники.УчетныеЗаписиЭлектроннойПочты.СистемнаяУчетнаяЗаписьЭлектроннойПочты
    print(f"Учётка: {ref.АдресЭлектроннойПочты}, SMTP {ref.СерверИсходящейПочты}:{ref.ПортСервераИсходящейПочты}, "
          f"для отправки: {buh.String(ref.ИспользоватьДляОтправки)}")

    q = buh.NewObject("Запрос")
    q.Text = Q_STORAGE
    q.SetParameter("Владелец", ref)
    sel = q.Execute().Выбрать()
    sel.Следующий()
    if sel.Кол == 0:
        print("Пароль в безопасном хранилище BuhBud не записан — тест SMTP пропущен.")
        sys.exit(2)
    print("Пароль в хранилище: есть")

    profile = buh.РаботаСПочтовымиСообщениямиСлужебный.ИнтернетПочтовыйПрофиль(ref, False)
    mail = buh.NewObject("ИнтернетПочта")
    try:
        mail.Подключиться(profile)
        mail.Отключиться()
    except Exception as e:
        msg = e.excepinfo[2] if getattr(e, 'excepinfo', None) else e
        print(f"SMTP: ОШИБКА подключения — {msg}")
        sys.exit(1)
    print("SMTP: подключение и авторизация успешны")


if __name__ == "__main__":
    main()
