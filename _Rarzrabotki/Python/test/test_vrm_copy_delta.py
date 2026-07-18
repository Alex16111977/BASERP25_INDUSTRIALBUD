# -*- coding: utf-8 -*-
"""Сценарий пользователя: копия последнего ВРМ + доввод процента -> проведение пишет ДЕЛЬТУ.

Get-or-create по маркеру в Комментарии (документы НЕ удаляются, в конце — отмена
проведения + пометка удаления, обратимо). Проверки:
  1. Программная копия записывается и проводится без ошибок.
  2. Обороты РН по регистратору-копии = только дельта довведённого процента.
  3. Перепроведение не дублирует обороты.
"""
import sys
from datetime import timedelta

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MARKER = "ТЕСТ CLAUDE: копия с довводом процентов (можно удалить)"
BUMP = 10.0  # довводимый процент

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

fails = []


def check(label, ok, detail=""):
    print(("OK  " if ok else "FAIL") + f" {label}" + (f" | {detail}" if detail else ""))
    if not ok:
        fails.append(label)


# --- 1. Последний проведённый документ (не тестовый) ---
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1
    Док.Ссылка КАК Ссылка
ИЗ Документ.А_ВыполнениеРаботМодулей КАК Док
ГДЕ Док.Проведен И НЕ Док.ПометкаУдаления И ВЫРАЗИТЬ(Док.Комментарий КАК Строка(200)) <> &Маркер
УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ
"""
q.SetParameter("Маркер", MARKER)
r = q.Execute().Выгрузить()
if r.Количество() == 0:
    print("SKIP: нет проведённых документов ВРМ")
    sys.exit(1)
src_ref = r.Получить(0).Ссылка
src = src_ref.ПолучитьОбъект()
n_mod = src.КоличествоМодулей
print(f"Исходник: {src_ref.Номер} от {src_ref.Дата}, модулей={n_mod}, строк ТЧ={src.Работы.Количество()}")

# --- 2. Get-or-create тестовой копии по маркеру ---
q2 = erp.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка
ИЗ Документ.А_ВыполнениеРаботМодулей КАК Док
ГДЕ ВЫРАЗИТЬ(Док.Комментарий КАК Строка(200)) = &Маркер
"""
q2.SetParameter("Маркер", MARKER)
r2 = q2.Execute().Выгрузить()
if r2.Количество() > 0:
    copy_obj = r2.Получить(0).Ссылка.ПолучитьОбъект()
    if copy_obj.Проведен:
        copy_obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        copy_obj = r2.Получить(0).Ссылка.ПолучитьОбъект()
    copy_obj.ПометкаУдаления = False
    copy_obj.Работы.Очистить()
    print("Тестовый документ переиспользован")
else:
    copy_obj = erp.Документы.А_ВыполнениеРаботМодулей.СоздатьДокумент()
    print("Тестовый документ создан")

# --- 3. «Копия»: шапка + ТЧ 1:1, затем доввод процента ---
copy_obj.Дата = src.Дата + timedelta(days=1)
copy_obj.Организация = src.Организация
copy_obj.СтруктураСебестоимости = src.СтруктураСебестоимости
copy_obj.Подразделение = src.Подразделение
copy_obj.КоличествоМодулей = n_mod
copy_obj.Ответственный = src.Ответственный
copy_obj.Комментарий = MARKER

bump_info = None
for i in range(src.Работы.Количество()):
    s = src.Работы.Получить(i)
    d = copy_obj.Работы.Добавить()
    d.Работа = s.Работа
    d.Этап = s.Этап
    d.НомерМодуля = s.НомерМодуля
    d.Процент = s.Процент
    d.ПланЧасыМодуля = s.ПланЧасыМодуля
    d.ПроцентНаНачало = s.ПроцентНаНачало
    d.ЧасыВыполнение = s.ЧасыВыполнение
    d.Комментарий = s.Комментарий
    if bump_info is None and s.Процент <= 100 - BUMP and s.ПланЧасыМодуля > 0:
        d.Процент = s.Процент + BUMP
        d.ЧасыВыполнение = d.ПланЧасыМодуля * d.Процент / 100
        bump_info = (str(s.Работа), s.НомерМодуля, s.ПланЧасыМодуля)

check("копия шапки/ТЧ собрана", copy_obj.Работы.Количество() == src.Работы.Количество(),
      f"строк={copy_obj.Работы.Количество()}")
if bump_info is None:
    print("SKIP: не нашлась клетка для доввода процента (всё >= 90%)")
    sys.exit(1)
expected_delta = bump_info[2] * BUMP / 100
print(f"Доввод: работа='{bump_info[0]}', модуль М{bump_info[1]}, план_модуля={bump_info[2]}, ожидаемая дельта={expected_delta:.3f} ч")

# --- 4. Проведение ---
try:
    copy_obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    check("запись и проведение копии", True)
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
    check("запись и проведение копии", False, msg[:300])
    sys.exit(1)

copy_ref = copy_obj.Ссылка


def oborot_by_registrator(ref):
    qq = erp.NewObject("Запрос")
    qq.Text = """
    ВЫБРАТЬ ЕСТЬNULL(СУММА(РН.ЧасыВыполнение), 0) КАК Часы, КОЛИЧЕСТВО(*) КАК Строк
    ИЗ РегистрНакопления.А_ВыполнениеРаботМодулей КАК РН
    ГДЕ РН.Регистратор = &Док
    """
    qq.SetParameter("Док", ref)
    rr = qq.Execute().Выгрузить()
    return float(rr.Получить(0).Часы), int(rr.Получить(0).Строк)


delta, rows = oborot_by_registrator(copy_ref)
check("обороты копии = дельта доввода", abs(delta - expected_delta) < 0.01,
      f"факт={delta:.3f}, ожидание={expected_delta:.3f}, движений={rows}")

# --- 5. Перепроведение без дублей ---
copy_obj = copy_ref.ПолучитьОбъект()
copy_obj.Записать(erp.РежимЗаписиДокумента.Проведение)
delta2, rows2 = oborot_by_registrator(copy_ref)
check("перепроведение без дублей", abs(delta2 - delta) < 0.001 and rows2 == rows,
      f"факт={delta2:.3f}, движений={rows2}")

# --- 6. Уборка: отмена проведения + пометка удаления (обратимо) ---
copy_obj = copy_ref.ПолучитьОбъект()
copy_obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
copy_obj = copy_ref.ПолучитьОбъект()
copy_obj.ПометкаУдаления = True
copy_obj.Записать()
delta3, rows3 = oborot_by_registrator(copy_ref)
check("уборка: движений нет, помечен на удаление", rows3 == 0)

print("\n" + ("TEST PASSED" if not fails else f"TEST FAILED: {fails}"))
sys.exit(0 if not fails else 1)
