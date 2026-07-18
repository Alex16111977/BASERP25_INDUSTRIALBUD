"""
Smoke: перепровести РаспределениеФ2 №000000026 через COM →
проверить что создалась ВКассу с правильным А_РаспределениеФ2 и Σ=348800.

Запуск: C:\\Python313\\python.exe test_create_vedomost_vkassu_iz_f2_smoke.py
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EXPECTED_SUM = 348800.00
EXPECTED_ZARP_ROWS = 31   # число строк ТЧ Распределение Ф2 №000000026 (= ТЧ Зарплата ВКассу hotfix v2)
EXPECTED_UNIQ_FL = 21     # число уникальных ФЛ (= ТЧ Состав ВКассу)
EXPECTED_SPOSOB = "Зарплата за месяц"
EXPECTED_PROCENT = 100


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN_ERP)

    # 1. Найти эталон Ф2
    q = erp.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1
    Док.Ссылка КАК Ссылка
ИЗ Документ.РаспределениеФ2 КАК Док
ГДЕ Док.Номер = "000000026" И Док.Проведен
"""
    rs = q.Выполнить()
    if rs.Пустой():
        print("FAIL: РаспределениеФ2 №000000026 не найден")
        sys.exit(1)

    sel = rs.Выбрать(); sel.Следующий()
    f2_ref = sel.Ссылка
    print(f"Ф2: {f2_ref}")

    # 2. Перепровести Ф2 → ПриЗаписи триггерит создание ВКассу
    obj = f2_ref.ПолучитьОбъект()
    try:
        obj.Записать(erp.РежимЗаписиДокумента.Проведение)
        print("OK: РаспределениеФ2 перепроведено → ПриЗаписи отработал")
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"FAIL: перепроведение Ф2: {info}")
        sys.exit(1)

    # 3. Перечитать Ф2 (обратная ссылка должна обновиться)
    f2_re = erp.Документы.РаспределениеФ2.ПолучитьСсылку(f2_ref.УникальныйИдентификатор())
    f2_obj = f2_re.ПолучитьОбъект()
    if not f2_obj.А_ВедомостьВКассу or str(f2_obj.А_ВедомостьВКассу) == "":
        print("FAIL: Ф2.А_ВедомостьВКассу не заполнена после перепроведения")
        sys.exit(1)

    vk_ref = f2_obj.А_ВедомостьВКассу
    print(f"ВКассу: {vk_ref}")

    # 4. Проверить ВКассу
    vk = vk_ref.ПолучитьОбъект()

    # Шапка: Подразделение = ПУСТО (hotfix v2)
    if not vk.Подразделение.Пустая():
        podr_name = str(vk.Подразделение.Наименование)
        print(f"FAIL: Подразделение должно быть ПУСТО, а получили '{podr_name}'")
        sys.exit(1)
    print(f"  Шапка.Подразделение = ПУСТО ✓")

    # Шапка: СпособВыплаты = предопределённый "Зарплата" (Наим="Зарплата за месяц")
    sposob_name = str(vk.СпособВыплаты.Наименование) if not vk.СпособВыплаты.Пустая() else ""
    if sposob_name != EXPECTED_SPOSOB:
        print(f"FAIL: СпособВыплаты должен быть '{EXPECTED_SPOSOB}', получили '{sposob_name}'")
        sys.exit(1)
    print(f"  Шапка.СпособВыплаты = '{sposob_name}' ✓")

    # Шапка: ПроцентВыплаты = 100
    if int(vk.ПроцентВыплаты) != EXPECTED_PROCENT:
        print(f"FAIL: ПроцентВыплаты = {vk.ПроцентВыплаты}, ожидалось {EXPECTED_PROCENT}")
        sys.exit(1)
    print(f"  Шапка.ПроцентВыплаты = {vk.ПроцентВыплаты} ✓")

    # Шапка: Организация = Касса.Владелец (если касса найдена)
    if not vk.Касса.Пустая():
        org_uid_vk = erp.string(vk.Организация.УникальныйИдентификатор())
        org_uid_kassa = erp.string(vk.Касса.Владелец.УникальныйИдентификатор())
        if org_uid_vk != org_uid_kassa:
            print(f"FAIL: Организация ({vk.Организация.Наименование}) != Касса.Владелец ({vk.Касса.Владелец.Наименование})")
            sys.exit(1)
        print(f"  Шапка.Организация = Касса.Владелец = '{vk.Организация.Наименование}' ✓")
    else:
        print("  WARN: Касса не определена — проверку Организация=Касса.Владелец пропускаем")

    # Шапка: А_РаспределениеФ2 ссылка корректная
    vk_f2_uid = erp.string(vk.А_РаспределениеФ2.УникальныйИдентификатор())
    f2_uid = erp.string(f2_ref.УникальныйИдентификатор())
    if vk_f2_uid != f2_uid:
        print(f"FAIL: А_РаспределениеФ2.UID не совпадает: {vk_f2_uid} != {f2_uid}")
        sys.exit(1)
    print(f"  Шапка.А_РаспределениеФ2.UID = {vk_f2_uid} ✓")

    # ТЧ Зарплата = 31 строк (= ТЧ Распределение Ф2), Σ = 348800
    cnt_zp = vk.Зарплата.Количество()
    sum_zp = 0
    for i in range(cnt_zp):
        sum_zp += float(vk.Зарплата.Получить(i).КВыплате)
    print(f"  Зарплата: {cnt_zp} строк, Σ={sum_zp:,.2f}")
    if cnt_zp != EXPECTED_ZARP_ROWS:
        print(f"FAIL: ожидалось {EXPECTED_ZARP_ROWS} строк Зарплата, получили {cnt_zp}")
        sys.exit(1)
    if abs(sum_zp - EXPECTED_SUM) > 0.01:
        print(f"FAIL: Σ Зарплата = {sum_zp}, ожидалось {EXPECTED_SUM}")
        sys.exit(1)

    # ТЧ Зарплата: ГруппаУчётаНачислений = "Зарплата (661)" во всех строках
    no_group = 0
    for i in range(cnt_zp):
        r = vk.Зарплата.Получить(i)
        if r.ГруппаУчетаНачислений.Пустая():
            no_group += 1
    if no_group > 0:
        print(f"FAIL: {no_group} строк ТЧ Зарплата без ГруппыУчётаНачислений")
        sys.exit(1)
    print(f"  Зарплата.ГруппаУчётаНачислений заполнена во всех строках ✓")

    # ТЧ Расшифровка
    cnt_rs = vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()
    print(f"  А_Расшифровка: {cnt_rs} строк")
    if cnt_rs == 0:
        print("FAIL: ТЧ А_Расшифровка пуста")
        sys.exit(1)

    # ТЧ Состав = 21 уникальных ФЛ
    cnt_sost = vk.Состав.Количество()
    print(f"  Состав: {cnt_sost} строк (уникальные ФЛ)")
    if cnt_sost != EXPECTED_UNIQ_FL:
        print(f"FAIL: Состав должен иметь {EXPECTED_UNIQ_FL} ФЛ, а получили {cnt_sost}")
        sys.exit(1)

    # СуммаПоДокументу
    if abs(float(vk.СуммаПоДокументу) - EXPECTED_SUM) > 0.01:
        print(f"FAIL: СуммаПоДокументу = {vk.СуммаПоДокументу}, ожидалось {EXPECTED_SUM}")
        sys.exit(1)
    print(f"  СуммаПоДокументу = {vk.СуммаПоДокументу} ✓")

    # ВКассу проводится автоматически
    if not vk.Проведен:
        print(f"FAIL: документ должен быть Проведён=True, а Проведен={vk.Проведен}")
        sys.exit(1)
    print(f"  Проведена = True ✓")

    print(f"\nPASS: smoke test пройден")


if __name__ == "__main__":
    main()
