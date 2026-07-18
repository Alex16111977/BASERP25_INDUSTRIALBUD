"""
Idempotency: дважды перепровести Ф2 → проверить что UUID ВКассу не меняется
и в базе не появляется второй документ.
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN_ERP)

    q = erp.NewObject("Запрос")
    q.Текст = '''ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка
ИЗ Документ.РаспределениеФ2 КАК Док
ГДЕ Док.Номер = "000000026" И Док.Проведен'''
    rs = q.Выполнить()
    if rs.Пустой():
        print("FAIL: Ф2 №000000026 не найден")
        sys.exit(1)
    sel = rs.Выбрать(); sel.Следующий()
    f2_ref = sel.Ссылка

    # Первое перепроведение
    obj = f2_ref.ПолучитьОбъект()
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    f2_obj = f2_ref.ПолучитьОбъект()
    vk_uid_1 = erp.string(f2_obj.А_ВедомостьВКассу.УникальныйИдентификатор())
    vk_sum_1 = float(f2_obj.А_ВедомостьВКассу.СуммаПоДокументу)
    vk_zarp_1 = f2_obj.А_ВедомостьВКассу.ПолучитьОбъект().Зарплата.Количество()
    print(f"Первый прогон: UID = {vk_uid_1}, Σ = {vk_sum_1:,.2f}, Зарплата = {vk_zarp_1}")

    # Второе перепроведение
    obj2 = f2_ref.ПолучитьОбъект()
    obj2.Записать(erp.РежимЗаписиДокумента.Проведение)
    f2_obj2 = f2_ref.ПолучитьОбъект()
    vk_uid_2 = erp.string(f2_obj2.А_ВедомостьВКассу.УникальныйИдентификатор())
    vk_sum_2 = float(f2_obj2.А_ВедомостьВКассу.СуммаПоДокументу)
    vk_zarp_2 = f2_obj2.А_ВедомостьВКассу.ПолучитьОбъект().Зарплата.Количество()
    print(f"Второй прогон: UID = {vk_uid_2}, Σ = {vk_sum_2:,.2f}, Зарплата = {vk_zarp_2}")

    # Проверки
    if vk_uid_1 != vk_uid_2:
        print(f"FAIL: UUID изменился: {vk_uid_1} → {vk_uid_2}")
        sys.exit(1)
    if abs(vk_sum_1 - vk_sum_2) > 0.01:
        print(f"FAIL: Σ изменилась: {vk_sum_1} → {vk_sum_2}")
        sys.exit(1)
    if vk_zarp_1 != vk_zarp_2:
        print(f"FAIL: Зарплата изменилась: {vk_zarp_1} → {vk_zarp_2}")
        sys.exit(1)
    print("UUID, Σ и количество строк совпали ✓")

    # Проверка отсутствия дублей в базе
    qd = erp.NewObject("Запрос")
    qd.Текст = '''ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол
ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу КАК Вед
ГДЕ Вед.А_РаспределениеФ2 = &Ф2 И НЕ Вед.ПометкаУдаления'''
    qd.УстановитьПараметр("Ф2", f2_ref)
    rs2 = qd.Выполнить()
    sel2 = rs2.Выбрать(); sel2.Следующий()
    cnt = int(sel2.Кол)
    if cnt != 1:
        print(f"FAIL: В базе {cnt} ведомостей по этому Ф2 (должна быть 1)")
        sys.exit(1)
    print(f"В базе ровно 1 ВКассу для этого Ф2 ✓")

    print("\nPASS: idempotency test пройден")


if __name__ == "__main__":
    main()
