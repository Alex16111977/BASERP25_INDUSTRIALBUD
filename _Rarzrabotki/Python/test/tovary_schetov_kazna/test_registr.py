# -*- coding: utf-8 -*-
"""Регистр А_ТоварыСчетовПоставщиков загружен в BaseERP с полями по спеку."""
import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
r = erp.Метаданные.РегистрыСведений.Найти("А_ТоварыСчетовПоставщиков")
ok = r is not None
if ok:
    def names(coll):
        return {coll.Получить(i).Имя for i in range(coll.Количество())}
    need = [{"СчетКазнаУИД", "НомерПозиции"},
            {"Номенклатура", "НазваниеВСчете", "Количество", "Единица", "Цена", "Сумма"},
            {"ЗаказПоставщику", "НомерСчета", "ДатаСчета", "Контрагент", "Подразделение", "СуммаСчета", "ИмяФайла",
             "СпособРаспознавания", "Предупреждение"}]
    got = [names(r.Измерения), names(r.Ресурсы), names(r.Реквизиты)]
    for n, g in zip(need, got):
        if not n <= g:
            print("нет полей:", n - g); ok = False
print("OK" if ok else "FAIL")
sys.exit(0 if ok else 1)
