# -*- coding: utf-8 -*-
"""Відкат ВСІХ АмортизацияОС до стандартного стану.
Видаляє всі РАСХОД записи ВНА з кожного документа.
"""
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Знайти всі проведені АмортизацияОС
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Д.Ссылка ИЗ Документ.АмортизацияОС КАК Д
ГДЕ Д.Проведен = ИСТИНА УПОРЯДОЧИТЬ ПО Д.Дата"""
docs = q.Execute().Unload()

for d in range(docs.Count()):
    doc_ref = docs.Get(d).Ссылка
    print(f"\nДок: {S(doc_ref)}")

    nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
    nabor.Filter.Регистратор.Set(doc_ref)
    nabor.Read()
    original_count = nabor.Count()

    # Видалити РАСХОД записи (ВидДвижения = Расход)
    i = nabor.Count() - 1
    deleted = 0
    while i >= 0:
        rec = nabor.Get(i)
        vd_xml = conn.XMLString(rec.RecordType)
        if vd_xml == "Expense":
            nabor.Delete(i)
            deleted += 1
        i -= 1

    if deleted > 0:
        print(f"  {original_count} -> {nabor.Count()} (видалено {deleted} РАСХОД)")
        nabor.Write(True)
        print(f"  Записано!")
    else:
        print(f"  Немає РАСХОД записів ({original_count} записів)")

print("\nВідкат завершено!")
