"""Прогон оновленого СформироватьЗапросСверткиPL — тест на синтаксис + перевірка коментаря."""
import sys
import re
from datetime import datetime
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
BSL_PATH = r"C:\Configuration_downloads\BASERP25\Documents\А_ФинРез_PL\Ext\ObjectModule.bsl"


def extract_query():
    with open(BSL_PATH, encoding='utf-8') as f:
        src = f.read()
    # Витягую тіло запиту з функції
    m = re.search(r'Функция СформироватьЗапросСверткиPL\(\)\s*Возврат\s*"(.*?)";', src, re.DOTALL)
    if not m:
        return None
    raw = m.group(1)
    # Прибираю BSL-префікс |\t для багаторядкового рядка та екранування ""
    lines = raw.split('\n')
    cleaned = []
    for line in lines:
        if line.startswith('\t|'):
            cleaned.append(line[2:])
        elif line.startswith('|'):
            cleaned.append(line[1:])
        else:
            cleaned.append(line)
    return '\n'.join(cleaned).replace('""', '"')


def main():
    q_text = extract_query()
    if not q_text:
        print("FAIL: query not extracted")
        return 1

    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)

    org_q = erp.NewObject("Запрос")
    org_q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Организации ГДЕ КодПоЕДРПОУ = \"40645273\""
    org = org_q.Execute().Выгрузить().Получить(0).Ссылка

    dds_q = erp.NewObject("Запрос")
    dds_q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.СтатьиДвиженияДенежныхСредств ГДЕ А_ПриёмникСебестоимостиПродажPL И НЕ ПометкаУдаления"
    dds_rows = dds_q.Execute().Выгрузить()
    dds = dds_rows.Получить(0).Ссылка if dds_rows.Количество() else None

    print(f"Орг: {org}, ДДСCoGS: {dds}")
    print(f"Query length: {len(q_text)} chars")

    q = erp.NewObject("Запрос")
    q.Text = q_text
    q.УстановитьПараметр("НачалоПериода", datetime(2026, 2, 1))
    q.УстановитьПараметр("КонецПериода", datetime(2026, 2, 28, 23, 59, 59))
    q.УстановитьПараметр("Организация", org)
    q.УстановитьПараметр("ДДСCoGS", dds)

    try:
        result = q.Выполнить().Выгрузить()
        cnt = result.Количество()
        print(f"OK: rows={cnt}")
        with_comment = 0
        for i in range(cnt):
            row = result.Получить(i)
            if row.Комментарий and row.Комментарий != "":
                with_comment += 1
                if with_comment <= 5:
                    src = str(row.Источник)
                    print(f"  [{src}] {row.Комментарий[:80]}")
        print(f"\nWith Комментарий: {with_comment} / {cnt}")
    except Exception as e:
        err = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        print(f"FAIL: {err}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
