"""
Тест: реквізит А_НомерДокументаИзБух присутній у 24 документах ERP
(22 нових з deploy 2026-04-25 + 2 еталонних які вже мали).

Перевіряє через COM запит "ВЫБРАТЬ ПЕРВЫЕ 1 А_НомерДокументаИзБух ИЗ Документ.<DocName>".
Якщо реквізит у БД є — запит виконається без помилки.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from utils.com_connect import connect_erp  # noqa: E402

NEW_DOCS = [
    "ПриходныйКассовыйОрдер", "ВозвратТоваровПоставщику", "ЗаказКлиента",
    "ЗаказПоставщику", "ПеремещениеТоваров", "ПоступлениеБезналичныхДенежныхСредств",
    "СписаниеБезналичныхДенежныхСредств", "АвансовыйОтчет", "РасходныйКассовыйОрдер",
    "РеализацияТоваровУслуг", "ПриобретениеТоваровУслуг", "СборкаТоваров",
    "АктВыполненныхРабот", "ПокупкаПродажаВалюты", "ЗаявкаНаПокупкуПродажуВалюты",
    "ВнутреннееПотреблениеТоваров", "ПриобретениеУслугПрочихАктивов",
    "ВзаимозачетЗадолженности", "А_ОтражениеЗарплатыВУчете",
    "ПередачаМатериаловВПроизводство", "ПринятиеКУчетуОС", "ПринятиеКУчетуНМА",
]
EXISTING_DOCS = ["ВедомостьНаВыплатуЗарплатыВБанк", "СписаниеНедостачТоваров"]


def check(conn, doc):
    q = conn.NewObject("Запрос")
    q.Текст = f"ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, А_НомерДокументаИзБух ИЗ Документ.{doc}"
    try:
        q.Выполнить()
        return True, ""
    except Exception as e:
        return False, str(e)[:200]


def main():
    erp = connect_erp()
    fails = []
    print(f"Перевіряю {len(NEW_DOCS)} нових + {len(EXISTING_DOCS)} еталонних = {len(NEW_DOCS)+len(EXISTING_DOCS)} документів...")
    print()
    for doc in NEW_DOCS + EXISTING_DOCS:
        ok, err = check(erp, doc)
        marker = "OK  " if ok else "FAIL"
        print(f"  [{marker}] {doc}")
        if not ok:
            fails.append((doc, err))
    print()
    if fails:
        print(f"FAIL: {len(fails)} з {len(NEW_DOCS)+len(EXISTING_DOCS)} документів не мають реквізиту:")
        for d, e in fails:
            print(f"  {d}: {e}")
        sys.exit(1)
    print(f"ALL OK: {len(NEW_DOCS)+len(EXISTING_DOCS)} документів мають реквізит А_НомерДокументаИзБух")


if __name__ == "__main__":
    main()
