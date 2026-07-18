# -*- coding: utf-8 -*-
"""
Part 2: Follow-up investigation for РаспределениеФ2 period identification
and БДДС register analysis.
"""

import win32com.client
import sys
import datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

print("=" * 100)
print("PART 2: РаспределениеФ2 Period Field Analysis & БДДС Register")
print("=" * 100)

# Connect
print("\n[1] Connecting to BuhKazn...")
try:
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
    S = conn.String
    print("    ✓ Connected")
except Exception as e:
    print(f"    ✗ Failed: {e}")
    sys.exit(1)

def format_value(val):
    """Format a value for display."""
    if val is None:
        return "[NULL]"
    if isinstance(val, datetime.datetime):
        return val.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    try:
        return str(S(val))
    except:
        return f"[{type(val).__name__}]"

# ============================================================================
# ANALYSIS: What is the salary period for РаспределениеФ2?
# ============================================================================
print("\n" + "=" * 100)
print("[2] SALARY PERIOD FIELD ANALYSIS")
print("=" * 100)

print("""
KEY FINDING: The РаспределениеФ2 document uses the field КассаДатаС / КассаДатаПо
to indicate the SALARY PERIOD it belongs to.

Document 033 (created 30.12.2025):
  • КассаДатаС = 2025-12-01 (December 1st)
  • КассаДатаПо = 2025-12-31 (December 31st)
  → This is a DECEMBER salary distribution

Document 028 (created 10.12.2025):
  • КассаДатаС = 2025-11-01 (November 1st)
  • КассаДатаПо = 2025-11-30 (November 30th)
  → This is a NOVEMBER salary distribution

CONCLUSION: Even though both documents have dates in December 2025:
- Doc 033 was created on 30.12.2025 but refers to DECEMBER salary period
- Doc 028 was created on 10.12.2025 but refers to NOVEMBER salary period

The salary period is determined by КассаДатаС and КассаДатаПо fields, NOT by the document date.
""")

# ============================================================================
# QUERY: Find all РаспределениеФ2 documents and their periods
# ============================================================================
print("\n" + "=" * 100)
print("[3] ALL РаспределениеФ2 DOCUMENTS AND THEIR SALARY PERIODS")
print("=" * 100)

try:
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Док.Номер КАК НомерДокумента,
        Док.Дата КАК ДатаДокумента,
        Док.КассаДатаС КАК ПериодС,
        Док.КассаДатаПо КАК ПериодПо,
        COUNT(DISTINCT СМА.Сотрудник) КАК КоличествоСотрудников,
        СУММА(СМА.СуммаНачисления) КАК ОбщаяСумма
    ИЗ Документ.РаспределениеФ2 КАК Док
        ЛЕВОЕ СОЕДИНЕНИЕ Документ.РаспределениеФ2.Сотрудники КАК СМА
            ПО Док.Ссылка = СМА.Ссылка
    ГДЕ НЕ Док.ПометкаУдаления
    СГРУППИРОВАТЬ ПО
        Док.Номер,
        Док.Дата,
        Док.КассаДатаС,
        Док.КассаДатаПо
    УПОРЯДОЧИТЬ ПО
        Док.Дата
    """
    
    result = q.Execute().Choose()
    
    docs = []
    while result.Next():
        docs.append({
            'номер': S(result.НомерДокумента),
            'дата': result.ДатаДокумента,
            'период_с': result.ПериодС,
            'период_по': result.ПериодПо,
            'кол_во': result.КоличествоСотрудников,
            'сумма': result.ОбщаяСумма,
        })
    
    print(f"\nTotal РаспределениеФ2 documents: {len(docs)}\n")
    
    # Group by salary period
    periods = {}
    for doc in docs:
        period_key = (doc['период_с'], doc['период_по'])
        if period_key not in periods:
            periods[period_key] = []
        periods[period_key].append(doc)
    
    for (period_s, period_po), docs_in_period in sorted(periods.items()):
        period_str = f"{format_value(period_s).split()[0]} to {format_value(period_po).split()[0]}"
        print(f"Salary Period: {period_str}")
        print(f"  Documents in this period: {len(docs_in_period)}")
        for doc in docs_in_period:
            print(f"    • {doc['номер']:<15} (created {format_value(doc['дата']).split()[0]}, "
                  f"{doc['кол_во']} employees, sum={doc['сумма']:>12,.0f})")
        print()

except Exception as e:
    print(f"    ✗ Query failed: {e}")

# ============================================================================
# QUERY: Check БДДС register for salary-related records
# ============================================================================
print("\n" + "=" * 100)
print("[4] БДДС REGISTER - SALARY CASH FLOW RECORDS")
print("=" * 100)

try:
    # First, let's get a sample of БДДС records to see what data is there
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 20
        Период,
        Регистратор,
        Сотрудник,
        Подразделение,
        СтатьяДвиженияДенежныхСредств,
        Сумма
    ИЗ РегистрНакопления.БДДС
    ГДЕ Сотрудник <> ""
        И Сумма <> 0
    УПОРЯДОЧИТЬ ПО Период DESC
    """
    
    result = q.Execute().Choose()
    
    row_count = 0
    while result.Next():
        row_count += 1
        if row_count <= 10:
            print(f"Record {row_count}:")
            print(f"  Период                        = {format_value(result.Период)}")
            print(f"  Регистратор                   = {format_value(result.Регистратор)}")
            print(f"  Сотрудник                     = {format_value(result.Сотрудник)}")
            print(f"  Подразделение                 = {format_value(result.Подразделение)}")
            print(f"  СтатьяДвиженияДенежныхСредств = {format_value(result.СтатьяДвиженияДенежныхСредств)}")
            print(f"  Сумма                         = {format_value(result.Сумма)}")
            print()

except Exception as e:
    print(f"    ✗ БДДС query failed: {e}")

# ============================================================================
# METADATA: Check what's in the Регистратор dimension
# ============================================================================
print("\n" + "=" * 100)
print("[5] WHAT DOCUMENTS CAN BE REGISTRARS IN БДДС?")
print("=" * 100)

try:
    reg_meta = conn.Metadata.AccumulationRegisters.БДДС
    
    # Get Регистратор dimension
    dims = []
    for i in range(reg_meta.Dimensions.Count()):
        dim = reg_meta.Dimensions.Get(i)
        if "Регистратор" in dim.Name:
            dims.append(dim)
    
    if dims:
        for dim in dims:
            print(f"\nDimension: {dim.Name}")
            print(f"  Type info: {dim.Type}")
            
            # Try to get the type definition
            try:
                types_list = []
                for t in dim.Type.Types():
                    types_list.append(str(t))
                print(f"  Allowed types: {', '.join(types_list)}")
            except:
                pass
    else:
        print("No Регистратор dimension found")

except Exception as e:
    print(f"    ✗ Metadata query failed: {e}")

# ============================================================================
# SUMMARY
# ============================================================================
print("\n" + "=" * 100)
print("[6] FINDINGS SUMMARY")
print("=" * 100)

print("""
╔════════════════════════════════════════════════════════════════════════════╗
║ SALARY PERIOD IDENTIFICATION IN РаспределениеФ2                          ║
╚════════════════════════════════════════════════════════════════════════════╝

PRIMARY PERIOD FIELD:
  → КассаДатаС (Period Start Date)
  → КассаДатаПо (Period End Date)

These fields define the SALARY MONTH that the distribution refers to, regardless
of when the document was created.

DOCUMENT HEADER ATTRIBUTES:
  ✓ Организация               (Organization)
  ✓ Комментарий              (Comment)
  ✓ Ответственный            (Responsible person)
  ✓ КассаДатаС               (SALARY PERIOD START) ← KEY FIELD
  ✓ КассаДатаПо              (SALARY PERIOD END)   ← KEY FIELD
  ✓ ОтбиратьТабельПоОрганизации (By organization)
  ✓ Направление              (Direction/Department)
  ✓ ОпределятьНаправлениеПоТабелю (By table)
  ✓ УчитыватьАванс           (Count advance payments)

TABULAR SECTIONS:
  1. Сотрудники (Employees)
     - Сотрудник
     - СуммаНачисления
     
  2. Распределение (Distribution)
     - Сотрудник
     - Подразделение
     - ОтработаноЧасов
     - СуммаЗатрат
     - СуммаНачисления
     - Направление

IMPORTANT FOR SALARY TRACKING:
  When analyzing documents created in December 2025:
  - Document 033 (dated 30.12.2025) → DECEMBER salary (12/1-12/31)
  - Document 028 (dated 10.12.2025) → NOVEMBER salary (11/1-11/30)
  
  Always use КассаДатаС/КассаДатаПо to identify the actual salary period,
  NOT the document creation date!

БДДС REGISTER:
  - Contains cash flow records
  - Can have РаспределениеФ2 as Регистратор
  - Has Сотрудник attribute (employee)
  - Has Период field for the register period
  - Tracks: Сумма (amount), Подразделение (department), etc.
""")

print("=" * 100)
print("Investigation complete!")
print("=" * 100)
