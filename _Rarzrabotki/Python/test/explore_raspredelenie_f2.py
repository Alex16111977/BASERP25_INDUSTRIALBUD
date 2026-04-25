# -*- coding: utf-8 -*-
"""
Investigation script for РаспределениеФ2 document structure in BuhKazn.

Objectives:
1. Get metadata structure of Document.РаспределениеФ2
2. Query specific example documents (033 and 028)
3. Check БДДС register structure
4. Identify salary period fields
"""

import win32com.client
import sys
import datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

print("=" * 100)
print("INVESTIGATION: РаспределениеФ2 Document Structure in BuhKazn")
print("=" * 100)

# ============================================================================
# CONNECT TO BuhKazn
# ============================================================================
print("\n[1] CONNECTING TO BuhKazn...")
try:
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
    S = conn.String
    print("    ✓ Connected to BuhKazn")
except Exception as e:
    print(f"    ✗ Connection failed: {e}")
    sys.exit(1)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def iterate_collection(collection):
    """Safely iterate over a 1C collection."""
    items = []
    try:
        for item in collection:
            items.append(item)
    except TypeError:
        try:
            count = collection.Count()
            for i in range(count):
                items.append(collection.Get(i))
        except:
            pass
    return items

def get_type_desc(type_obj):
    """Get human-readable type description."""
    try:
        types_list = []
        for t in iterate_collection(type_obj.Types()):
            try:
                types_list.append(str(t))
            except:
                types_list.append("?")
        return ", ".join(types_list) if types_list else "?"
    except:
        return "?"

def safe_get_field(obj, field_name):
    """Safely get field value from object."""
    try:
        val = getattr(obj, field_name)
        if val is None:
            return None
        return val
    except:
        return None

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
# PART 1: METADATA STRUCTURE
# ============================================================================
print("\n" + "=" * 100)
print("[2] DOCUMENT METADATA: Document.РаспределениеФ2")
print("=" * 100)

try:
    meta = conn.Metadata.Documents.РаспределениеФ2
    
    # Standard attributes
    print("\n--- STANDARD ATTRIBUTES ---")
    std_attrs = iterate_collection(meta.StandardAttributes)
    print(f"Total: {len(std_attrs)}")
    for attr in std_attrs:
        try:
            print(f"  • {attr.Name}")
        except:
            pass
    
    # Header attributes (Реквизиты)
    print("\n--- HEADER ATTRIBUTES (Реквизиты) ---")
    attrs = iterate_collection(meta.Attributes)
    print(f"Total: {len(attrs)}")
    for attr in attrs:
        try:
            attr_name = attr.Name
            type_desc = get_type_desc(attr.Type)
            print(f"  • {attr_name:<40} : {type_desc}")
        except Exception as e:
            print(f"  • [ERROR: {e}]")
    
    # Tabular sections
    print("\n--- TABULAR SECTIONS ---")
    tab_sections = iterate_collection(meta.TabularSections)
    print(f"Total: {len(tab_sections)}")
    
    for ts in tab_sections:
        try:
            ts_name = ts.Name
            ts_attrs = iterate_collection(ts.Attributes)
            print(f"\n  Section: {ts_name} ({len(ts_attrs)} columns)")
            print(f"  {'-' * 80}")
            for ts_attr in ts_attrs:
                try:
                    attr_name = ts_attr.Name
                    type_desc = get_type_desc(ts_attr.Type)
                    print(f"    • {attr_name:<40} : {type_desc}")
                except:
                    pass
        except Exception as e:
            print(f"  [ERROR: {e}]")
    
except Exception as e:
    print(f"    ✗ Metadata query failed: {e}")

# ============================================================================
# PART 2: FIND AND DUMP EXAMPLE DOCUMENTS
# ============================================================================
print("\n" + "=" * 100)
print("[3] QUERYING EXAMPLE DOCUMENTS")
print("=" * 100)

def find_and_dump_doc(doc_number, doc_date_str):
    """Find document by number and date, then dump all its fields."""
    print(f"\n--- Document: Распределение ф2 {doc_number} от {doc_date_str} ---")
    
    try:
        # Parse date
        try:
            doc_date = datetime.datetime.strptime(doc_date_str, "%d.%m.%Y")
        except:
            print(f"  ✗ Invalid date format: {doc_date_str}")
            return
        
        # Query for document
        q = conn.NewObject("Запрос")
        q.Text = """
        ВЫБРАТЬ
            Док.Ссылка
        ИЗ Документ.РаспределениеФ2 КАК Док
        ГДЕ Док.Номер = &ДокНомер
            И Док.Дата МЕЖДУ &ДатаОт И &ДатаПо
            И НЕ Док.ПометкаУдаления
        """
        q.SetParameter("ДокНомер", doc_number)
        q.SetParameter("ДатаОт", doc_date)
        q.SetParameter("ДатаПо", datetime.datetime(doc_date.year, doc_date.month, doc_date.day, 23, 59, 59))
        
        result = q.Execute().Choose()
        
        if not result.Next():
            print(f"  ✗ Document not found: {doc_number}")
            return
        
        doc_ref = result.Ссылка
        
        # Get the document object
        doc_obj = doc_ref.GetObject()
        
        print(f"  ✓ Document found: {format_value(doc_obj.Ссылка)}")
        
        # Dump header fields
        print(f"\n  HEADER FIELDS:")
        print(f"  {'-' * 80}")
        
        meta = conn.Metadata.Documents.РаспределениеФ2
        attrs = iterate_collection(meta.Attributes)
        
        for attr in attrs:
            try:
                field_name = attr.Name
                field_val = safe_get_field(doc_obj, field_name)
                formatted = format_value(field_val)
                print(f"    {field_name:<40} = {formatted}")
            except Exception as e:
                print(f"    {field_name:<40} = [ERROR: {e}]")
        
        # Dump tabular sections
        tab_sections = iterate_collection(meta.TabularSections)
        if tab_sections:
            print(f"\n  TABULAR SECTIONS:")
            for ts in tab_sections:
                try:
                    ts_name = ts.Name
                    ts_table = safe_get_field(doc_obj, ts_name)
                    if ts_table is None:
                        print(f"    {ts_name}: [NULL]")
                        continue
                    
                    row_count = ts_table.Count()
                    print(f"    {ts_name}: {row_count} rows")
                    
                    if row_count > 0:
                        ts_attrs = iterate_collection(ts.Attributes)
                        print(f"    {'-' * 76}")
                        
                        # Print first 3 rows only
                        for row_idx in range(min(3, row_count)):
                            row = ts_table.Get(row_idx)
                            print(f"      Row {row_idx + 1}:")
                            for ts_attr in ts_attrs:
                                try:
                                    field_name = ts_attr.Name
                                    field_val = safe_get_field(row, field_name)
                                    formatted = format_value(field_val)
                                    print(f"        {field_name:<38} = {formatted}")
                                except:
                                    pass
                        
                        if row_count > 3:
                            print(f"      ... and {row_count - 3} more rows")
                
                except Exception as e:
                    print(f"    {ts_name}: [ERROR: {e}]")
    
    except Exception as e:
        print(f"  ✗ Error: {e}")

# Find documents
find_and_dump_doc("000000033", "30.12.2025")
find_and_dump_doc("000000028", "10.12.2025")

# ============================================================================
# PART 3: CHECK БДДС REGISTER FOR РаспределениеФ2 RECORDS
# ============================================================================
print("\n" + "=" * 100)
print("[4] БДДС (CASH FLOW) REGISTER STRUCTURE")
print("=" * 100)

try:
    reg_meta = conn.Metadata.AccumulationRegisters.БДДС
    
    # Dimensions
    print("\n--- DIMENSIONS ---")
    dimensions = iterate_collection(reg_meta.Dimensions)
    print(f"Total: {len(dimensions)}")
    for dim in dimensions:
        try:
            dim_name = dim.Name
            type_desc = get_type_desc(dim.Type)
            print(f"  • {dim_name:<40} : {type_desc}")
        except:
            pass
    
    # Resources
    print("\n--- RESOURCES ---")
    resources = iterate_collection(reg_meta.Resources)
    print(f"Total: {len(resources)}")
    for res in resources:
        try:
            res_name = res.Name
            type_desc = get_type_desc(res.Type)
            print(f"  • {res_name:<40} : {type_desc}")
        except:
            pass
    
    # Attributes
    print("\n--- ATTRIBUTES ---")
    attrs = iterate_collection(reg_meta.Attributes)
    print(f"Total: {len(attrs)}")
    for attr in attrs:
        try:
            attr_name = attr.Name
            type_desc = get_type_desc(attr.Type)
            print(f"  • {attr_name:<40} : {type_desc}")
        except:
            pass

except Exception as e:
    print(f"    ✗ БДДС metadata query failed: {e}")

# ============================================================================
# PART 4: QUERY БДДС RECORDS FOR РаспределениеФ2
# ============================================================================
print("\n" + "=" * 100)
print("[5] БДДС RECORDS WITH REGISTRATOR = РаспределениеФ2")
print("=" * 100)

try:
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 10
        Период,
        Регистратор,
        Сотрудник,
        СМА,
        СчетДебет,
        СчетКредит,
        Аналитика1,
        Аналитика2,
        Аналитика3,
        Сумма,
        ПериодНачисления,
        КодЗарплатыВКазне,
        КодМесяца,
        КодПлатежа
    ИЗ РегистрНакопления.БДДС
    ГДЕ Регистратор.Метаданные().ПолноеИмя = "Документ.РаспределениеФ2"
    """
    
    result = q.Execute().Choose()
    
    row_count = 0
    period_values = set()
    period_nach_values = set()
    kod_mesyaca_values = set()
    
    while result.Next():
        row_count += 1
        
        if row_count <= 5:
            print(f"\n  Record {row_count}:")
            print(f"    Период                   = {format_value(result.Период)}")
            print(f"    Регистратор              = {format_value(result.Регистратор)}")
            print(f"    Сотрудник                = {format_value(result.Сотрудник)}")
            print(f"    СМА                      = {format_value(result.СМА)}")
            print(f"    СчетДебет                = {format_value(result.СчетДебет)}")
            print(f"    СчетКредит               = {format_value(result.СчетКредит)}")
            print(f"    Сумма                    = {format_value(result.Сумма)}")
            try:
                print(f"    ПериодНачисления         = {format_value(result.ПериодНачисления)}")
            except:
                pass
            try:
                print(f"    КодЗарплатыВКазне       = {format_value(result.КодЗарплатыВКазне)}")
            except:
                pass
            try:
                print(f"    КодМесяца                = {format_value(result.КодМесяца)}")
            except:
                pass
            try:
                print(f"    КодПлатежа               = {format_value(result.КодПлатежа)}")
            except:
                pass
        
        # Collect values
        try:
            period_values.add(format_value(result.Период))
        except:
            pass
        try:
            period_nach_values.add(format_value(result.ПериодНачисления))
        except:
            pass
        try:
            kod_mesyaca_values.add(format_value(result.КодМесяца))
        except:
            pass
    
    print(f"\n  Total БДДС records with РаспределениеФ2 registrator: {row_count}")
    
    if period_values:
        print(f"\n  Unique Period values found: {sorted(period_values)}")
    if period_nach_values:
        print(f"  Unique ПериодНачисления values: {sorted(period_nach_values)}")
    if kod_mesyaca_values:
        print(f"  Unique КодМесяца values: {sorted(kod_mesyaca_values)}")

except Exception as e:
    print(f"    ✗ БДДС query failed: {e}")

# ============================================================================
# SUMMARY
# ============================================================================
print("\n" + "=" * 100)
print("[6] SUMMARY & FINDINGS")
print("=" * 100)

print("""
Key observations to identify:
1. Which field in РаспределениеФ2 indicates the SALARY PERIOD?
   - Could be: ПериодНачисления, МесяцЗарплаты, ДатаНачисления, ПериодС/ПериодПо
   
2. Are documents 033 (30.12.2025) and 028 (10.12.2025) from the same period?
   - Look at their period-related fields
   
3. How does БДДС register track periods?
   - Check if it has ПериодНачисления or КодМесяца fields
   
4. Is there a salary month code field?
   - КодМесяца, КодПлатежа, etc.
""")

print("=" * 100)
print("Investigation complete!")
print("=" * 100)
