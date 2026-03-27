# -*- coding: utf-8 -*-
"""
Test script to check structure of bank payment documents in ERP.
"""

import win32com.client

OUTPUT_FILE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\bank_docs_structure_output.txt"

def iterate_collection(collection):
    items = []
    try:
        for item in collection:
            items.append(item)
    except Exception:
        try:
            count = collection.Count()
            for i in range(count):
                items.append(collection.Get(i))
        except Exception:
            pass
    return items

def type_to_str(type_desc):
    parts = []
    try:
        for t in iterate_collection(type_desc.Types()):
            try:
                parts.append(str(t))
            except:
                parts.append("?")
    except:
        return "?"
    return ", ".join(parts) if parts else "?"

def analyze_document(conn, doc_name, f):
    line = "=" * 80
    f.write(f"{line}\n")
    f.write(f"DOCUMENT: {doc_name}\n")
    f.write(f"{line}\n")

    meta = getattr(conn.Metadata.Documents, doc_name)

    # Standard attributes
    f.write("\n--- Standard Attributes ---\n")
    for attr in iterate_collection(meta.StandardAttributes):
        try:
            f.write(f"  {attr.Name}\n")
        except:
            pass

    # Header attributes
    attrs = iterate_collection(meta.Attributes)
    f.write(f"\n--- Header Attributes: {len(attrs)} total ---\n")
    for attr in attrs:
        try:
            f.write(f"  {attr.Name}\n")
        except:
            pass

    # Tabular sections
    tab_sections = iterate_collection(meta.TabularSections)
    f.write(f"\n--- Tabular Sections: {len(tab_sections)} total ---\n")
    for ts in tab_sections:
        try:
            ts_attrs = iterate_collection(ts.Attributes)
            f.write(f"\n  TabSection: {ts.Name} ({len(ts_attrs)} columns)\n")
            f.write(f"  {'~' * 60}\n")
            for tsAttr in ts_attrs:
                try:
                    f.write(f"    {tsAttr.Name}\n")
                except:
                    pass
        except Exception as e:
            f.write(f"  (error: {e})\n")

    f.write("\n\n")

def main():
    print("Connecting to ERP...")
    v8 = win32com.client.Dispatch("V83.COMConnector")
    CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
    conn = v8.Connect(CONN_ERP)
    print("Connected OK.")

    docs = [
        "СписаниеБезналичныхДенежныхСредств",
        "ПоступлениеБезналичныхДенежныхСредств",
    ]

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        for doc_name in docs:
            try:
                analyze_document(conn, doc_name, f)
            except Exception as e:
                f.write(f"ERROR analyzing {doc_name}: {e}\n")

    # Print output to console too
    with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
        print(f.read())

    print(f"Output saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
