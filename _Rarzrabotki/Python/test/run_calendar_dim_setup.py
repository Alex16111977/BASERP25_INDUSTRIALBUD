"""Виконати calendar_dim_olapbaserp.sql на OlapBASERP."""
import sys
import re
import pyodbc

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = (
    'Driver={ODBC Driver 17 for SQL Server};'
    'Server=localhost;Database=OlapBASERP;UID=sa;PWD=Brw739182465!;'
    'Encrypt=no;TrustServerCertificate=yes'
)
SCRIPT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap\scripts\calendar_dim_olapbaserp.sql"


def split_batches(sql_text):
    """Split SQL by 'GO' delimiter (на власному рядку, без коментарів)."""
    return [b.strip() for b in re.split(r'^\s*GO\s*$', sql_text, flags=re.MULTILINE) if b.strip()]


def main():
    with open(SCRIPT, encoding='utf-8') as f:
        sql_text = f.read()

    batches = split_batches(sql_text)
    print(f"Batches: {len(batches)}")

    conn = pyodbc.connect(CONN, autocommit=False)
    cur = conn.cursor()

    for i, batch in enumerate(batches, 1):
        first_line = batch.split('\n', 1)[0][:80]
        print(f"  [{i}] {first_line}")
        try:
            cur.execute(batch)
            # Read all result sets if any
            while True:
                if cur.description:
                    rows = cur.fetchall()
                    cols = [d[0] for d in cur.description]
                    print(f"      ↳ {len(rows)} rows, cols={cols}")
                    for r in rows[:5]:
                        print(f"        {r}")
                if not cur.nextset():
                    break
        except pyodbc.Error as e:
            print(f"      FAIL: {e}")
            conn.rollback()
            return 1

    conn.commit()
    conn.close()
    print("\nOK — committed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
