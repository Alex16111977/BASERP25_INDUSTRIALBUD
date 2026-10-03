import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
_v8 = None
_c = {}
def connect(which='buh'):
    global _v8
    if which in _c:
        return _c[which]
    if _v8 is None:
        _v8 = win32com.client.Dispatch("V83.COMConnector")
    _c[which] = _v8.Connect(CONN_BUH if which == 'buh' else CONN_ERP)
    return _c[which]

def q(c, text, **params):
    Q = c.NewObject("Запрос")
    Q.Text = text
    for k, v in params.items():
        Q.SetParameter(k, v)
    try:
        vt = Q.Execute().Выгрузить()
    except Exception as e:
        if hasattr(e, 'excepinfo') and e.excepinfo:
            print("FAIL:", e.excepinfo[2]); raise
        raise
    cols = [vt.Колонки.Получить(i).Имя for i in range(vt.Колонки.Количество())]
    S = c.String
    out = []
    for i in range(vt.Количество()):
        r = vt.Получить(i)
        d = {}
        for j, col in enumerate(cols):
            v = r.Получить(j)
            if isinstance(v, (str, int, float, bool)) or v is None:
                d[col] = v
            else:
                try:
                    d[col] = S(v)
                except Exception:
                    d[col] = str(v)
        out.append(d)
    return out

def pr(rows, cols=None, limit=500):
    if not rows:
        print("(пусто)"); return
    cols = cols or list(rows[0].keys())
    print(" | ".join(cols))
    for r in rows[:limit]:
        vals = []
        for c in cols:
            v = r.get(c)
            if isinstance(v, float):
                v = round(v, 3)
                if v == int(v): v = int(v)
            vals.append(str(v).strip() if v is not None else "")
        print(" | ".join(vals))
    print(f"-- строк: {len(rows)}")
