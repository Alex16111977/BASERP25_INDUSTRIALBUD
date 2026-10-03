import win32com.client, sys, pythoncom
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
_v8 = None
def connect(which='buh'):
    global _v8
    if _v8 is None:
        _v8 = win32com.client.Dispatch("V83.COMConnector")
    return _v8.Connect(CONN_BUH if which=='buh' else CONN_ERP)
def q(c, text, params=None):
    Q = c.NewObject("Запрос")
    Q.Text = text
    for k, v in (params or {}).items():
        Q.SetParameter(k, v)
    try:
        return Q.Execute().Выгрузить()
    except Exception as e:
        if hasattr(e, 'excepinfo') and e.excepinfo:
            raise RuntimeError(e.excepinfo[2])
        raise
def rows(c, vt, cols):
    S = c.String
    out = []
    for i in range(vt.Количество()):
        r = vt.Получить(i)
        d = {}
        for col in cols:
            v = r.Получить(vt.Колонки.Найти(col).Индекс) if False else getattr(r, col)
            if isinstance(v, (str, int, float, bool)) or v is None:
                d[col] = v
            else:
                try:
                    d[col] = S(v)
                except Exception:
                    d[col] = str(v)
        out.append(d)
    return out
