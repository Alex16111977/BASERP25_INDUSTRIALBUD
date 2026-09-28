# -*- coding: utf-8 -*-
"""
Rule #-1: претест запитів обробки А_РозрахункиПДВЗаПершоюПодією на живій базі bas_industrialbud.

Витягує тексти з функцій ТекстЗапроса*() модуля об'єкта (1:1 як у BSL) і виконує їх з параметрами движка.
Запуск:  python test_pdv6442_zapros.py [шлях_до_ObjectModule.bsl]
"""
import os, re, sys, datetime
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_BSL = os.path.join(HERE, '..', '..', 'Обработки', 'А_РозрахункиПДВЗаПершоюПодією', 'Ext', 'ObjectModule.bsl')
CONNS = ['Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"',
         'Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"']


def query_texts(path):
    src = open(path, encoding='utf-8-sig').read()
    result = {}
    for m in re.finditer(r'Функция (ТекстЗапроса\w+)\(\)\s*\n(.*?)КонецФункции', src, re.S):
        body = m.group(2)
        lines = []
        for line in body.splitlines():
            t = line.strip()
            if t.startswith('"'):
                lines.append(t[1:])
            elif t.startswith('|'):
                lines.append(t[1:])
        text = '\n'.join(lines)
        assert text.endswith('";'), m.group(1)
        result[m.group(1)] = text[:-2].replace('""', '"')
    return result


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    last = None
    for s in CONNS:
        try:
            return v8.Connect(s)
        except Exception as e:  # noqa
            last = e
    raise SystemExit(f"FAIL connect: {last}")


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_BSL
    texts = query_texts(path)
    print('запитів у модулі:', len(texts), sorted(texts))
    c = connect()
    today = datetime.datetime.now()
    base = dict(
        ДатаНачала=datetime.datetime(2026, 1, 1, 0, 0, 0),
        ДатаОкончания=datetime.datetime(today.year, today.month, today.day, 23, 59, 59),
        Гривня=c.Константы.ВалютаРегламентированногоУчета.Получить(),
        ОтборОрганизация=False, Организация=c.Справочники.Организации.ПустаяСсылка(),
        ОтборКонтрагент=False, Контрагент=c.Справочники.Контрагенты.ПустаяСсылка(),
        ОтборДоговор=False, ДоговорКонтрагента=c.Справочники.ДоговорыКонтрагентов.ПустаяСсылка(),
    )
    extra = {
        'ТекстЗапросаСальдо6442': dict(
            КонецПериода=base['ДатаОкончания'],
            Счет6442=c.ПланыСчетов.Хозрасчетный.НалоговыйКредитНеподтвержденный,
            Договоры=None),
        'ТекстЗапросаУзлыОбменаЕРП': {},
        'ТекстЗапросаДатаЗапрета': {},
    }
    fails = 0
    for name, text in sorted(texts.items()):
        z = c.NewObject("Запрос")
        z.Text = text
        params = extra.get(name, base)
        for k, v in params.items():
            if k == 'Договоры':
                arr = c.NewObject("Массив")
                sel = c.NewObject("Запрос")
                sel.Text = 'ВЫБРАТЬ ПЕРВЫЕ 50 Д.Ссылка ИЗ Справочник.ДоговорыКонтрагентов КАК Д'
                vs = sel.Execute().Выбрать()
                while vs.Следующий():
                    arr.Добавить(vs.Ссылка)
                v = arr
            z.SetParameter(k, v)
        try:
            tz = z.Execute().Выгрузить()
            print(f"OK   {name}: рядків={tz.Количество()}")
        except Exception as e:
            fails += 1
            info = e.excepinfo[2] if getattr(e, 'excepinfo', None) else e
            print(f"FAIL {name}: {info}")
    # фільтри з заповненими відборами
    z = c.NewObject("Запрос")
    z.Text = texts['ТекстЗапросаДвижения']
    org = c.NewObject("Запрос")
    org.Text = 'ВЫБРАТЬ ПЕРВЫЕ 1 О.Ссылка ИЗ Справочник.Организации КАК О ГДЕ О.КодПоЕДРПОУ = "40645273"'
    o = org.Execute().Выбрать()
    o.Следующий()
    p = dict(base, ОтборОрганизация=True, Организация=o.Ссылка)
    for k, v in p.items():
        z.SetParameter(k, v)
    try:
        print('OK   Движения з відбором організації: рядків=', z.Execute().Выгрузить().Количество())
    except Exception as e:
        fails += 1
        print('FAIL Движения з відбором:', e.excepinfo[2] if getattr(e, 'excepinfo', None) else e)
    print('ПІДСУМОК:', 'УСПІХ' if fails == 0 else f'ПОМИЛОК {fails}')
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
