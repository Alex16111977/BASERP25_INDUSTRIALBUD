# -*- coding: utf-8 -*-
"""
Наскрізне приймання перепроведення обробки А_РозрахункиПДВЗаПершоюПодією (ЗМІНЮЄ базу: перепроводить 1–2 документи).

Серверний контекст: фонове завдання ДлительныеОперации.ВыполнитьПроцедуруМодуляОбъектаОбработки → ПерепровестиВФоне
(зовнішнє COM-з'єднання BuhBud не проводить документи з підписками server-only модулів).

Документи тесту:
  A) приклад користувача — розбіжні документи ключа «Рахунок на оплату постачальника ІБ00-003315» (оплата 000013825);
  B) один розбіжний документ відкритого періоду, якого НЕМАЄ в черзі обміну з ЕРП (перевірка зняття з черги).

Інваріанти після перепроведення:
  1. по ключах A і B |факт ПС − очікувано| ≤ 0,05 для кожного документа (ключ зник з аналізу);
  2. сальдо 6442 по ключу A = сума незареєстрованих ПН (тут 0: ПН 433,33 зареєстрована);
  3. оплата 000013825 не має проводки Дт 6442;
  4. документ B не з'явився в черзі обміну; документи, що були в черзі, там і лишились.
Запуск:  python test_pdv6442_perepr.py
"""
import os, sys, time, datetime
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
EPF = os.path.abspath(os.path.join(HERE, '..', '..', 'Обработки', 'А_РозрахункиПДВЗаПершоюПодією.epf'))
CONN = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
FAILS = []


def check(cond, msg):
    print(('OK   ' if cond else 'FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def q(c, text, **params):
    z = c.NewObject("Запрос")
    z.Text = text
    for k, v in params.items():
        z.SetParameter(k, v)
    return z.Execute().Выгрузить()


def registered(c, ref):
    name = ref.Метаданные().Имя
    tz = q(c, f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ Документ.{name}.Изменения КАК Изм ГДЕ Изм.Ссылка = &С", С=ref)
    return int(tz.Получить(0).Кол) > 0


def analyse(c, date_from):
    proc = c.ВнешниеОбработки.Создать(EPF, False)
    today = datetime.datetime.now()
    proc.ДатаНачала = date_from
    proc.ДатаОкончания = datetime.datetime(today.year, today.month, today.day, 12, 0, 0)
    proc.Анализировать()
    return proc


def main():
    c = win32com.client.Dispatch("V83.COMConnector").Connect(CONN)
    date_from = datetime.datetime(2026, 8, 1, 12, 0, 0)
    proc = analyse(c, date_from)
    keyA, docsA, docB, keyB = None, [], None, None
    for i in range(proc.Документы.Количество()):
        r = proc.Документы.Получить(i)
        if '000013825' in c.String(r.Документ):
            keyA = r.КлючИД
    for i in range(proc.Документы.Количество()):
        r = proc.Документы.Получить(i)
        if keyA and r.КлючИД == keyA and r.Отметка:
            docsA.append(r.Документ)
    already_fixed = keyA is None
    for i in range(0 if already_fixed else proc.Документы.Количество()):
        r = proc.Документы.Получить(i)
        if r.Отметка and r.КлючИД != keyA and not registered(c, r.Документ):
            docB, keyB = r.Документ, r.КлючИД
            break
    if already_fixed:
        print('приклад 000013825 уже без розбіжності — перевіряю лише інваріанти (без перепроведення)')
    else:
        check(len(docsA) > 0, f'ключ прикладу знайдено, документів до перепроведення: {len(docsA)}')
    print('документ B (не в черзі обміну):', c.String(docB) if docB is not None else '—')
    dealA = None
    for i in range(proc.Ключи.Количество()):
        r = proc.Ключи.Получить(i)
        if r.КлючИД == keyA:
            dealA = r.Сделка
    test_docs = list(docsA) + ([docB] if docB is not None else [])
    reg_before = {c.String(d): registered(c, d) for d in test_docs}
    print('у черзі обміну ДО:', reg_before)

    if already_fixed:
        test_docs = []
    # фонове перепроведення
    arr = c.NewObject("Массив")
    for d in test_docs:
        arr.Добавить(d)
    pe = c.NewObject("Структура")
    pe.Вставить("Документы", arr)
    pe.Вставить("РазрешитьЗакрытыйПериод", False)
    params = c.NewObject("Структура")
    params.Вставить("ЭтоВнешняяОбработка", True)
    params.Вставить("ИмяОбработки", EPF)
    params.Вставить("ИмяМетода", "ПерепровестиВФоне")
    params.Вставить("ПараметрыВыполнения", pe)
    pv = c.ДлительныеОперации.ПараметрыВыполненияВФоне(None)
    pv.ЗапуститьВФоне = True
    pv.ОжидатьЗавершение = 0
    pv.АдресРезультата = "строка"
    pv.НаименованиеФоновогоЗадания = "Тест А_РозрахункиПДВЗаПершоюПодією"
    if not test_docs:
        return finish(c, date_from, keyA, keyB, None, [], {})
    res = c.ДлительныеОперации.ВыполнитьВФоне("ДлительныеОперации.ВыполнитьПроцедуруМодуляОбъектаОбработки", params, pv)
    job = c.BackgroundJobs.FindByUUID(res.ИдентификаторЗадания)
    t0 = time.time()
    job = job.WaitForCompletion(600)
    job = c.BackgroundJobs.FindByUUID(res.ИдентификаторЗадания)
    print(f'фон: {time.time() - t0:.0f} c, стан {c.String(job.State)}')
    if job.ErrorInfo is not None:
        print('ПОМИЛКА ФОНУ:', c.String(job.ErrorInfo.Описание) if hasattr(job.ErrorInfo, 'Описание') else c.String(job.ErrorInfo))
    check(c.String(job.State) in ('Завершено', 'Completed', 'Задание выполнено', 'Завдання виконано'),
          'фонове завдання завершено без помилок')
    msgs = job.ПолучитьСообщенияПользователю(True) if hasattr(job, 'ПолучитьСообщенияПользователю') else None
    # журнал: останній запис тесту
    tz = c.NewObject("ТаблицаЗначений")
    otb = c.NewObject("Структура")
    otb.Вставить("Событие", "А_РозрахункиПДВЗаПершоюПодією")
    otb.Вставить("ДатаНачала", datetime.datetime.now() - datetime.timedelta(minutes=15))
    c.ВыгрузитьЖурналРегистрации(tz, otb, "Дата,Уровень,Комментарий")
    for i in range(tz.Количество()):
        r = tz.Получить(i)
        print('  ЖР:', c.String(r.Дата), c.String(r.Уровень), r.Комментарий)

    return finish(c, date_from, keyA, keyB, dealA, test_docs, reg_before)


def finish(c, date_from, keyA, keyB, dealA, test_docs, reg_before):
    # інваріанти
    proc2 = analyse(c, date_from)
    left = {}
    for i in range(proc2.Документы.Количество()):
        r = proc2.Документы.Получить(i)
        if r.КлючИД in (keyA, keyB) and r.ЕстьРасхождение:
            left.setdefault(r.КлючИД, []).append((c.String(r.Документ), r.ОжидаемоПС, r.ФактПС))
    check(keyA not in left, f'ключ A без розбіжностей {left.get(keyA, "")}')
    if keyB is not None:
        check(keyB not in left, f'ключ B без розбіжностей {left.get(keyB, "")}')
    if dealA is None:
        dealA = q(c, 'ВЫБРАТЬ С.Ссылка ИЗ Документ.СчетНаОплатуПоставщика КАК С ГДЕ С.Номер = "ІБ00-003315" И ГОД(С.Дата) = 2026').Получить(0).Ссылка
    if dealA is not None:
        tz = q(c, """ВЫБРАТЬ ЕСТЬNULL(СУММА(О.СуммаОстатокДт - О.СуммаОстатокКт), 0) КАК С
            ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет = &Сч, , Субконто3 = &Сд) КАК О""",
               Сч=c.ПланыСчетов.Хозрасчетный.НалоговыйКредитНеподтвержденный, Сд=dealA)
        bal = float(tz.Получить(0).С)
        print('сальдо 6442 по ключу A:', bal)
        check(abs(bal) < 0.01, 'сальдо 6442 по ключу A = 0 (ПН 433,33 зареєстрована)')
    pay = q(c, 'ВЫБРАТЬ П.Ссылка ИЗ Документ.СписаниеСРасчетногоСчета КАК П ГДЕ П.Номер = "000013825" И ГОД(П.Дата) = 2026').Получить(0).Ссылка
    tz = q(c, """ВЫБРАТЬ Х.СчетДт.Код КАК Дт, Х.СчетКт.Код КАК Кт, Х.Сумма КАК С, Х.Содержание КАК Сод
        ИЗ РегистрБухгалтерии.Хозрасчетный КАК Х ГДЕ Х.Регистратор = &Д""", Д=pay)
    rows = [(r.Дт, r.Кт, float(r.С), r.Сод) for r in (tz.Получить(i) for i in range(tz.Количество()))]
    print('проводки 000013825 ПІСЛЯ:', rows)
    check(not any(r[0] == '6442' for r in rows), 'оплата 000013825 без Дт 6442')
    reg_after = {c.String(d): registered(c, d) for d in test_docs}
    print('у черзі обміну ПІСЛЯ:', reg_after)
    check(all(reg_after[k] == v for k, v in reg_before.items()),
          'черга обміну: хто не був — не з\'явився, хто був — лишився')

    print('ПІДСУМОК:', 'УСПІХ' if not FAILS else f'ПОМИЛОК {len(FAILS)}')
    sys.exit(1 if FAILS else 0)


if __name__ == '__main__':
    main()
