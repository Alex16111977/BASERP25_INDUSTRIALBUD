# -*- coding: utf-8 -*-
"""Разове вирівнювання ієрархії Справочник.Подразделения («Підрозділи Казна») у BuhBud (bas_industrialbud) за деревом казни.

Обмін Казна→Бух (план КазначействоBASБухгалтерияУкраины) переносить лише Код і Наименование, Родитель не приходить —
тому в BuhBud довідник плоский. Скрипт ставить елементу BuhBud того самого батька, що в казні.
Відповідність казна→BuhBud — як шукає обмін: спершу UUID, потім Код (дублів кодів немає).
Змінюється лише Родитель; коди й назви не чіпаються; елементи BuhBud без пари в казні не чіпаються.
Запис без реєстрації змін в обмінах (довідник входить у план ОбменУправлениеПредприятиемБухгалтерия20 → ERP).
Нові підрозділи обмін і далі привозить без батька — повторний запуск дорівнює їх (вже вирівняні пропускає).
    python buh_podrazdeleniya_ierarhiya_iz_kazny.py                    — dry-run, звіт CSV
    python buh_podrazdeleniya_ierarhiya_iz_kazny.py --apply            — запис (бекап поточних батьків у json)
    python buh_podrazdeleniya_ierarhiya_iz_kazny.py --rollback=<json>  — повернути батьків із бекапу"""
import csv, datetime, io, json, os, sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
ЗАПИС = "--apply" in sys.argv
ВІДКАТ = next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--rollback=")), None)
ПАПКА = os.path.dirname(os.path.abspath(__file__))
com = win32com.client.Dispatch("V83.COMConnector")
kz = com.Connect('Srvr="SQLSERVER";Ref="kazna";Usr="cfo";Pwd="2442"')
bh = com.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

ТЕКСТ = """ВЫБРАТЬ С.Ссылка КАК Ссылка, С.Код КАК Код, С.Наименование КАК Имя, С.Родитель КАК Родитель
ИЗ Справочник.Подразделения КАК С"""


def довідник(c):
    q = c.NewObject("Запрос")
    q.Text = ТЕКСТ
    s = q.Execute().Выбрать()
    d = {}
    while s.Следующий():
        d[c.String(s.Ссылка.УникальныйИдентификатор())] = {
            "реф": s.Ссылка, "код": s.Код.strip(), "назва": s.Имя,
            "батько": "" if s.Родитель.Пустая() else c.String(s.Родитель.УникальныйИдентификатор())}
    return d


def змін_у_обміні():
    q = bh.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ Справочник.Подразделения.Изменения КАК Изм"
    s = q.Execute().Выбрать()
    s.Следующий()
    return s.К


def записати(uuid, батько_uuid):
    о = Б[uuid]["реф"].ПолучитьОбъект()
    о.Родитель = Б[батько_uuid]["реф"] if батько_uuid else bh.Справочники.Подразделения.ПустаяСсылка()
    # зміна локальна для BuhBud: не реєструвати в обмінах (БСП ОбменДаннымиСобытия)
    о.ДополнительныеСвойства.Вставить("ОтключитьМеханизмРегистрацииОбъектов", True)
    о.Записать()


def записати_всі(зміни):
    до = змін_у_обміні()
    записано, помилки = 0, []
    for uuid, батько_uuid in зміни:
        try:
            записати(uuid, батько_uuid)
            записано += 1
        except Exception as e:
            помилки.append((Б[uuid]["код"], str(e.excepinfo[2] if getattr(e, "excepinfo", None) else e)[:150]))
    print(f"записано: {записано}; помилок: {len(помилки)} {помилки[:5]}; "
          f"реєстрацій в обміні: було {до}, стало {змін_у_обміні()}")


def назва(d, uuid):
    return d[uuid]["назва"] if uuid in d else ""


К, Б = довідник(kz), довідник(bh)

if ВІДКАТ:
    with io.open(ВІДКАТ, encoding="utf-8") as f:
        бекап = json.load(f)
    записати_всі([(р["uuid"], р["було"]) for р in бекап if р["uuid"] in Б])
    sys.exit(0)

за_кодом = {}
for u, е in Б.items():
    за_кодом.setdefault(е["код"], []).append(u)


def пара(uk):
    if uk in Б:
        return uk, "UUID"
    кандидати = за_кодом.get(К[uk]["код"], [])
    if len(кандидати) == 1:
        return кандидати[0], "код"
    return None, "дубль коду" if кандидати else "немає в BuhBud"


def глибина(uk):
    г = 0
    while К[uk]["батько"] in К and г < 20:
        uk, г = К[uk]["батько"], г + 1
    return г


відп = {u: пара(u) for u in К}
рядки, зміни = [], []
for uk in sorted(К, key=lambda u: (глибина(u), К[u]["код"])):
    ub, як = відп[uk]
    к = К[uk]
    рядок = {"код": к["код"], "назва казни": к["назва"], "назва BuhBud": назва(Б, ub) if ub else "",
             "знайдено": як, "батько в казні": назва(К, к["батько"]),
             "батько BuhBud було": назва(Б, Б[ub]["батько"]) if ub else "", "дія": ""}
    рядки.append(рядок)
    if ub is None:
        рядок["дія"] = "пропуск: " + як
        continue
    ціль = відп.get(к["батько"], (None, ""))[0] if к["батько"] else ""
    if ціль is None:
        рядок["дія"] = "пропуск: батька немає в BuhBud"
    elif Б[ub]["батько"] == ціль:
        рядок["дія"] = "вже як у казні"
    else:
        рядок["дія"] = "поставити батька" if ціль else "очистити батька"
        зміни.append((ub, ціль))

підсумок = {}
for р in рядки:
    підсумок[р["дія"]] = підсумок.get(р["дія"], 0) + 1
print(f"казна {len(К)}, BuhBud {len(Б)}; знайдено за UUID {sum(1 for _, я in відп.values() if я == 'UUID')}, "
      f"за кодом {sum(1 for _, я in відп.values() if я == 'код')}")
print("підсумок:", підсумок)
пари = {ub for ub, _ in відп.values() if ub}
print("в BuhBud без пари в казні (не чіпаються):", [(е["код"], е["назва"]) for u, е in Б.items() if u not in пари])
print("назви різні (казна / BuhBud; не чіпаються):",
      [(р["код"], р["назва казни"], р["назва BuhBud"]) for р in рядки if р["назва BuhBud"] and р["назва казни"] != р["назва BuhBud"]])
звіт = os.path.join(ПАПКА, "buh_podrazdeleniya_ierarhiya_report.csv")
with io.open(звіт, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(рядки[0].keys()), delimiter=";")
    w.writeheader()
    w.writerows(рядки)
print("звіт:", звіт)

if not ЗАПИС:
    print(f"DRY-RUN, до запису: {len(зміни)}. Для запису: --apply")
    sys.exit(0)
бекап = os.path.join(ПАПКА, f"buh_podrazdeleniya_ierarhiya_backup_{datetime.datetime.now():%Y%m%d_%H%M}.json")
with io.open(бекап, "w", encoding="utf-8") as f:
    json.dump([{"uuid": ub, "код": Б[ub]["код"], "назва": Б[ub]["назва"], "було": Б[ub]["батько"]} for ub, _ in зміни],
              f, ensure_ascii=False, indent=1)
print("бекап:", бекап)
записати_всі(зміни)
