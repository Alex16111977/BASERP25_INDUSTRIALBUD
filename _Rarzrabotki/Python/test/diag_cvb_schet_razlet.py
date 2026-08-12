# -*- coding: utf-8 -*-
"""ЦВБ: пообъектная расшифровка расхождения по ОДНОМУ банковскому счёту (READ-ONLY).

Отвечает на вопрос «плагин говорит: 809 документів співпало, з діями 0 — а где разлёт?».
Плагин сверяет ПОТОК КАЖДОГО документа ERP с потоком ЕГО пары в Бух и молчит, если пары
совпали; документ, которого в Бух нет ВООБЩЕ (и который не подпадает под правило A/переказ),
он покажет, а вот перекос между СОСТАВОМ оборотов двух баз на уровне счёта — нет.

Здесь строится честное сравнение двух списков оборотов по счёту:
  ERP: РегистрНакопления.ДенежныеСредстваБезналичные.Обороты (по регистраторам)
  Бух: РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто, СчетДт/СчетКт В ИЕРАРХИИ
       СчетаВБанках и Субконто1 = банковский счёт (та же логика, что в плагине)
и сводятся по UUID документа (обмен хранит UUID 1:1).

Запуск:
  python diag_cvb_schet_razlet.py --iban UA973005280000026004000010559 --from 2026-04-01 --to 2026-04-30
  python diag_cvb_schet_razlet.py --uid ebcb5323-e5fc-11eb-a208-000c299fb278 --from 2026-04-01 --to 2026-04-30
"""
import argparse
import datetime
import sys
from pathlib import Path

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
REPORTS = Path(__file__).resolve().parent / "reports"

parser = argparse.ArgumentParser()
parser.add_argument("--iban")
parser.add_argument("--uid")
parser.add_argument("--from", dest="dt_from", required=True)
parser.add_argument("--to", dest="dt_to", required=True)
parser.add_argument("--out", default="cvb_schet_razlet.md")
args = parser.parse_args()
assert args.iban or args.uid, "нужен --iban или --uid счёта"

nach = datetime.datetime.strptime(args.dt_from, "%Y-%m-%d")
kon = datetime.datetime.strptime(args.dt_to, "%Y-%m-%d").replace(hour=23, minute=59, second=59)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)
buh = v8.Connect(CONN_BUH)
S = erp.String
SB = buh.String

# --- счёт в ERP ---------------------------------------------------------------
if args.uid:
    schet = erp.Справочники.БанковскиеСчетаОрганизаций.ПолучитьСсылку(
        erp.NewObject("УникальныйИдентификатор", args.uid))
else:
    q = erp.NewObject("Запрос")
    q.SetParameter("IBAN", args.iban)
    q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 БС.Ссылка КАК Ссылка
ИЗ Справочник.БанковскиеСчетаОрганизаций КАК БС
ГДЕ БС.НомерСчета = &IBAN
"""
    rr = q.Execute().Выгрузить()
    assert rr.Количество(), f"счёт с IBAN {args.iban} не найден в ERP"
    schet = rr.Получить(0).Ссылка

schet_uid = str(S(schet.УникальныйИдентификатор()))
print(f"=== Счёт ERP: {S(schet)} (UID {schet_uid}) ===", flush=True)
print(f"    период {nach:%d.%m.%Y} — {kon:%d.%m.%Y}", flush=True)

# --- обороты ERP по регистраторам ---------------------------------------------
q = erp.NewObject("Запрос")
q.SetParameter("Нач", nach)
q.SetParameter("Кон", kon)
q.SetParameter("Счет", schet)
q.Text = """
ВЫБРАТЬ
	Об.Регистратор КАК Регистратор,
	ПРЕДСТАВЛЕНИЕ(Об.Регистратор) КАК Представление,
	Об.СуммаПриход КАК Приход,
	Об.СуммаРасход КАК Расход
ИЗ
	РегистрНакопления.ДенежныеСредстваБезналичные.Обороты(
		&Нач, &Кон, Авто, БанковскийСчет = &Счет) КАК Об
"""
tz = q.Execute().Выгрузить()
erp_docs = {}
for i in range(tz.Количество()):
    row = tz.Получить(i)
    uid = str(S(row.Регистратор.УникальныйИдентификатор()))
    erp_docs[uid] = {
        "предст": str(S(row.Представление)),
        "тип": str(S(row.Регистратор.Метаданные().Имя)),
        "приход": float(row.Приход),
        "расход": float(row.Расход),
    }
erp_prih = sum(d["приход"] for d in erp_docs.values())
erp_rash = sum(d["расход"] for d in erp_docs.values())
print(f"    ERP: документов {len(erp_docs)}, приход {erp_prih:,.2f}, расход {erp_rash:,.2f}",
      flush=True)

# --- счёт в Бух (по UUID, затем по IBAN) --------------------------------------
buh_schet = None
try:
    probe = buh.Справочники.БанковскиеСчета.ПолучитьСсылку(
        buh.NewObject("УникальныйИдентификатор", schet_uid))
    if probe.ПолучитьОбъект() is not None:
        buh_schet = probe
except Exception:
    pass
if buh_schet is None:
    iban = str(S(schet.НомерСчета))
    qb = buh.NewObject("Запрос")
    qb.SetParameter("НомерСчета", iban)
    qb.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 БС.Ссылка КАК Ссылка
ИЗ Справочник.БанковскиеСчета КАК БС
ГДЕ БС.НомерСчета = &НомерСчета И БС.Наименование <> ""
УПОРЯДОЧИТЬ ПО БС.Наименование
"""
    rb = qb.Execute().Выгрузить()
    assert rb.Количество(), f"счёт {iban} не найден в BuhBud"
    buh_schet = rb.Получить(0).Ссылка
print(f"=== Счёт Бух: {SB(buh_schet)} ===", flush=True)

# --- обороты Бух по регистраторам ---------------------------------------------
qb = buh.NewObject("Запрос")
qb.SetParameter("Нач", nach)
qb.SetParameter("Кон", kon)
qb.SetParameter("Банк", buh_schet)
qb.Text = """
ВЫБРАТЬ
	Т.Регистратор КАК Регистратор,
	ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК Представление,
	СУММА(Т.Приход) КАК Приход,
	СУММА(Т.Расход) КАК Расход
ИЗ
	(ВЫБРАТЬ Д.Регистратор КАК Регистратор, Д.Сумма КАК Приход, 0 КАК Расход
	ИЗ РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(&Нач, &Кон, , , ) КАК Д
	ГДЕ Д.СчетДт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках))
		И Д.СубконтоДт1 = &Банк

	ОБЪЕДИНИТЬ ВСЕ

	ВЫБРАТЬ Д.Регистратор, 0, Д.Сумма
	ИЗ РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(&Нач, &Кон, , , ) КАК Д
	ГДЕ Д.СчетКт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках))
		И Д.СубконтоКт1 = &Банк) КАК Т
СГРУППИРОВАТЬ ПО
	Т.Регистратор
"""
tb = qb.Execute().Выгрузить()
buh_docs = {}
for i in range(tb.Количество()):
    row = tb.Получить(i)
    uid = str(SB(row.Регистратор.УникальныйИдентификатор()))
    buh_docs[uid] = {
        "предст": str(SB(row.Представление)),
        "приход": float(row.Приход),
        "расход": float(row.Расход),
    }
buh_prih = sum(d["приход"] for d in buh_docs.values())
buh_rash = sum(d["расход"] for d in buh_docs.values())
print(f"    Бух: документов {len(buh_docs)}, приход {buh_prih:,.2f}, расход {buh_rash:,.2f}",
      flush=True)
print(f"\n    Δ приход = {erp_prih - buh_prih:,.2f}; Δ расход = {erp_rash - buh_rash:,.2f}",
      flush=True)

# --- сведение по UUID ----------------------------------------------------------
# У переводов пара в Бух живёт под UUID парного документа ERP (списания) —
# подтягиваем связь, иначе выписка выглядит «сиротой».
pary = {}
qp = erp.NewObject("Запрос")
qp.SetParameter("Нач", nach)
qp.SetParameter("Кон", kon)
qp.Text = """
ВЫБРАТЬ
	П.Ссылка КАК Выписка,
	П.А_СписаниеБезналичныхДенежныхСредств КАК Списание
ИЗ
	Документ.ПоступлениеБезналичныхДенежныхСредств КАК П
ГДЕ
	П.Дата МЕЖДУ &Нач И &Кон
	И П.А_СписаниеБезналичныхДенежныхСредств <> ЗНАЧЕНИЕ(Документ.СписаниеБезналичныхДенежныхСредств.ПустаяСсылка)
"""
tp = qp.Execute().Выгрузить()
for i in range(tp.Количество()):
    row = tp.Получить(i)
    pary[str(S(row.Выписка.УникальныйИдентификатор()))] = \
        str(S(row.Списание.УникальныйИдентификатор()))

only_erp, only_buh, both = [], [], []
matched_buh = set()
for uid, d in erp_docs.items():
    alt = pary.get(uid)
    b = buh_docs.get(uid) or (buh_docs.get(alt) if alt else None)
    if b:
        matched_buh.add(uid if uid in buh_docs else alt)
        if abs((d["приход"] - d["расход"]) - (b["приход"] - b["расход"])) >= 0.005:
            both.append((uid, d, b))
    else:
        only_erp.append((uid, d))
for uid, b in buh_docs.items():
    if uid not in matched_buh and uid not in erp_docs:
        only_buh.append((uid, b))

only_erp.sort(key=lambda x: -abs(x[1]["приход"] - x[1]["расход"]))
only_buh.sort(key=lambda x: -abs(x[1]["приход"] - x[1]["расход"]))

out = [f"# Расшифровка расхождения по счёту {S(schet)}",
       f"Период {args.dt_from} .. {args.dt_to}", "",
       f"| Сторона | Документов | Приход | Расход |", "|---|---|---|---|",
       f"| ERP | {len(erp_docs)} | {erp_prih:,.2f} | {erp_rash:,.2f} |",
       f"| Бух | {len(buh_docs)} | {buh_prih:,.2f} | {buh_rash:,.2f} |",
       f"| **Δ** | | **{erp_prih - buh_prih:,.2f}** | **{erp_rash - buh_rash:,.2f}** |", ""]

def block(title, rows, side):
    out.append(f"## {title}: {len(rows)}")
    out.append("")
    if not rows:
        out.append("(нет)")
        out.append("")
        return 0.0
    out.append("| Нетто | Тип | Документ |")
    out.append("|---|---|---|")
    total = 0.0
    for uid, d in rows:
        netto = d["приход"] - d["расход"]
        total += netto
        out.append(f"| {netto:,.2f} | {d.get('тип', '')} | {d['предст']} |")
    out.append("")
    out.append(f"**Σ нетто = {total:,.2f}**")
    out.append("")
    print(f"\n--- {title}: {len(rows)}, Σ нетто {total:,.2f} ---", flush=True)
    for uid, d in rows[:20]:
        print(f"  {d['приход'] - d['расход']:>16,.2f}  {d.get('тип', ''):<38} {d['предст']}",
              flush=True)
    return total

block("Есть в ERP, нет в Бух (обороты по счёту)", only_erp, "erp")
block("Есть в Бух, нет в ERP (обороты по счёту)", only_buh, "buh")

out.append(f"## Пара найдена, но потоки разные: {len(both)}")
out.append("")
if both:
    out.append("| ERP нетто | Бух нетто | Δ | Документ ERP |")
    out.append("|---|---|---|---|")
    for uid, d, b in both:
        de, db = d["приход"] - d["расход"], b["приход"] - b["расход"]
        out.append(f"| {de:,.2f} | {db:,.2f} | {de - db:,.2f} | {d['предст']} |")
        print(f"  РАЗНЫЕ ПОТОКИ: ERP {de:,.2f} vs Бух {db:,.2f} — {d['предст']}", flush=True)
else:
    out.append("(нет)")

REPORTS.mkdir(exist_ok=True)
rep = REPORTS / args.out
rep.write_text("\n".join(out), encoding="utf-8")
print(f"\nОтчёт: {rep}", flush=True)

erp = None
buh = None
print("DIAG DONE", flush=True)
