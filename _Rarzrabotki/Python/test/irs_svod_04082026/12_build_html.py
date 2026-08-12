# -*- coding: utf-8 -*-
"""Аналитическая записка МД IRS 2026 (HTML)."""
import json, sys, os
sys.stdout.reconfigure(encoding='utf-8')

OUT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\А_ПланФактныйПроизводствоПолный\Письмо директору\04082026"
os.makedirs(OUT, exist_ok=True)

byON = json.load(open("data_byon_calc.json", encoding="utf-8"))
mat = json.load(open("data_materials.json", encoding="utf-8"))
kazna = json.load(open("data_kazna.json", encoding="utf-8"))
xl = json.load(open("data_excel_koshtoris.json", encoding="utf-8"))
t = mat["tot"]


def g(v, dec=2):
    s = ("{:,." + str(dec) + "f}").format(v).replace(",", "\u00a0").replace(".", ",")
    return s


eco = [d for d in byON if d["Відхилення"] > 0.005 and d["ПланГрн"] > 0]
per = [d for d in byON if d["Відхилення"] < -0.005]
eco_s = sum(d["Відхилення"] for d in eco)
per_s = sum(-d["Відхилення"] for d in per)

kz_mat = sum(x["Розхід"] for x in kazna if x["Стаття"] == "Строительные материалы")
kz_in = sum(x["Прихід"] for x in kazna)
kz_out = sum(x["Розхід"] for x in kazna)
kz_zp = sum(x["Розхід"] for x in kazna if x["Стаття"] == "Зарплата производственного персонала")

lip = xl["15"]["график"]["липень"][1] + xl["30"]["график"]["липень"][1]
ser = xl["15"]["график"]["серпень"][1] + xl["30"]["график"]["серпень"][1]

rows_eco = "".join(
    '<tr><td>{}</td><td class="num">{}</td><td class="num">{}</td>'
    '<td class="num good">{}</td><td class="num">{}</td></tr>'.format(
        d["ОН"], g(d["ПланГрн"]), g(d["Прогноз"]), g(d["Відхилення"]),
        g(d["Прогноз"] / d["ПланГрн"] * 100, 1))
    for d in sorted(eco, key=lambda x: -x["Відхилення"])[:8])

rows_per = "".join(
    '<tr><td>{}</td><td class="num">{}</td><td class="num">{}</td>'
    '<td class="num bad">{}</td><td class="num">{}</td></tr>'.format(
        d["ОН"], g(d["ПланГрн"]) if d["ПланГрн"] else "—", g(d["Прогноз"]),
        g(-d["Відхилення"]),
        g(d["Прогноз"] / d["ПланГрн"] * 100, 1) if d["ПланГрн"] else "—")
    for d in sorted(per, key=lambda x: x["Відхилення"])[:8])

rows_buy = "".join(
    '<tr><td>{}</td><td class="num">{}</td><td class="num">{}</td>'
    '<td class="num">{}</td><td class="num">{}</td></tr>'.format(
        d["ОН"], g(d["ПланГрн"]), g(d["ФактГрн"]), g(d["Осталось"]), g(d["Процент"], 1))
    for d in sorted(byON, key=lambda x: -x["Осталось"])[:12])

CSS = """
  :root{--ink:#1f2933;--muted:#6b7683;--line:#dfe3e8;--good:#1b7f4b;--bad:#b3261e;
        --accent:#1f4e79;--accent-bg:#eef3f8;--warn:#8a5a00;--warn-bg:#fff8e6}
  *{box-sizing:border-box}
  body{margin:0;padding:32px 28px 48px;background:#f4f6f8;color:var(--ink);
       font:15px/1.65 "Segoe UI",Tahoma,Arial,sans-serif}
  .sheet{max-width:980px;margin:0 auto;background:#fff;padding:40px 44px;
         border:1px solid var(--line);border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
  h1{font-size:19px;margin:0 0 22px;color:var(--accent)}
  h2{font-size:15px;margin:28px 0 8px;padding-bottom:5px;
     border-bottom:2px solid var(--accent-bg);color:var(--accent)}
  h3{font-size:14px;margin:18px 0 6px;color:var(--ink)}
  p{margin:9px 0}
  table{width:100%;border-collapse:collapse;margin:8px 0 4px;font-size:14px}
  th,td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:right}
  th:first-child,td:first-child{text-align:left}
  thead th{background:#f7f9fb;font-weight:600;color:var(--muted);font-size:13px;
           border-bottom:1px solid #c9d2da;white-space:nowrap}
  tr.total td{font-weight:700;background:#f7f9fb;border-top:1px solid #c9d2da}
  .num{font-variant-numeric:tabular-nums;white-space:nowrap}
  .good{color:var(--good);font-weight:600}
  .bad{color:var(--bad);font-weight:600}
  .warn{margin:14px 0;padding:11px 15px;background:var(--warn-bg);
        border-left:3px solid var(--warn);border-radius:3px;font-size:14px}
  .warn b{color:var(--warn)}
  .att{margin:26px 0 0;padding:10px 14px;background:#f7f9fb;border:1px solid var(--line);
       border-radius:4px;font-size:14px}
  .att b{color:var(--accent)}
  .att ol{margin:6px 0 0 18px;padding:0}
  .att li{margin:3px 0}
  .src{margin-top:20px;font-size:12.5px;color:var(--muted);border-top:1px solid var(--line);padding-top:10px}
  @media print{body{background:#fff;padding:0}.sheet{border:0;box-shadow:none;max-width:none;padding:0}}
"""

HTML = """<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="utf-8">
<title>МД IRS 2026 — закупівлі, кошторис, бюджет станом на 04.08.2026</title>
<style>{css}</style>
</head>
<body>
<div class="sheet">

<p>Доброго дня,</p>

<h1>МД IRS 2026 — матеріали, кошторис і бюджет станом на 04.08.2026</h1>

<p>Розріз — тільки підрозділ «МД IRS 2026» (13 будинків: 6 × IRS 15 м і 7 × IRS 30 м).
Усі суми — <b>з ПДВ</b>, накопичувально від початку проєкту.</p>

<h2>1. Скільки закуплено і скільки лишилось закупити</h2>

<table>
  <thead><tr><th>Показник</th><th>Сума, грн</th><th>% плану</th></tr></thead>
  <tbody>
    <tr><td>План СС «15 м» (184 позиції × 6 будинків)</td><td class="num">{p15}</td><td class="num">—</td></tr>
    <tr><td>План СС «30 м» (170 позицій × 7 будинків)</td><td class="num">{p30}</td><td class="num">—</td></tr>
    <tr class="total"><td>План разом</td><td class="num">{plan}</td><td class="num">100,0</td></tr>
    <tr><td>Закуплено (57 документів, липень 2026)</td><td class="num">{fakt}</td><td class="num">{faktp}</td></tr>
    <tr><td>Лишилось закупити</td><td class="num">{ost}</td><td class="num">{ostp}</td></tr>
    <tr class="total"><td>Прогноз (закуплено + лишилось)</td><td class="num">{prog}</td><td class="num">{progp}</td></tr>
    <tr><td>Відхилення від плану — економія</td><td class="num good">{otkl}</td><td class="num">{otklp}</td></tr>
  </tbody>
</table>

<p>Прогноз нижчий за кошторис на <b>{otkl} грн ({otklp} %)</b>. Тиждень тому (31.07) економія
складала 444 523,90 грн — за тиждень вона <b>зменшилась на 20 295,08 грн</b>.</p>

<p>Економія сальдова: по <b>{neco} позиціях</b> зниження на <b>{ecos} грн</b>,
по <b>{nper} позиціях</b> — перевитрата на <b>{pers} грн</b>.</p>

<h3>Основна економія</h3>
<table>
  <thead><tr><th>Загальна назва</th><th>План, грн</th><th>Прогноз, грн</th><th>Економія, грн</th><th>%</th></tr></thead>
  <tbody>{rows_eco}</tbody>
</table>

<h3>Основна перевитрата</h3>
<table>
  <thead><tr><th>Загальна назва</th><th>План, грн</th><th>Прогноз, грн</th><th>Перевитрата, грн</th><th>%</th></tr></thead>
  <tbody>{rows_per}</tbody>
</table>

<div class="warn"><b>Увага.</b> Обидві картки СС мають <b>однаковий підрозділ закупівель «МД IRS 2026»</b> —
матеріали купуються спільним котлом на всі 13 будинків. Тому розріз 15 м / 30 м коректний
<b>тільки по плану</b>; факт закупівель поділити між проєктами неможливо ані у звіті, ані в обліку.</div>

<h2>2. Що лишилось закупити — найбільші позиції</h2>

<table>
  <thead><tr><th>Загальна назва</th><th>План, грн</th><th>Закуплено, грн</th><th>Лишилось, грн</th><th>Викон., %</th></tr></thead>
  <tbody>{rows_buy}</tbody>
</table>

<p>Повний перелік по <b>141 загальній назві</b> — у додатку, аркуш «Ще закупити».
Найбільша незакрита позиція — <b>генератор: 805 000 грн, закупівель немає взагалі</b>
(13 шт по 61 923,08 грн). Далі — бойлери електричні 192 400 грн і фарба 139 384 грн, теж по нулях.</p>

<h2>3. Кошторис Excel проти структури собівартості</h2>

<table>
  <thead><tr><th>Проєкт</th><th>Кошторис Excel, грн</th><th>План СС (1С), грн</th><th>Розбіжність, грн</th></tr></thead>
  <tbody>
    <tr><td>IRS 15 (6 будинків)</td><td class="num">{x15}</td><td class="num">{p15}</td><td class="num bad">+6 071,97</td></tr>
    <tr><td>IRS 30 (7 будинків)</td><td class="num">{x30}</td><td class="num">{p30}</td><td class="num bad">−16 406,60</td></tr>
    <tr class="total"><td>Разом</td><td class="num">{xtot}</td><td class="num">{plan}</td><td class="num bad">−10 334,60</td></tr>
  </tbody>
</table>

<p>Розбіжність невелика і <b>повністю локалізована у чотирьох позиціях труб</b>:
Труба 40*40*3, Труба 50*50*4, Труба металева Ст3 100*100*3, Труба металева Ст3 50*50*2.
По решті позицій кошторис і СС збігаються копійка в копійку. Деталізація — аркуш «Кошторис vs СС».</p>

<h2>4. Графік закупівель</h2>

<table>
  <thead><tr><th>Місяць</th><th>План IRS 15</th><th>План IRS 30</th><th>План разом</th><th>Факт</th><th>Відхилення</th></tr></thead>
  <tbody>
    <tr><td>Липень</td><td class="num">{l15}</td><td class="num">0,00</td><td class="num">{lip}</td><td class="num">{fakt}</td><td class="num good">+{dlip}</td></tr>
    <tr><td>Серпень</td><td class="num">{s15}</td><td class="num">{s30}</td><td class="num">{ser}</td><td class="num">0,00</td><td class="num bad">−{ser}</td></tr>
    <tr class="total"><td>Разом</td><td class="num">{x15s}</td><td class="num">{x30s}</td><td class="num">{gtot}</td><td class="num">{fakt}</td><td class="num">—</td></tr>
  </tbody>
</table>

<p>Липень закрито з випередженням на <b>{dlip} грн</b>. Основне навантаження — <b>серпень:
{ser} грн</b>, з них {s30} грн по IRS 30. Станом на 04.08 серпневих закупівель ще немає.
Колонка «вересень» у кошторисі — артефакт формули (там продубльовано ПДВ серпня), реального
плану вересня немає.</p>

<h2>5. Гроші: бюджет проти каси</h2>

<p>Факт грошей узято з бази <b>Казна, регістр БДДС</b>, по підрозділу «МД IRS 2026» в ієрархії
(липень 2026):</p>

<table>
  <thead><tr><th>Стаття ДДС</th><th>Каса (Казна), грн</th><th>Факт у бюджеті Excel</th><th>Розбіжність</th></tr></thead>
  <tbody>
    <tr><td>Будівельні матеріали</td><td class="num">{kzm}</td><td class="num">3 570 426,34</td><td class="num">0,00</td></tr>
    <tr><td>Зарплата виробничого персоналу</td><td class="num">{kzzp}</td><td class="num">63 504,63</td><td class="num bad">+19 365,31</td></tr>
    <tr><td>Транспортні витрати</td><td class="num">19 995,52</td><td class="num">19 995,52</td><td class="num">0,00</td></tr>
    <tr><td>МШП</td><td class="num">14 549,08</td><td class="num">14 549,08</td><td class="num">0,00</td></tr>
    <tr><td>Послуги сторонніх організацій</td><td class="num">14 434,00</td><td class="num">14 434,00</td><td class="num">0,00</td></tr>
    <tr><td>Оренда сторонньої техніки</td><td class="num">14 000,00</td><td class="num">14 000,00</td><td class="num">0,00</td></tr>
    <tr><td>Інші податки</td><td class="num">—</td><td class="num">1 239,98</td><td class="num bad">−1 239,98</td></tr>
    <tr class="total"><td>Списання разом</td><td class="num">{kzout}</td><td class="num">3 707 388,31</td><td class="num bad">+18 125,33</td></tr>
    <tr><td>Надходження (повернення коштів)</td><td class="num">{kzin}</td><td class="num">41 332,61</td><td class="num">0,00</td></tr>
  </tbody>
</table>

<p>Колонка «Факт» у бюджетному файлі — це зріз Казни, і по більшості статей вона збігається
<b>копійка в копійку</b>. Розбіжність лише дві: недобрано <b>19 365,31 грн зарплати</b>
і зайво показано <b>1 239,98 грн інших податків</b>, яких по цьому підрозділу в Казні немає.
Нетто — <b>18 125,33 грн</b>. Файл варто перезняти.</p>

<div class="warn"><b>Розрив «закуплено ↔ оплачено».</b> Прихід ТМЦ по ЕРП — <b>{fakt} грн</b>,
оплачено через Казну — <b>{kzm} грн</b>. Різниця <b>{kred} грн</b> — це неоплачена
кредиторська заборгованість перед постачальниками матеріалів.</div>

<h2>6. Чому бюджет не порівнюється з P&amp;L</h2>

<p>По підрозділу «МД IRS 2026» у P&amp;L за липень–жовтень визнано витрат лише
<b>43 779,60 грн</b> (інші виробничі 42 926,00 і ЗП 853,60). Статті «Будівельні матеріали»
в P&amp;L немає взагалі — <b>собівартість визнається в момент реалізації</b>, а будинки ще
не передані замовнику. Закуплені матеріали лежать у запасах і незавершеному виробництві.</p>

<p>Тому на цій стадії проєкту порівнювати бюджет із P&amp;L нема з чим. Робочі бази контролю —
<b>закупівлі</b> (нарахування, розділ 1–2) і <b>каса</b> (Казна, розділ 5). До P&amp;L
повернемось після першої реалізації будинків.</p>

<h2>7. Аудит загальних назв</h2>

<p>Прогін евристики по свіжих закупівлях (з попереднього прогону 26.07 факт зріс з 2,26 до 4,45 млн).
Головний висновок: <b>масових помилок у загальних назвах немає</b>. Те, що на перший погляд
виглядає як «закупівля без плану» (13 груп на 144 445,20 грн), насправді має іншу причину.</p>

<h3>7.1. Порожні рядки кошторису — 16 позицій (СС «15 м»)</h3>
<p>У картці СС «МД IRS 2026 15 м» є <b>16 рядків із нульовою кількістю і нульовою ціною</b>:
Плитка, Поручень, Плінтус на стільницю, Кутик і Заглушка на стільницю, Профіль для підлоги (2),
Муфта, Американка, Куточок сантехнічний, Заклепка, Штора, Захист від перенапруги,
Кабель нагріваючий, Замок, Канал вентиляційний.</p>
<p>По <b>13 із них уже закуплено на 144 445,20 грн</b>. Загальні назви на картках номенклатури
проставлені <b>вірно</b> — вони збігаються дослівно. Це не помилка НСІ, а <b>недозаповнений
кошторис</b>: позиції внесено, але не оцінено. Поки вони нульові, ці витрати назавжди
залишатимуться «перевитратою» і спотворюватимуть прогноз.</p>

<h3>7.2. Реальна помилка загальної назви — 4 картки</h3>
<p>У групі <b>«АВР»</b> (план 2 760 грн/шт — автомат введення резерву) лежать звичайні
модульні автомати:</p>
<table>
  <thead><tr><th>Код</th><th>Картка</th><th>Кіл.</th><th>Ціна факту</th><th>Сума</th><th>Стало ЗН</th></tr></thead>
  <tbody>
    <tr><td>00-00007600</td><td>Автоматичний вимикач 2р В 40А 6кА 2М</td><td class="num">6</td><td class="num">456,30</td><td class="num">2 737,80</td><td>Вимикач</td></tr>
    <tr><td>00-00000412</td><td>Автоматичний вимикач 1п 25А</td><td class="num">6</td><td class="num">71,21</td><td class="num">427,25</td><td>Вимикач</td></tr>
    <tr><td>00-00003558</td><td>Автоматичний вимикач 1п 16А</td><td class="num">6</td><td class="num">69,62</td><td class="num">417,74</td><td>Вимикач</td></tr>
    <tr><td>00-00003557</td><td>Автоматичний вимикач 1п 10А</td><td class="num">6</td><td class="num">68,47</td><td class="num">410,83</td><td>Вимикач</td></tr>
    <tr class="total"><td>Разом</td><td></td><td></td><td></td><td class="num">3 993,62</td><td></td></tr>
  </tbody>
</table>
<p>Обґрунтування: АВР — це апарат автоматичного введення резерву, у групі має лишитись
<b>тільки картка 00-00009969 «Автомат введення резерву 2п 63A»</b> (факт 2 808,00 проти
плану 2 760,00 — розбіжність 1,7 %). Решта чотирьох — модульні автомати захисту ліній;
у плані СС «15 м» вони закладені під ЗН <b>«Вимикач»</b> як «Модульний автомат 10А/16А/25А»
по 124,00–142,50 грн. <b>Правки в базу я не вносив</b> — проставити ЗН на картках потрібно вручну.</p>

<h3>7.3. Розбіжність методики між двома кошторисами</h3>
<p>Ті самі модульні автомати у СС «15 м» закладені під ЗН «Вимикач», а у СС «30 м» — під
ЗН «Диференційний автомат» (план 9 746,10 грн, закупівель нуль). Дифавтомат — інший апарат
(автомат із ПЗВ). Кошториси варто привести до одного довідника.</p>

<h3>7.4. Поза кошторисом</h3>
<p><b>Коліно дворострубне Profil 75 графіт</b> — 12 шт на 3 667,20 грн у групі «коліно»
(планова ціна 68,92 грн — каналізаційне коліно). Це елемент <b>водостічної системи Profil</b>
(разом з ним куплені ринва, труба водостічна, з'єднувач і держак ринви). Окремої позиції
для коліна водостоку в кошторисі немає — це питання до кошторису, а не до загальної назви.</p>

<h3>7.5. Завищені планові ціни (загальні назви вірні)</h3>
<p><b>Шайба</b>: план 19,10 грн/шт при фактичній ціні 0,47 грн (шайби М12) — план 17 304,60 грн
на 906 шт при реальній вартості близько 430 грн. <b>Розподільча коробка</b>: план 100,47 грн
проти факту 24,14 грн, план 32 754,00 грн на 326 шт. Призначення в обох випадках збігається,
загальні назви вірні — це помилка ціни в кошторисі. Групи «мшп», «Гайка», «Шпилька» перевірено
і відсіяно як хибні спрацювання (родові категорії та витратні матеріали).</p>

<h2>Висновки</h2>
<ol>
  <li>Матеріали йдуть з економією <b>{otkl} грн (3,8 %)</b>, але вона <b>зменшується</b>
      (тиждень тому було 444 523,90).</li>
  <li>Липень закрито з випередженням графіка; уся вага — <b>серпень, {ser} грн</b>,
      закупівель поки нуль.</li>
  <li>Критично не закуплено: <b>генератори 805 000 грн</b>, бойлери 192 400 грн, фарба 139 384 грн.</li>
  <li>Неоплачена кредиторка за матеріали — <b>{kred} грн</b>.</li>
  <li>Треба дозаповнити <b>16 нульових рядків кошторису «15 м»</b> — інакше 144 445,20 грн
      закупівель назавжди лишаться «поза планом».</li>
  <li>Виправити загальні назви на <b>4 картках автоматів</b> (3 993,62 грн) — список у п. 7.2.</li>
  <li>Перезняти колонку «Факт» у бюджетному файлі — розбіжність з Казною 18 125,33 грн.</li>
</ol>

<div class="att"><b>Додаток:</b>
<ol>
  <li>«Сверка_МД_IRS_2026_04-08-2026.xlsx» — 8 аркушів: зведення, 141 загальна назва,
      «ще закупити», кошторис vs СС, графік закупівель, бюджет vs каса, аудит ЗН,
      57 документів закупівлі.</li>
</ol></div>

<div class="src">Джерела: BaseERP — звіт «А_ПланФактныйПроизводствоПолный» (метод ПолучитьДанные),
регістр СебестоимостьТоваров (прихід / Приобретение товаров и услуг, підрозділ закупівель
в ієрархії), довідник А_СтруктураСебестоимости; Казна — регістр БДДС; звіт А_ОтчетPL;
файли «МАТЕРИАЛИ.xlsx» і «Бюджет_Серпень26-Жовтень26.xlsx». Усі суми з ПДВ.
Дані станом на 04.08.2026. Змін до баз не вносилось.</div>

</div>
</body>
</html>
"""

html = HTML.format(
    css=CSS,
    p15=g(4227672.06), p30=g(7069053.18), plan=g(t["ПланГрн"]),
    fakt=g(t["ФактГрн"]), faktp=g(t["ФактГрн"] / t["ПланГрн"] * 100, 1),
    ost=g(t["Осталось"]), ostp=g(t["Осталось"] / t["ПланГрн"] * 100, 1),
    prog=g(t["Прогноз"]), progp=g(t["Прогноз"] / t["ПланГрн"] * 100, 1),
    otkl=g(t["ОтклонениеПрогноз"]), otklp=g(t["ОтклонениеПрогноз"] / t["ПланГрн"] * 100, 1),
    neco=len(eco), ecos=g(eco_s), nper=len(per), pers=g(per_s),
    rows_eco=rows_eco, rows_per=rows_per, rows_buy=rows_buy,
    x15=g(xl["15"]["СумаВсього"]), x30=g(xl["30"]["СумаВсього"]),
    xtot=g(xl["15"]["СумаВсього"] + xl["30"]["СумаВсього"]),
    l15=g(xl["15"]["график"]["липень"][1]), lip=g(lip),
    s15=g(xl["15"]["график"]["серпень"][1]), s30=g(xl["30"]["график"]["серпень"][1]),
    ser=g(ser), dlip=g(t["ФактГрн"] - lip),
    x15s=g(xl["15"]["СумаВсього"]), x30s=g(xl["30"]["СумаВсього"]),
    gtot=g(lip + ser),
    kzm=g(kz_mat), kzzp=g(kz_zp), kzout=g(kz_out), kzin=g(kz_in),
    kred=g(t["ФактГрн"] - kz_mat),
)

path = os.path.join(OUT, "Лист_МД_IRS_2026_04-08-2026.html")
open(path, "w", encoding="utf-8").write(html)
print("OK ->", path, len(html), "символів")
