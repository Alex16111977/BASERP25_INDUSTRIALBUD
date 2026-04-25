"""Render pl_faq.md from _pl_aggregates.json — prose-heavy for NotebookLM RAG."""
import json
import sys
from pathlib import Path

AGG = Path(__file__).parent / "_pl_aggregates.json"
OUT = Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_PL\pl_faq.md")

UA_MONTHS = {
    "2025-12": "Грудень 2025",
    "2026-01": "Січень 2026",
    "2026-02": "Лютий 2026",
}

PERIOD_ORDER = ["2025-12", "2026-01", "2026-02"]


def fmt(x):
    if x is None or abs(x) < 0.01:
        return "—"
    neg = x < 0
    a = abs(x)
    i = int(a)
    f = a - i
    s = f"{i:,}".replace(",", " ")
    if f > 0.005:
        s += f".{int(round(f*100)):02d}"
    return ("−" + s) if neg else s


def num_anchor(x):
    """Semantic anchors for a monetary value — lets RAG match on words like
    'млрд', 'мільйонів' etc. even if the user doesn't type the exact digits.
    """
    if x is None or abs(x) < 0.01:
        return ""
    ax = abs(x)
    parts = []
    if ax >= 1_000_000_000:
        parts.append(f"~{ax/1_000_000_000:.2f} млрд ₴")
        parts.append(f"{int(ax/1_000_000):,} млн ₴".replace(",", " "))
    elif ax >= 1_000_000:
        parts.append(f"~{ax/1_000_000:.2f} млн ₴")
        whole_m = int(round(ax / 1_000_000))
        parts.append(f"приблизно {whole_m} мільйонів ₴")
    elif ax >= 1_000:
        parts.append(f"~{ax/1_000:.0f} тисяч ₴")
    else:
        parts.append(f"{int(ax)} ₴")
    return ", ".join(parts)


def ord_word(n):
    """Порядкове слово: 1 → 'Перший', 2 → 'Другий' ... для QA-заголовків."""
    words = {1: "Перший", 2: "Другий", 3: "Третій", 4: "Четвертий", 5: "П'ятий"}
    return words.get(n, f"{n}-й")


def dump_file_for(ym, bucket):
    """Повертає шлях на pl_dump_*.md для посилань.
    bucket ∈ {'summary','income','cost','opex','marketing','cash'}
    """
    YM = ym.replace("-", "_")
    mapping = {
        "summary": f"pl_dump_{YM}_01_summary.md",
        "income":  f"pl_dump_{YM}_02_income_revenue.md",
        "cost":    f"pl_dump_{YM}_03_cost_of_goods.md",
        "opex":    f"pl_dump_{YM}_04_opex_admin.md",
        "marketing": f"pl_dump_{YM}_05_marketing_fin.md",
        "cash":    f"pl_dump_{YM}_06_cash_anomalies.md",
    }
    return mapping[bucket]


def _render_sec11(out, data):
    """Sec 11 — Precision Q&A для фінансистів.
    ~115 питань-відповідей з точними цифрами у 3 форматах (точна ₴, округлена, пропис)
    для максимального retrieval-матчу у NotebookLM RAG.
    """
    out.append("## 11. Точні суми за запитом (precision Q&A)")
    out.append("")
    out.append(
        "Цей розділ — **115+ готових питань-відповідей** із **точними сумами** для швидкого retrieval-матчу. "
        "Кожна відповідь дублює ключове число у **3 форматах**: жирна точна цифра (₴), округлення у млн/млрд, "
        "і текст-якір для пошуку словами («приблизно 193 мільйони», «~1.69 млрд»). Це критично для фінансових "
        "запитів де помилка у класифікації коштує реальних грошей. **Завжди звіряйте з джерелом** — воно вказано у кожній відповіді."
    )
    out.append("")

    # ─── 11.1 TL;DR цифри (4 × 3 = 12 Q&A) ───
    out.append("### 11.1 Ключові цифри періоду — TL;DR точні значення")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        t = data[ym]["totals"]
        src_sum = dump_file_for(ym, "summary")
        src_cash = dump_file_for(ym, "cash")

        out.append(f"#### Q: Скільки План PL склав у {label} ТОВ ІНДАСТРІАЛБУД?")
        out.append("")
        out.append(
            f"**Відповідь:** **{fmt(t['plan'])} ₴** ({num_anchor(t['plan'])}). "
            f"Сумарний плановий бюджет усіх ЦО-підрозділів групи ІНДАСТРІАЛБУД за {label}, "
            f"режим ВключатьДочерние=Ложь (тільки ЦО-документи, без СВОД, щоб уникнути задвоєння). "
            f"Кількість ЦО-документів А_ОтчетPL: **{t['n_plan_docs']}**. "
            f"Джерело: `{src_sum}`, секція TL;DR + Contextual metadata."
        )
        out.append("")

        out.append(f"#### Q: Скільки Факт витрат ЕРП склав у {label}?")
        out.append("")
        out.append(
            f"**Відповідь:** **{fmt(t['fact_rash'])} ₴** ({num_anchor(t['fact_rash'])}). "
            f"Фактичні нарахування за {label} з регістра **ПрочиеРасходы** (метод нарахування, не каса). "
            f"Охоплює **{t['n_rash_rows']}** рядків-регістраторів (документи Приобретение, Акт, РКО, Списание). "
            f"Джерело: `{src_sum}` (TL;DR) і розбивка по статтях у "
            f"`{dump_file_for(ym,'cost')}` / `{dump_file_for(ym,'opex')}` / `{dump_file_for(ym,'marketing')}`."
        )
        out.append("")

        out.append(f"#### Q: Який касовий приплив у Казну за {label}?")
        out.append("")
        out.append(
            f"**Відповідь:** **{fmt(t['cash_in'])} ₴** ({num_anchor(t['cash_in'])}). "
            f"Сума всіх рухів 'Приход' у регістрі **А_ДвиженияДенегИзКазны** за {label}. "
            f"Кількість рядків руху (в обидва боки): **{t['n_cash_rows']}**. "
            f"**Увага:** включає **всі рухи групи ІНДАСТРІАЛБУД**, у т.ч. внутрішньогрупові перекази між підрозділами. "
            f"Не всі з них прив'язані до PL-статей (див. «Каса БЕЗ PL» у `{src_cash}`)."
        )
        out.append("")

        out.append(f"#### Q: Який касовий відплив з Казни за {label}?")
        out.append("")
        out.append(
            f"**Відповідь:** **{fmt(t['cash_out'])} ₴** ({num_anchor(t['cash_out'])}). "
            f"Сума всіх рухів 'Расход' у регістрі А_ДвиженияДенегИзКазны за {label}. "
            f"Розходження між припливом і відпливом відображає зміну залишків на рахунках/касах і очікувано. "
            f"Джерело: `{src_cash}`, секція «Касовий відплив топ-документів»."
        )
        out.append("")

    out.append("---")
    out.append("")

    # ─── 11.2 Топ-5 підрозділів за планом (5 × 3 = 15 Q&A) ───
    out.append("### 11.2 Топ-5 підрозділів за плановим бюджетом PL")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_podr_plan"][:5]
        src = dump_file_for(ym, "summary")
        for i, p in enumerate(top, 1):
            ow = ord_word(i)
            out.append(f"#### Q: {ow} підрозділ за плановим бюджетом PL у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **{p['подр']}** із плановим бюджетом **{fmt(p['план'])} ₴** "
                f"({num_anchor(p['план'])}). Це **№{i}** у ренкінгу підрозділів за {label} "
                f"(ВключатьДочерние=Ложь, чистий ЦО-план без задвоєння). "
                f"Джерело: `{src}`, розділ «Топ-10 підрозділів за плановим бюджетом PL»."
            )
            out.append("")
    out.append("---")
    out.append("")

    # ─── 11.3 Топ-5 контрагентів за витратами (5 × 3 = 15 Q&A) ───
    out.append("### 11.3 Топ-5 контрагентів-постачальників за сумою витрат")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_kontr_rash"][:5]
        src_cost = dump_file_for(ym, "cost")
        for i, k in enumerate(top, 1):
            ow = ord_word(i)
            statti = ", ".join(k["статті"][:3]) if k.get("статті") else "—"
            out.append(f"#### Q: {ow} контрагент за витратами у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **{k['контрагент']}** — **{fmt(k['сума'])} ₴** "
                f"({num_anchor(k['сума'])}). Це **№{i}** за сумою фактичних витрат у {label}. "
                f"PL-статті у які підпадають витрати цього контрагента: {statti}. "
                f"Джерело: `{src_cost}` (якщо будівельні матеріали/послуги) або "
                f"`{dump_file_for(ym,'opex')}` (якщо адмін/загальновиробничі)."
            )
            out.append("")
    out.append("---")
    out.append("")

    # ─── 11.4 Топ-5 статей за фактом (5 × 3 = 15 Q&A) ───
    out.append("### 11.4 Топ-5 PL-статей за фактом витрат ЕРП")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_art_fact"][:5]
        for i, a in enumerate(top, 1):
            ow = ord_word(i)
            out.append(f"#### Q: {ow} за фактом PL-стаття у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **«{a['стаття']}»** з фактичною сумою **{fmt(a['факт'])} ₴** "
                f"({num_anchor(a['факт'])}). Це **№{i}** у ренкінгу PL-статей за фактом витрат ЕРП (регістр ПрочиеРасходы) у {label}. "
                f"Деталі по документам, контрагентам і підрозділам — у "
                f"`{dump_file_for(ym,'cost')}` або `{dump_file_for(ym,'opex')}` "
                f"(залежно від PL-групи статті — собівартість vs адмін/загальновиробн.)."
            )
            out.append("")
    out.append("---")
    out.append("")

    # ─── 11.5 Топ-5 документів-витрат (5 × 3 = 15 Q&A) ───
    out.append("### 11.5 Топ-5 документів-регістраторів витрат за сумою")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_docs_rash"][:5]
        for i, r in enumerate(top, 1):
            ow = ord_word(i)
            contr = r["контрагент"] if r.get("контрагент") else "(контрагент не вказаний)"
            num = r["номер"] if r.get("номер") else "(б/н)"
            date = r["дата"] if r.get("дата") and r["дата"] != "0100-01-01" else "(дата сервісна)"
            out.append(f"#### Q: {ow} найбільший документ-витрата у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **{r['тип']} №{num}** від {date}, контрагент **{contr}**, "
                f"підрозділ **{r['підр']}**, сума **{fmt(r['сума'])} ₴** "
                f"({num_anchor(r['сума'])}). PL-стаття через ДДС «{r['ддс']}». "
                f"Це **№{i}** серед документів-витрат {label}. "
                f"Джерело: `{dump_file_for(ym,'cost')}` або `{dump_file_for(ym,'opex')}`, розділ «Топ документи»."
            )
            out.append("")
    out.append("---")
    out.append("")

    # ─── 11.6 Топ-5 касових припливів (5 × 3 = 15 Q&A) ───
    out.append("### 11.6 Топ-5 касових припливів у Казну")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_cash_in"][:5]
        src_cash = dump_file_for(ym, "cash")
        for i, r in enumerate(top, 1):
            ow = ord_word(i)
            contr = r["контрагент"] if r.get("контрагент") else "(платник не вказаний, внутрішній перенос)"
            num = r["номер"] if r.get("номер") else "(б/н)"
            date = r["дата"] if r.get("дата") and r["дата"] != "0100-01-01" else "(сервіс)"
            out.append(f"#### Q: {ow} найбільший касовий приплив у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **{r['тип']} №{num}** від {date}, платник **{contr}**, "
                f"підрозділ **{r['підр']}**, сума **{fmt(r['сума'])} ₴** "
                f"({num_anchor(r['сума'])}). ДДС-стаття «{r['ддс']}». "
                f"Це **№{i}** серед припливів каси {label}. "
                f"Джерело: `{src_cash}`, розділ «Касовий приплив топ-документів»."
            )
            out.append("")
    out.append("---")
    out.append("")

    # ─── 11.7 Топ-5 PL-статей за планом (5 × 3 = 15 Q&A) ───
    out.append("### 11.7 Топ-5 PL-статей за плановим бюджетом")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_art_plan"][:5]
        src = dump_file_for(ym, "summary")
        for i, a in enumerate(top, 1):
            ow = ord_word(i)
            coms = a.get("коменти") or []
            ct = f" Коментар фінансиста з Excel: «{coms[0][:150]}»." if coms else ""
            out.append(f"#### Q: {ow} за плановим бюджетом PL-стаття у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **«{a['стаття']}»** із плановим бюджетом **{fmt(a['план'])} ₴** "
                f"({num_anchor(a['план'])}).{ct} Це **№{i}** серед PL-статей за плановим бюджетом {label} "
                f"(агреговано по всіх ЦО-документах). Джерело: `{src}`, розділ «Топ PL-статей за планом»."
            )
            out.append("")
    out.append("---")
    out.append("")

    # ─── 11.8 Delta між місяцями (4 × 2 = 8 Q&A) ───
    out.append("### 11.8 Порівняння між місяцями (delta)")
    out.append("")
    pairs = [("2025-12", "2026-01", "Січень 2026 vs Грудень 2025"),
             ("2026-01", "2026-02", "Лютий 2026 vs Січень 2026")]
    for prev_ym, curr_ym, label in pairs:
        t_prev = data[prev_ym]["totals"]
        t_curr = data[curr_ym]["totals"]
        delta_file = "pl_dump_delta_jan2026_vs_dec2025.md" if prev_ym == "2025-12" else "pl_dump_delta_feb2026_vs_jan2026.md"
        metrics = [
            ("План PL", "plan"),
            ("Факт витрат ЕРП", "fact_rash"),
            ("Каса приплив", "cash_in"),
            ("Каса відплив", "cash_out"),
        ]
        for metric_label, key in metrics:
            d_prev = t_prev[key]
            d_curr = t_curr[key]
            delta = d_curr - d_prev
            pct = (delta / d_prev * 100) if abs(d_prev) >= 0.01 else 0.0
            direction = "зріс" if delta > 0 else "впав"
            sign = "+" if delta > 0 else ""
            out.append(f"#### Q: Як змінився показник «{metric_label}» у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **{direction} на {fmt(abs(delta))} ₴** ({num_anchor(abs(delta))}), "
                f"тобто **{sign}{pct:.1f}%** у відносному вимірі. "
                f"Було (поперед. період): **{fmt(d_prev)} ₴**. Стало: **{fmt(d_curr)} ₴**. "
                f"Джерело: `{delta_file}`, секція Top moves."
            )
            out.append("")
    out.append("---")
    out.append("")

    # ─── 11.9 Ключові факти / тренди (5 Q&A) ───
    out.append("### 11.9 Ключові факти і тренди за 3 місяці (Груд 2025 – Лют 2026)")
    out.append("")

    # Total across 3 months
    total_plan = sum(data[ym]["totals"]["plan"] for ym in PERIOD_ORDER)
    total_fact = sum(data[ym]["totals"]["fact_rash"] for ym in PERIOD_ORDER)
    total_cin  = sum(data[ym]["totals"]["cash_in"] for ym in PERIOD_ORDER)
    total_cout = sum(data[ym]["totals"]["cash_out"] for ym in PERIOD_ORDER)
    total_ndocs = sum(data[ym]["totals"]["n_plan_docs"] for ym in PERIOD_ORDER)

    out.append("#### Q: Скільки всього План PL склав за 3 місяці (Грудень 2025 – Лютий 2026)?")
    out.append("")
    out.append(
        f"**Відповідь:** **{fmt(total_plan)} ₴** ({num_anchor(total_plan)}). "
        f"Сума планів трьох місяців: {fmt(data['2025-12']['totals']['plan'])} + "
        f"{fmt(data['2026-01']['totals']['plan'])} + {fmt(data['2026-02']['totals']['plan'])}. "
        f"Разом {total_ndocs} ЦО-документів А_ОтчетPL. "
        f"Джерела: `pl_dump_2025_12_01_summary.md`, `pl_dump_2026_01_01_summary.md`, `pl_dump_2026_02_01_summary.md`."
    )
    out.append("")

    out.append("#### Q: Скільки всього Факт витрат ЕРП за 3 місяці?")
    out.append("")
    out.append(
        f"**Відповідь:** **{fmt(total_fact)} ₴** ({num_anchor(total_fact)}). "
        f"Фактичні витрати (регістр ПрочиеРасходы) за період грудень 2025 – лютий 2026. "
        f"Джерело: summary-файли кожного місяця."
    )
    out.append("")

    out.append("#### Q: Скільки всього касового припливу за 3 місяці?")
    out.append("")
    out.append(
        f"**Відповідь:** **{fmt(total_cin)} ₴** ({num_anchor(total_cin)}). "
        f"Сумарний рух 'Приход' у А_ДвиженияДенегИзКазны включає внутрішньогрупові перекази. "
        f"Для чистих зовнішніх надходжень треба відняти статті «Транзит Вх» і «Внутригрупповые». "
        f"Джерело: cash_anomalies-файли кожного місяця."
    )
    out.append("")

    out.append("#### Q: Який підрозділ стабільно лідирує за плановим бюджетом PL?")
    out.append("")
    # Check Globino-2 is top in all 3 months
    leaders = [data[ym]["top_podr_plan"][0]["подр"] for ym in PERIOD_ORDER]
    if all(l == leaders[0] for l in leaders):
        leader = leaders[0]
        sums = {ym: data[ym]["top_podr_plan"][0]["план"] for ym in PERIOD_ORDER}
        out.append(
            f"**Відповідь:** **{leader}** — лідирує за плановим бюджетом PL у **всіх 3 місяцях**: "
            f"Грудень 2025 — {fmt(sums['2025-12'])} ₴, Січень 2026 — {fmt(sums['2026-01'])} ₴, "
            f"Лютий 2026 — {fmt(sums['2026-02'])} ₴. Разом за 3 міс — {fmt(sum(sums.values()))} ₴ "
            f"({num_anchor(sum(sums.values()))}). "
            f"Джерело: розділ 11.2 цього FAQ та summary-файли."
        )
    else:
        out.append(f"**Відповідь:** У кожному місяці свій лідер: {', '.join(f'{UA_MONTHS[ym]}: {l}' for ym, l in zip(PERIOD_ORDER, leaders))}.")
    out.append("")

    out.append("#### Q: Яка PL-стаття стабільно найбільша за фактом витрат?")
    out.append("")
    art_leaders = [data[ym]["top_art_fact"][0]["стаття"] for ym in PERIOD_ORDER]
    if all(a == art_leaders[0] for a in art_leaders):
        leader = art_leaders[0]
        sums = {ym: data[ym]["top_art_fact"][0]["факт"] for ym in PERIOD_ORDER}
        out.append(
            f"**Відповідь:** **«{leader}»** — найбільша за фактом у **всіх 3 місяцях**: "
            f"Грудень 2025 — {fmt(sums['2025-12'])} ₴, Січень 2026 — {fmt(sums['2026-01'])} ₴, "
            f"Лютий 2026 — {fmt(sums['2026-02'])} ₴. Це очікувано для будівельної компанії — "
            f"купівля стройматеріалів (цемент, арматура, метал, ізоляція) складає найбільшу частку витрат. "
            f"Джерело: розділ 11.4 цього FAQ."
        )
    else:
        out.append(f"**Відповідь:** У кожному місяці свій лідер: {', '.join(f'{UA_MONTHS[ym]}: «{a}»' for ym, a in zip(PERIOD_ORDER, art_leaders))}.")
    out.append("")

    out.append("---")
    out.append("")


def main():
    data = json.loads(AGG.read_text(encoding="utf-8"))
    out = []
    out.append("# PL FAQ — Прямі відповіді фінансистам ІНДАСТРІАЛБУД")
    out.append("")
    out.append("Дата генерації: 2026-04-22  ")
    out.append("Тип: ЗМІННИЙ (перегенеровується після закриття кожного місяця)  ")
    out.append("Призначення: прозові відповіді на типові фінансистські запитання, спеціально побудовано для NotebookLM RAG (короткі прозові блоки + малі таблиці).  ")
    out.append("Пов'язані файли: pl_dump_*.md (детальні виписки по місяцях), pl_methodology.md (методологія), pl_dds_mapping.md (маппінг).")
    out.append("")
    out.append("## Як читати цей файл")
    out.append("")
    out.append("Кожний розділ нижче починається з короткої прозової відповіді з конкретними числами. Там же — посилання на файл-джерело, де повні деталі. Ключова перевага: NotebookLM індексує прозу і retrievable навіть коли великі таблиці у місячних виписках не повертаються.")
    out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 1: Ключові підсумки за місяць ═══
    out.append("## 1. Ключові підсумки за період (TL;DR)")
    out.append("")

    for ym in PERIOD_ORDER:
        d = data[ym]
        t = d["totals"]
        label = UA_MONTHS[ym]
        out.append(f"### 1.{PERIOD_ORDER.index(ym)+1} {label}")
        out.append("")
        out.append(
            f"За {label} загальний **План PL** по ТОВ ІНДАСТРІАЛБУД склав **{fmt(t['plan'])} ₴** "
            f"(сумарно по {t['n_plan_docs']} документах А_ОтчетPL у режимі ВключатьДочерние=Ложь). "
            f"Загальний **Факт витрат ЕРП** (регістр ПрочиеРасходы) — **{fmt(t['fact_rash'])} ₴** "
            f"({t['n_rash_rows']} рядків-регістраторів). "
            f"Грошовий потік з Казни: **приплив {fmt(t['cash_in'])} ₴**, **відплив {fmt(t['cash_out'])} ₴** "
            f"({t['n_cash_rows']} рядків руху). Джерело: `pl_dump_{ym.replace('-', '_')}_01_summary.md`, секція TL;DR."
        )
        out.append("")
        out.append(f"**Ключові цифри {label}** — для швидкого довідування:")
        out.append("")
        out.append("| Показник | Сума ₴ |")
        out.append("|---|---|")
        out.append(f"| План PL (всі підрозділи) | {fmt(t['plan'])} |")
        out.append(f"| Факт витрат ЕРП | {fmt(t['fact_rash'])} |")
        out.append(f"| Каса приплив | {fmt(t['cash_in'])} |")
        out.append(f"| Каса відплив | {fmt(t['cash_out'])} |")
        out.append(f"| Документів планів | {t['n_plan_docs']} |")
        out.append("")

    out.append("**Важливо:** касові суми високі бо включають **всі рухи групи ІНДАСТРІАЛБУД** (включно з внутрішньогруповими переказами). Не всі з них прив'язані до PL (див. розділ 'Каса БЕЗ PL' у pl_dump_*.md).")
    out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 2: Топ підрозділів за планом ═══
    out.append("## 2. Топ підрозділів за плановим бюджетом PL")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_podr_plan"][:10]
        out.append(f"### 2.{PERIOD_ORDER.index(ym)+1} {label}")
        out.append("")
        if top:
            leader = top[0]
            followups = ", ".join(p["подр"] + " (" + fmt(p["план"]) + " ₴)" for p in top[1:5])
            out.append(
                f"У {label} найбільший плановий бюджет мав підрозділ **{leader['подр']}** — {fmt(leader['план'])} ₴. "
                f"Далі йдуть: {followups}."
            )
            out.append("")
            out.append("**Топ-10 підрозділів за планом:**")
            out.append("")
            out.append("| № | Підрозділ | План ₴ |")
            out.append("|---|---|---|")
            for i, p in enumerate(top, 1):
                out.append(f"| {i} | {p['подр']} | {fmt(p['план'])} |")
            out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 3: Топ контрагентів за витратами ═══
    out.append("## 3. Топ контрагенти-постачальники за сумою витрат")
    out.append("")
    out.append("Контрагенти, яким компанія найбільше платила за період (факт регістра ПрочиеРасходы, з документів-витрат типу Приобретение/Акт/РКО/Списание).")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_kontr_rash"][:10]
        out.append(f"### 3.{PERIOD_ORDER.index(ym)+1} {label}")
        out.append("")
        if top:
            leader = top[0]
            out.append(
                f"У {label} найбільша сума витрат (фактичне нарахування) пішла на **{leader['контрагент']}** — {fmt(leader['сума'])} ₴ "
                f"(у статтях: {', '.join(leader['статті'])}). Далі: "
                + "; ".join(f"**{k['контрагент']}** ({fmt(k['сума'])} ₴)" for k in top[1:5])
                + "."
            )
            out.append("")
            out.append("**Топ-10 контрагентів за витратами:**")
            out.append("")
            out.append("| № | Контрагент | Сума ₴ | У яких PL-статтях |")
            out.append("|---|---|---|---|")
            for i, k in enumerate(top, 1):
                out.append(f"| {i} | {k['контрагент']} | {fmt(k['сума'])} | {', '.join(k['статті'][:3])} |")
            out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 4: Топ статті за планом ═══
    out.append("## 4. Топ PL-статей за плановим бюджетом")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_art_plan"][:10]
        out.append(f"### 4.{PERIOD_ORDER.index(ym)+1} {label}")
        out.append("")
        if top:
            leader = top[0]
            lc = leader.get("коменти") or []
            lc_txt = f" Коментар фінансиста з Excel: «{lc[0]}»." if lc else ""
            out.append(
                f"У {label} найбільший план мала стаття **«{leader['стаття']}»** — {fmt(leader['план'])} ₴.{lc_txt} "
                + "Далі: "
                + ", ".join(f"«{a['стаття']}» ({fmt(a['план'])} ₴)" for a in top[1:5])
                + "."
            )
            out.append("")
            out.append("**Топ-10 статей за планом:**")
            out.append("")
            out.append("| № | PL-стаття | План ₴ | Коментар (з Excel, якщо є) |")
            out.append("|---|---|---|---|")
            for i, a in enumerate(top, 1):
                c = a.get("коменти") or []
                ct = f"«{c[0][:60]}»" if c else "—"
                out.append(f"| {i} | {a['стаття']} | {fmt(a['план'])} | {ct} |")
            out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 5: Топ статті за фактом ═══
    out.append("## 5. Топ PL-статей за фактом витрат ЕРП")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_art_fact"][:10]
        out.append(f"### 5.{PERIOD_ORDER.index(ym)+1} {label}")
        out.append("")
        if top:
            leader = top[0]
            out.append(
                f"За фактом ЕРП у {label} найбільше списано по **«{leader['стаття']}»** — {fmt(leader['факт'])} ₴. "
                + "Далі: "
                + ", ".join(f"«{a['стаття']}» ({fmt(a['факт'])} ₴)" for a in top[1:5])
                + "."
            )
            out.append("")
            out.append("**Топ-10 за фактом:**")
            out.append("")
            out.append("| № | PL-стаття | Факт ₴ |")
            out.append("|---|---|---|")
            for i, a in enumerate(top, 1):
                out.append(f"| {i} | {a['стаття']} | {fmt(a['факт'])} |")
            out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 6: Топ документи-витрати ═══
    out.append("## 6. Топ документів-регістраторів витрат за сумою")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_docs_rash"][:10]
        out.append(f"### 6.{PERIOD_ORDER.index(ym)+1} {label}")
        out.append("")
        if top:
            leader = top[0]
            out.append(
                f"Найбільший документ-витрата у {label} — **{leader['тип']} №{leader['номер']}** від {leader['дата']}, "
                f"контрагент **{leader['контрагент']}**, підрозділ {leader['підр']}, сума {fmt(leader['сума'])} ₴ "
                f"(PL стаття через ДДС «{leader['ддс']}»)."
            )
            out.append("")
            out.append("**Топ-10 документів-витрат:**")
            out.append("")
            out.append("| № | Тип | Номер | Дата | Контрагент | Підрозділ | Сума ₴ | ДДС |")
            out.append("|---|---|---|---|---|---|---|---|")
            for i, r in enumerate(top, 1):
                out.append(f"| {i} | {r['тип']} | {r['номер']} | {r['дата']} | {r['контрагент']} | {r['підр']} | {fmt(r['сума'])} | {r['ддс']} |")
            out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 7: Топ касових надходжень ═══
    out.append("## 7. Топ касових надходжень (приплив у Казну)")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        top = data[ym]["top_cash_in"][:10]
        out.append(f"### 7.{PERIOD_ORDER.index(ym)+1} {label}")
        out.append("")
        if top:
            leader = top[0]
            out.append(
                f"Найбільший касовий приплив у {label} — {fmt(leader['сума'])} ₴ через **{leader['тип']} №{leader['номер']}** "
                f"від {leader['дата']} від платника **{leader['контрагент']}** "
                f"(підрозділ {leader['підр']}, ДДС «{leader['ддс']}»)."
            )
            out.append("")
            out.append("**Топ-10 припливів каси:**")
            out.append("")
            out.append("| № | Тип | Номер | Дата | Платник | Підрозділ | Сума ₴ | ДДС |")
            out.append("|---|---|---|---|---|---|---|---|")
            for i, r in enumerate(top, 1):
                out.append(f"| {i} | {r['тип']} | {r['номер']} | {r['дата']} | {r['контрагент']} | {r['підр']} | {fmt(r['сума'])} | {r['ддс']} |")
            out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 8: Методологічні Q&A ═══
    out.append("## 8. Типові методологічні питання фінансиста")
    out.append("")
    out.append("### Q: Як класифікувати штраф від ДПС (податковий) чи штраф постачальнику?")
    out.append("")
    out.append("**Відповідь:** Штраф завжди записується у **PL-стаття 000000060 «Штраф»** (група «Административные затраты»). Мапається на ДДС «Штраф» (код УТ-001151). Це правило незалежне від суми чи природи штрафу — штрафи НЕ входять у себестойкість, бо це ризики управління, а не операційні витрати. Джерело: `pl_methodology.md` секція «Група «Административные затраты»».")
    out.append("")
    out.append("### Q: Купили основний засіб (обладнання, авто). У яку PL-статтю?")
    out.append("")
    out.append("**Відповідь:** У **PL-стаття 000000028 «Расход от инвест. Деятельности (ОС)»** (група «Дополнительные расходы»). Це **whitelist-стаття**: дозволена БЕЗ header-ДДС, бо це капітальні витрати, а не операційні. У ТЧ Статьи мапається на дві ДДС: «Покупка виробничих ОС» (УТ-001058) і «Прочие ОС» (УТ-001073). Джерело: `pl_dds_mapping.md`, `pl_methodology.md`.")
    out.append("")
    out.append("### Q: Що таке 'shared ДДС' і коли одна ДДС у кількох PL-статтях?")
    out.append("")
    out.append("**Відповідь:** 'Shared ДДС' — це виняток з правила уникальності. За замовчуванням кожна ДДС-стаття входить лише в одну PL-статтю. АЛЕ якщо ця ДДС за останні **12 місяців у регістрі ПрочиеРасходы не мала жодного обороту**, її можна додати до кількох PL-статей (zontik). Це правило адаптивної уникальності, перевіряє модуль об'єкта `Справочник.А_Статьи_PL.ПередЗаписью`. Приклад shared-zontik: ДДС «Персонал» використовується у статтях Премії (у виробничній і адмін групах). Джерело: `pl_dds_mapping.md` секція «Shared ДДС».")
    out.append("")
    out.append("### Q: Чому стаття 'Финансовые расходы' (код 000000056) не має ДДС?")
    out.append("")
    out.append("**Відповідь:** Це **whitelist-стаття**: відсотки по кредитах, курсові різниці і банківські збитки не прив'язуються до грошового потоку за ДДС-класифікацією. Всього таких whitelist-кодів два: **000000028** (інвест-діяльність, ОС) і **000000056** (фінансові витрати). Модуль об'єкта А_Статьи_PL дозволяє у цих двох випадках порожнє header-поле СтатьяДвиженияДенежныхСредств. Всі інші статті-Расход ЗОБОВ'ЯЗАНІ мати header-ДДС. Джерело: `pl_methodology.md`, `pl_articles_catalog.md`.")
    out.append("")
    out.append("### Q: У чому різниця між Excel-листом ЦО і СВОД для одного підрозділу?")
    out.append("")
    out.append("**Відповідь:** **ЦО-лист** (Центр Відповідальності) — план для ОДНОГО конкретного підрозділу, без дочірніх. Документ у 1С має `ВключатьДочерние=Ложь`. **СВОД-лист** — план для підрозділу ПЛЮС усі дочірні у цьому напрямку. Документ має `ВключатьДочерние=Істина`. Наприклад, для «Виробництва» у лютому 2026 є два документи: №000000096 (СВОД, ~20.3 млн ₴) і №000000097 (ЦО, ~1.2 млн ₴). Якщо підсумувати обидва — буде подвоєння. Звіт А_ОтчетPL за замовчуванням фільтрує тільки ЦО (ВключатьДочерние=Ложь) через СКД-параметр, щоб уникнути задвоєння. Джерело: `pl_report_architecture_analyst.md`, `pl_methodology.md`.")
    out.append("")
    out.append("### Q: Чому каса (Казна) розходиться з фактом витрат?")
    out.append("")
    out.append("**Відповідь:** Це нормальне і очікуване явище. Каса — **грошовий метод**, факт витрат — **метод нарахування**. Розходження легально в таких випадках: (1) **передоплата постачальнику**: гроші пішли з Казни але витрата ще не визнана у ПрочиеРасходы; (2) **постоплата**: витрату нарахували але гроші ще не відправили; (3) **перекриття боргів за минулі місяці**: гроші йдуть у цей період, а витрата була в попередньому. Збіг каси з фактом — виняток, розходження — правило. Джерело: `pl_report_architecture_analyst.md` секція «Коли касовий потік легально розходиться з витратами».")
    out.append("")
    out.append("### Q: Як зв'язуються PL-статті з фактичними витратами в ЕРП?")
    out.append("")
    out.append("**Відповідь:** Через **СтатьяДвиженияДенежныхСредств (ДДС)**, яка є посередником. Кожна PL-стаття типу 'Расход' має header-ДДС + ТЧ Статьи з переліком рядків-ДДС. У ЕРП витрати у регістрі `ПрочиеРасходы.СтатьяРасходов.А_СтатьяДвиженияДенежныхСредств` ідентифікують ту саму ДДС. Сумм. Схема: **Excel план** → **Документ.А_ОтчетPL** → **PL-стаття.Статьи.ДДС** ↔ **ПрочиеРасходы.СтатьяРасходов.ДДС** → **Факт**. Джерело: `pl_report_architecture_analyst.md`, `pl_dds_mapping.md`.")
    out.append("")
    out.append("### Q: Скільки PL-статей у довіднику? Скільки автоматично створених?")
    out.append("")
    out.append("**Відповідь:** У довіднику **Справочник.А_Статьи_PL** всього **68 статей** (усі елементи, немає груп — група хранится через окремий атрибут `Группа`). З них **7 автоматично створені** скриптом (`А_СозданоАвтоматически=Істина`): Внутригрупповые затраты (000000061), Зарплата механизаторов (000000062), Возврат денежных средств (000000064), Удержание с зарплаты механизаторов (000000065), Начисления на зарплату механизаторов (000000066), Зарплата кладовщика (000000067), Амортизация (000000068). Джерело: `pl_articles_catalog.md`.")
    out.append("")
    out.append("### Q: Які PL-групи існують і що вони означають економічно?")
    out.append("")
    out.append("**Відповідь:** 8 груп: (1) **Операционный доход** (код 000000006) — виручка від основної діяльності; (2) **Себестоимость** (000000001) — змінні виробничі витрати; (3) **Общепроизводственные затраты** (000000003) — інфраструктурні витрати; (4) **Маркетинговые затраты** (000000005) — реклама; (5) **Административные затраты** (000000002) — управлінський персонал, офіс, банки; (6) **Налоги и сборы** (000000008) — ПДВ і місцеві податки; (7) **Финансовая деятельность** (000000004) — відсотки, дивіденди, податок на прибуток; (8) **Дополнительные расходы** (000000007) — капітальні. Формула: **Виручка − Собівартість − Операційні − Маркетинг − Адмін − Податки − Фінансові ± Додаткові = Чистий PL**. Джерело: `pl_methodology.md` початок.")
    out.append("")
    out.append("### Q: Чому ЗП виробничого персоналу розділена на кілька статей (000000003, 000000004, 000000005)?")
    out.append("")
    out.append("**Відповідь:** Для чистоти обліку: **000000003 «ЗП производственного персонала»** = нетто (що реально видано на руки), **000000004 «Удержание...»** = ПДФО+ВЗ (утримано з брутто), **000000005 «Начисления...»** = ЄСВ 22% (сплачено зверху). Всі три разом складають повну gross-вартість праці. Аналогічна структура для ІТР (06/07/08), ПМ (41/42/43), адмін (38/39/40), механізаторів (62/65/66). Це дозволяє окремо планувати чисту зарплату, утримання і податки. Джерело: `pl_methodology.md` секція «Група Себестоимость».")
    out.append("")
    out.append("### Q: Чому у статті 000000022 «Прочие производственные затраты» три різні ДДС?")
    out.append("")
    out.append("**Відповідь:** Це контрольна стаття-буфер для витрат які не вкладаються в інші специфічні. У ТЧ Статьи вона мапається на три ДДС: УТ-001069 «Прочие производственные затраты» (header), УТ-001192 «Прочие затраты охраны труда», УТ-001052 «Прочие затраты». Це єдиний приклад 1:3 мапінгу в довіднику. Якщо ви бачите витрату у ній — це означає витрата з одної з цих ДДС. Джерело: `pl_dds_mapping.md` Forward view для 000000022.")
    out.append("")
    out.append("### Q: Як у звіті PL видно надходження від донорів (ООН, ВООЗ, ПРООН)?")
    out.append("")
    out.append("**Відповідь:** Надходження від міжнародних донорів (ООН, ВООЗ, ПРООН) приходять як **виручка** через Документи Реалізації або Акти виконаних робіт, на підрозділах **«МД ООН 2025»**, **«МД ВООЗ 2025»**, **«МД ПРООН Черкаси ДСНС»**. Мапається на PL-стаття 000000001 «Выручка от продаж» (група Операционный доход) через СтатьяДоходов «Выручка от продаж». Це ДОХІДНА частина — не плутати з внутрішніми витратами цих підрозділів. Джерело: `pl_dump_*.md` розділ відповідних підрозділів.")
    out.append("")
    out.append("### Q: Що таке 'ЕРП-витрати БЕЗ PL-прив'язки'? Чи це помилка?")
    out.append("")
    out.append("**Відповідь:** Це **DATA_GAP** — витрати з регістра ПрочиеРасходы, чиї ДДС не входять у жодну PL-статтю (не в ТЧ Статьи жодної). Не помилка коду, а сигнал для фінансиста: або (а) додати цю ДДС до відповідної PL-статті, або (б) виправити статтю витрат у первинному документі. У кожному місячному pl_dump_*.md є секція «Аномалії: ЕРП-витрати БЕЗ PL». У лютому 2026 таких 0 (усе класифіковано), у січні було невелика кількість (див. відповідний файл). Джерело: `pl_dump_*.md` секція аномалій.")
    out.append("")
    out.append("### Q: Внутрішньогрупові перекази — як вони видно у звіті?")
    out.append("")
    out.append("**Відповідь:** Внутрішньогрупові перекази (між компаніями групи ІНДАСТРІАЛБУД) ідуть у PL-статтю **000000061 «Внутригрупповые затраты»** (ДДС УТ-001042). Це автоматично створена стаття з лютого 2026. У касі такі перекази зазвичай помічено як «Каса БЕЗ PL» якщо їх ДДС інша (наприклад, технічні транзитні статті). Для аналізу реальних зовнішніх витрат рекомендується фільтрувати цю статтю. Джерело: `pl_methodology.md`, `pl_articles_catalog.md`.")
    out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 9: Довідкові питання-відповіді ═══
    out.append("## 9. Довідкові відповіді по підрозділах, статтях, сумах")
    out.append("")

    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        d = data[ym]
        t = d["totals"]
        out.append(f"### Q: Скільки всього витрат по компанії за {label}?")
        out.append("")
        out.append(f"**Відповідь:** Фактичні витрати (регістр ПрочиеРасходы) склали **{fmt(t['fact_rash'])} ₴** по {t['n_rash_rows']} документах-регістраторах. Планові витрати (з Документ.А_ОтчетPL) — **{fmt(t['plan'])} ₴** по {t['n_plan_docs']} ЦО-документах підрозділів. Касовий відплив з Казни — **{fmt(t['cash_out'])} ₴** (включно з внутрішньогруповими). Джерело: `pl_dump_{ym.replace('-', '_')}_*.md` секція TL;DR.")
        out.append("")

        # Top 3 подразделений
        top_podr = d["top_podr_plan"][:3]
        out.append(f"### Q: Який підрозділ мав найбільший план у {label}?")
        out.append("")
        if top_podr:
            out.append(
                f"**Відповідь:** {top_podr[0]['подр']} — план {fmt(top_podr[0]['план'])} ₴. "
                f"На другому місці {top_podr[1]['подр']} ({fmt(top_podr[1]['план'])} ₴), третє — {top_podr[2]['подр']} ({fmt(top_podr[2]['план'])} ₴). "
                f"Джерело: `pl_dump_{ym.replace('-', '_')}_*.md`, розріз за підрозділами."
            )
        out.append("")

        # Top 3 контрагенти
        top_kontr = d["top_kontr_rash"][:3]
        out.append(f"### Q: Топ-3 контрагенти за витратами у {label}?")
        out.append("")
        if top_kontr:
            out.append(
                f"**Відповідь:** (1) **{top_kontr[0]['контрагент']}** — {fmt(top_kontr[0]['сума'])} ₴ (статті: {', '.join(top_kontr[0]['статті'][:3])}); "
                f"(2) **{top_kontr[1]['контрагент']}** — {fmt(top_kontr[1]['сума'])} ₴ (статті: {', '.join(top_kontr[1]['статті'][:3])}); "
                f"(3) **{top_kontr[2]['контрагент']}** — {fmt(top_kontr[2]['сума'])} ₴ (статті: {', '.join(top_kontr[2]['статті'][:3])})."
            )
        out.append("")

        # Top 3 статті за фактом
        top_af = d["top_art_fact"][:3]
        out.append(f"### Q: Які 3 статті найбільше коштували ЕРП у {label}?")
        out.append("")
        if top_af:
            out.append(
                f"**Відповідь:** (1) **{top_af[0]['стаття']}** — {fmt(top_af[0]['факт'])} ₴; "
                f"(2) **{top_af[1]['стаття']}** — {fmt(top_af[1]['факт'])} ₴; "
                f"(3) **{top_af[2]['стаття']}** — {fmt(top_af[2]['факт'])} ₴. Джерело: `pl_dump_{ym.replace('-', '_')}_*.md`, секції окремих статей."
            )
        out.append("")

        # Найбільший документ
        top_docs = d["top_docs_rash"][:1]
        if top_docs:
            r = top_docs[0]
            out.append(f"### Q: Який найбільший документ-витрата у {label}?")
            out.append("")
            out.append(
                f"**Відповідь:** **{r['тип']} №{r['номер']}** від {r['дата']}, контрагент **{r['контрагент']}**, "
                f"підрозділ {r['підр']}, сума **{fmt(r['сума'])} ₴**, PL-стаття через ДДС «{r['ддс']}»."
            )
            out.append("")

    out.append("---")
    out.append("")
    out.append("## 10. Джерела для поглибленого аналізу")
    out.append("")
    out.append("**Помісячні виписки — 6 файлів на місяць × 3 місяці = 18 файлів** (split-за-bucket для кращого RAG-retrieval):")
    out.append("")
    for ym in PERIOD_ORDER:
        label = UA_MONTHS[ym]
        YM = ym.replace("-", "_")
        out.append(f"- **{label}:** `pl_dump_{YM}_01_summary.md` (TL;DR + топ-10), `pl_dump_{YM}_02_income_revenue.md` (доходи), `pl_dump_{YM}_03_cost_of_goods.md` (собівартість), `pl_dump_{YM}_04_opex_admin.md` (адмін/загальновиробн.), `pl_dump_{YM}_05_marketing_fin.md` (маркетинг+фінанси), `pl_dump_{YM}_06_cash_anomalies.md` (каса+аномалії).")
    out.append("")
    out.append("**Дельти між місяцями** (Top moves, нові/зниклі контрагенти): `pl_dump_delta_jan2026_vs_dec2025.md`, `pl_dump_delta_feb2026_vs_jan2026.md`.")
    out.append("")
    out.append("**Статичні довідники (ПОСТІЙНІ, не змінюються від місяця):**")
    out.append("")
    out.append("- **Каталог PL-статей** (повна ієрархія 68 статей, whitelist, автосоздані): `pl_articles_catalog.md`.")
    out.append("- **Маппінг PL ↔ ДДС** (матриця відповідностей, forward + reverse, shared-zontik): `pl_dds_mapping.md`.")
    out.append("- **Методологія** (що економічно означає кожна група/стаття, антикейси): `pl_methodology.md`.")
    out.append("- **Архітектура звіту для аналітика** (план vs факт vs каса, ЦО vs СВОД, data gap): `pl_report_architecture_analyst.md`.")
    out.append("")
    out.append("---")
    out.append("")

    # ═══ Sec 11: Точні суми за запитом (precision Q&A) ═══
    _render_sec11(out, data)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(out), encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
