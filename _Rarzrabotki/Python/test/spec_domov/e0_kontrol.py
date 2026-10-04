# -*- coding: utf-8 -*-
"""Этап 0 (§10.1): сверка дампа с живой конфигурацией, планы обмена, регистрация СС, роли и пользователи — только чтение.
Выход out/e0_kontrol_<stamp>.json; раздел exchange.changes — справочно (состояние регистрации СС на день этапа 0). Приёмки
Task 28 и 44 сверяют регистрацию НЕ с ним, а начало и конец своего прогона: между днями таблицу меняют обмен и правки
пользователей (память feedback_invariantnye_testy_ne_baseline)."""
import collections
import datetime
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import preflight_config  # noqa: E402
from common import BB_MAIN, BB_WT, HERE, OUT, connect, q  # noqa: E402

STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M")
NEW_NAMES = [n for _, n in preflight_config.EXPECTED_NEW]


def main():
    B = connect()
    md = B.Metadata
    res = {"stamp": STAMP}
    db = preflight_config.db_objects(B)
    for nm, root in (("main", BB_MAIN), ("worktree", BB_WT)):
        f = {x for x in preflight_config.file_objects(os.path.join(root, "Configuration.xml")) if x[0] in {k for k, _ in db}}
        res[f"configuration_{nm}_vs_db"] = {"file": len(f), "db": len(db),
                                           "db_minus_file": sorted(f"{k}.{n}" for k, n in db - f),
                                           "file_minus_db": sorted(f"{k}.{n}" for k, n in f - db)}
    r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, "dump_check.py"), "--tag", "e0", "--objects",
                        "Справочник.СтруктураСебестоимости,Справочник.СтруктураСебестоимости.Форма.ФормаЭлемента,"
                        "Документ.РасчетКомплектаций,ПланОбмена.ОбменУправлениеПредприятиемБухгалтерия20"],
                       capture_output=True, text=True, encoding="utf-8")
    res["dump_vs_main"] = {"exit": r.returncode, "tail": r.stdout[-600:]}
    res["new_names_free"] = {n: (md.Catalogs.Find(n) is None and md.Enums.Find(n) is None and md.CommonModules.Find(n) is None)
                             for n in NEW_NAMES}
    probe = {"СтруктураСебестоимости": md.Catalogs.Find("СтруктураСебестоимости"),
             "РасчетКомплектаций": md.Documents.Find("РасчетКомплектаций")}
    plans = {k: [] for k in probe}
    for i in range(md.ExchangePlans.Count()):
        ep = md.ExchangePlans.Get(i)
        for k, m in probe.items():
            if ep.Content.Contains(m):
                plans[k].append(ep.Name)
    chg = q(B, """ВЫБРАТЬ
	Изм.Ссылка.Код КАК Код,
	ПРЕДСТАВЛЕНИЕ(Изм.Узел) КАК Узел,
	ЕСТЬNULL(Изм.НомерСообщения, -1) КАК НомерСообщения
ИЗ
	Справочник.СтруктураСебестоимости.Изменения КАК Изм
УПОРЯДОЧИТЬ ПО
	Код""")
    res["exchange"] = {"plans_with": plans,
                       "changes": [{"code": c["Код"], "node": c["Узел"], "msg": int(c["НомерСообщения"])} for c in chg]}
    cards = q(B, """ВЫБРАТЬ
	КОЛИЧЕСТВО(Компл.НомерСтроки) КАК Строк,
	СУММА(ВЫБОР КОГДА Компл.Номенклатура = ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка) ТОГДА 0 ИНАЧЕ 1 КОНЕЦ) КАК СКарточкой
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код""", Код="000000003")[0]
    res["ss003_cards"] = {"rows": int(cards["Строк"]), "with_card": int(cards["СКарточкой"])}
    roles = collections.Counter()
    users = B.ПользователиИнформационнойБазы.ПолучитьПользователей()
    for i in range(users.Количество()):
        u = users.Получить(i)
        names = {B.String(role_md.Имя) for role_md in u.Роли}    # КоллекцияРолейПользователя: только перебор, без Получить()
        key = ("ПолныеПрава" if "ПолныеПрава" in names else
               "+".join(sorted(names & {"ДобавлениеИзменениеДанныхБухгалтерии", "ЧтениеДанныхБухгалтерии", "ЗапускТонкогоКлиента"})) or "прочие")
        roles[key] += 1
    res["users_by_roles"] = dict(roles)
    write_ss = []
    for i in range(md.Roles.Count()):
        role = md.Roles.Get(i)
        if B.ПравоДоступа("Изменение", probe["СтруктураСебестоимости"], role):
            write_ss.append(role.Name)
    res["roles_write_ss"] = write_ss
    res["decisions"] = {"Д1": "правки в worktree на синхронизированных копиях, загрузка из основного каталога",
                        "Д2": "права (б), решение пользователя 04.10.2026: ЗапускТонкогоКлиента — полный набор (как у СС, правят все 24); ДобавлениеИзменениеДанныхБухгалтерии — полный набор; ЧтениеДанныхБухгалтерии — чтение",
                        "Д3": "«паков» не сводится к «пак»", "Т9": "Деталь без карточки — только в Чернетке",
                        "Т10": "зелёное «единственная сирота этапа»", "Т11": "узел-позиция ищется и по похожему тексту"}
    res["ok"] = {"main_eq_db": not res["configuration_main_vs_db"]["db_minus_file"]
                 and not res["configuration_main_vs_db"]["file_minus_db"],
                 "dump_eq_main": r.returncode == 0, "names_free": all(res["new_names_free"].values()),
                 "ss003_214": res["ss003_cards"] == {"rows": 214, "with_card": 214},
                 "rk_not_in_exchange": plans["РасчетКомплектаций"] == []}
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e0_kontrol_{STAMP}.json")
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({k: res[k] for k in ("ok", "users_by_roles", "roles_write_ss", "ss003_cards")}, ensure_ascii=False, indent=1))
    print("->", out)
    return 0 if all(res["ok"].values()) else 1


if __name__ == "__main__":
    sys.exit(main())
