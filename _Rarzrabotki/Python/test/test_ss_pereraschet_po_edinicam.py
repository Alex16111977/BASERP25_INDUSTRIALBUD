import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
c = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
M = c.А_СтруктураСебестоимостиСервер
Src = c.Перечисления.А_ИсточникиКоэффициентаПересчетаСС
fails=[]
def check(cond,msg):
    print(("OK   " if cond else "FAIL ")+msg)
    if not cond: fails.append(msg)
def snap(o):
    return [(r.НомерСтроки, round(r.Количество,3), round(r.Сумма,2), c.String(r.Единица), r.ЕдиницаСС, round(r.КоличествоСС_Оригинал,3), r.ЕдиницаСС_Оригинал) for r in o.Комплектующие]
def proto(p):
    print(f"  прот: Пересч={p.Пересчитано} НеНайд={p.НеНайдено} НетВК={p.НетВКомплектующих} Неодн={p.Неоднозначно} Σ {p.СуммаДо}->{p.СуммаПосле} Ош={p.Ошибка} {p.ТекстОшибки}")
    for i in range(p.Детали.Количество()): print("   ·", p.Детали.Получить(i))
def coef_row(o, nomss, ed):
    for r in o.КоэффициентыПересчета:
        if r.НоменклатураСС==nomss and r.ЕдиницаСС==ed: return r
ref = c.Справочники.СтруктураСебестоимости.НайтиПоНаименованию("МД ООН Админ", True)
NOM="Утеплювач базальтовий IZOVAT 100; стіни"
# --- T1: одна строка м3
o = ref.ПолучитьОбъект(); before = snap(o)
k3 = coef_row(o,NOM,"м3"); k3.Коэффициент=1.637931034; k3.ИсточникКоэффициента=Src.Вручную
p = M.ПересчитатьКоличество(o, k3.НомерСтроки); print("T1 рядок м3 №",k3.НомерСтроки); proto(p)
after = snap(o)
ch = [(b,a) for b,a in zip(before,after) if b!=a]
for b,a in ch: print("  изм:",b,"->",a)
check(len(ch)==1 and ch[0][1][0]==34 and ch[0][1][1]==12.666, "T1 изменена только строка 34 -> 12.666")
check(abs(p.СуммаДо-p.СуммаПосле)<0.001 and not p.Ошибка, "T1 Σ Сумма неизменна")
# --- T2: одна строка м2
o = ref.ПолучитьОбъект(); before = snap(o)
k2 = coef_row(o,NOM,"м2"); k2.Коэффициент=0.163987138; k2.ИсточникКоэффициента=Src.Вручную
p = M.ПересчитатьКоличество(o, k2.НомерСтроки); print("T2 рядок м2 №",k2.НомерСтроки); proto(p)
after = snap(o); ch=[(b,a) for b,a in zip(before,after) if b!=a]
for b,a in ch: print("  изм:",b,"->",a)
exp=round(66.6*0.163987138,3)
check(len(ch)==1 and ch[0][1][0]==239 and ch[0][1][1]==exp, f"T2 изменена только строка 239 -> {exp}")
# --- T3: одна строка с k=0 — ничего не меняется
o = ref.ПолучитьОбъект(); before = snap(o)
k2 = coef_row(o,NOM,"м2")
p = M.ПересчитатьКоличество(o, k2.НомерСтроки); print("T3 k=0"); proto(p)
check(snap(o)==before and p.Пересчитано==0, "T3 k=0: ничего не изменено")
# --- T4: полный пересчёт с обоими k — обе строки, неоднозначно=0 по IZOVAT
o = ref.ПолучитьОбъект()
k3 = coef_row(o,NOM,"м3"); k3.Коэффициент=1.637931034
k2 = coef_row(o,NOM,"м2"); k2.Коэффициент=0.163987138
p = M.ПересчитатьКоличество(o); print("T4 полный"); proto(p)
d={r.НомерСтроки:round(r.Количество,3) for r in o.Комплектующие}
check(d[34]==12.666 and d[239]==exp, "T4 обе строки позиции пересчитаны по своим единицам")
check(p.Неоднозначно==0, "T4 неоднозначно = 0")
# --- T5: все СС — полный пересчёт по сохранённым k: какие количества меняются
q=c.NewObject("Запрос"); q.Text='ВЫБРАТЬ С.Ссылка ИЗ Справочник.СтруктураСебестоимости КАК С ГДЕ НЕ С.ПометкаУдаления И С.КоэффициентыПересчета.НомерСтроки = 1'
t=q.Execute().Выгрузить()
print("T5 все СС с коэффициентами:", t.Количество())
for i in range(t.Количество()):
    o=t.Получить(i).Ссылка.ПолучитьОбъект(); before=snap(o)
    p=M.ПересчитатьКоличество(o); after=snap(o)
    ch=[(b,a) for b,a in zip(before,after) if b[1]!=a[1]]
    print(f" {o.Наименование}: изменено кол-в {len(ch)}; неодн={p.Неоднозначно}; ош={p.Ошибка}")
    for b,a in ch: print("   ",b,"->",a)
    for j in range(p.Детали.Количество()):
        s=p.Детали.Получить(j)
        if 'неоднозначно' in s: print("   ·",s)
print("FAILS:",len(fails))
