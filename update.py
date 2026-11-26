#!/usr/bin/env python3
"""Fetch an atomic, private RESURS (resursstore.com) dashboard snapshot. No credentials in HTML."""
import calendar,csv,datetime as dt,fcntl,html,io,json,os,time,uuid
from pathlib import Path
from decimal import Decimal
import requests

ROOT=Path(__file__).resolve().parent
def metrika(cfg,a,b,dimensions=None,site=False):
    q=dict(ids=90348201,date1=str(a),date2=str(b),metrics='ym:s:visits,ym:s:ecommerceRevenue,ym:s:goal297400586users,ym:s:goal3090348201reaches',accuracy='full',currency='RUB',limit=100000)
    if not site:q['filters']="ym:s:lastAdvEngine=='ya_direct'"
    if dimensions:q['dimensions']=dimensions
    for attempt in range(4):
        r=requests.get('https://api-metrika.yandex.net/stat/v1/data',params=q,headers={'Authorization':'OAuth '+cfg['metrika']},timeout=90)
        if r.status_code==200:break
        if r.status_code in (401,403) or attempt==3:raise RuntimeError('Metrika HTTP '+str(r.status_code)+': '+r.text[:800])
        time.sleep(5*(attempt+1))
    v=r.json()
    if v.get('sampled') or v.get('total_rows',0)>len(v.get('data',[])):raise RuntimeError('Incomplete Metrika report')
    return v

def direct(cfg,a,b,segment='Date'):
    body={'params':{'SelectionCriteria':{'DateFrom':str(a),'DateTo':str(b)},'FieldNames':[segment,'Impressions','Clicks','Cost'],'ReportName':'Resurs_'+uuid.uuid4().hex,'ReportType':'ACCOUNT_PERFORMANCE_REPORT','DateRangeType':'CUSTOM_DATE','Format':'TSV','IncludeVAT':'YES','IncludeDiscount':'NO'}}
    headers={'Authorization':'Bearer '+cfg['direct'],'Accept-Language':'ru','processingMode':'auto','returnMoneyInMicros':'false','skipReportHeader':'true','skipReportSummary':'true'}
    if cfg.get('client_login'):headers['Client-Login']=cfg['client_login']
    for _ in range(90):
        r=requests.post('https://api.direct.yandex.com/json/v5/reports',json=body,headers=headers,timeout=90)
        if r.status_code==200:
            rows=list(csv.DictReader(io.StringIO(r.text),delimiter='\t'))
            for row in rows:
                for k in ['Impressions','Clicks','Cost']:row[k]=float(Decimal(row[k]))
            return rows
        if r.status_code not in (201,202):raise RuntimeError('Direct HTTP '+str(r.status_code)+': '+r.text[:800])
        time.sleep(min(30,max(1,int(r.headers.get('retryIn',5)))))
    raise RuntimeError('Direct report timed out')

def e(x):return html.escape(str(x),quote=True)
def n(x,money=False):return f'{x:,.0f}'.replace(',',' ')+(' ₽' if money else '')
def delta(a,b):return f'{(b/a-1)*100:+.1f}%'.replace('.',',') if a else '—'
def table(headers,rows,total=None):
    def row(vals,tag='td'):return '<tr>'+''.join(f'<{tag}>{e(v)}</{tag}>' for v in vals)+'</tr>'
    return '<div class="table-wrap"><table><thead>'+row(headers,'th')+'</thead><tbody>'+''.join(row(r) for r in rows)+'</tbody>'+('<tfoot>'+row(total)+'</tfoot>' if total else '')+'</table></div>'
def bars(title,items,money=False):
    mx=max([v for _,v in items]+[1])
    return '<section class="trend-card"><h3>'+e(title)+'</h3>'+''.join('<div class="daily-line"><span>'+e(k)+'</span><span class="trend-track"><i class="year-26" style="width:'+str(max(0,v/mx*100))+'%"></i></span><b>'+n(v,money)+'</b></div>' for k,v in items)+'</section>'
def totals(rows,a,b):return [sum(r[k] for r in rows if str(a)<=r['Date']<=str(b)) for k in ['Cost','Impressions','Clicks']]
def period(a,b):return a.strftime('%d.%m')+'–'+b.strftime('%d.%m.%Y')
MONTHS=['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь']
def pct(a,b):return f'{a/b*100:.2f}%'.replace('.',',') if b else '—'
def rub2(a,b):return f'{a/b:,.2f}'.replace(',',' ').replace('.',',')+' ₽' if b else '—'
def rub0(a,b):return n(a/b,True) if b else '—'
def week_cells(v):
    cost,imp,clk,carts,buys=v[0],v[1],v[2],v[5],v[6]
    return [n(imp),n(clk),pct(clk,imp),rub2(cost,clk),n(cost,True),n(carts),pct(carts,clk),rub0(cost,carts),n(buys),pct(buys,carts),rub0(cost,buys)]
def weekly(cfg,daily,start,end,total):
    # Месяц → недели пн–вс, обрезанные по границам месяца и периода отчёта (как в сводной Excel).
    months={};d=start
    while d<=end:
        last=dt.date(d.year,d.month,calendar.monthrange(d.year,d.month)[1])
        b=min(d+dt.timedelta(days=6-d.weekday()),last,end)
        months.setdefault((d.year,d.month),[]).append((d,b));d=b+dt.timedelta(days=1)
    stats=lambda a,b:totals(daily,a,b)+metrika(cfg,a,b)['totals']
    head=['Период','Показы','Клики','CTR','CPC','Расход с НДС','Перешли в корзину · люди','CR в корзину','Стоимость перехода в корзину','Покупки','CR покупки из корзины','CPA']
    toggles='';css='';body=''
    for (y,m),weeks in months.items():
        cid=f'wk{y}{m:02d}'
        toggles+='<input type="checkbox" id="'+cid+'" class="wk-toggle" hidden>'
        css+='#'+cid+':checked~.table-wrap .'+cid+'-w{display:none}#'+cid+':checked~.table-wrap .'+cid+'-m .tog::before{content:"⊞"}'
        body+='<tr class="wk-month '+cid+'-m"><td><label for="'+cid+'"><span class="tog"></span> '+MONTHS[m-1]+' '+str(y)+'</label></td>'+''.join('<td>'+e(c)+'</td>' for c in week_cells(stats(weeks[0][0],weeks[-1][1])))+'</tr>'
        for a,b in weeks:
            body+='<tr class="wk-week '+cid+'-w"><td>'+e(a.strftime('%d.%m.%Y')+' - '+b.strftime('%d.%m.%Y'))+'</td>'+''.join('<td>'+e(c)+'</td>' for c in week_cells(stats(a,b)))+'</tr>'
    foot='<tr><td>Общий итог</td>'+''.join('<td>'+e(c)+'</td>' for c in week_cells(total))+'</tr>'
    html_='<div class="wk">'+toggles+'<div class="table-wrap"><table class="wk-table"><thead><tr>'+''.join('<th>'+e(h)+'</th>' for h in head)+'</tr></thead><tbody>'+body+'</tbody><tfoot>'+foot+'</tfoot></table></div></div>'
    return html_,css
def render(cfg):
    today=dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).date(); end=today-dt.timedelta(days=1)
    start=dt.date(2026,11,20)  # запуск рекламы РЕСУРС
    daily=direct(cfg,min(start,end-dt.timedelta(days=13)),end)
    met=metrika(cfg,start,end,'ym:s:date')
    by_date={r['dimensions'][0]['name']:r['metrics'] for r in met['data']}
    wa=end-dt.timedelta(days=6); pa=wa-dt.timedelta(days=7); pe=wa-dt.timedelta(days=1)
    last=totals(daily,wa,end)+metrika(cfg,wa,end)['totals']; old=totals(daily,pa,pe)+metrika(cfg,pa,pe)['totals']
    metrics=[('Расход с НДС',True),('Показы',False),('Клики',False),('Визиты',False),('Доход',True),('Перешли в корзину · люди',False),('Покупки',False)]
    s='<h1>РЕСУРС · resursstore.com</h1><p class="sub">Обновлено '+today.strftime('%d.%m.%Y')+' · данные по '+end.strftime('%d.%m.%Y')+'</p><h2>Итого · контекст за 7 дней</h2><p class="sub">'+period(wa,end)+' · сравнение с '+period(pa,pe)+'</p><div class="kpi-grid">'
    for (label,money),a,b in zip(metrics,old,last):s+='<div class="kpi"><div class="kpi-label">'+label+'</div><div class="kpi-value">'+n(b,money)+'</div><div class="kpi-delta">'+delta(a,b)+'</div></div>'
    s+='</div>'
    if last[5]==0:s+='<p class="sub">Корзины: цель «Корзина» не зарегистрировала посетителей за эти 7 дней. Требуется проверить настройку цели.</p>'
    s+='<h2>Контекст · ежедневная динамика</h2><p class="sub">'+period(start,end)+'</p>'
    days=[]
    for i in range((end-start).days+1):
        d=start+dt.timedelta(days=i); dr=totals(daily,d,d); mr=by_date.get(str(d),[0,0,0,0]); days.append((d,dr,mr))
    s+='<div class="trend-grid">'+bars('Расход с НДС',[(d.strftime('%d.%m'),dr[0]) for d,dr,mr in days],True)+bars('Перешли в корзину · люди',[(d.strftime('%d.%m'),mr[2]) for d,dr,mr in days])+'</div>'
    total_m=met['totals']; total_d=totals(daily,start,end)
    s+=table(['Дата','Расход с НДС','Показы','Клики','Доход','Перешли в корзину · люди','Покупки'],[[d.strftime('%d.%m'),n(dr[0],True),n(dr[1]),n(dr[2]),n(mr[1],True),n(mr[2]),n(mr[3])] for d,dr,mr in days],['Итого',n(total_d[0],True),n(total_d[1]),n(total_d[2]),n(total_m[1],True),n(total_m[2]),n(total_m[3])])
    wk_html,wk_css=weekly(cfg,daily,start,end,total_d+total_m)
    s+='<h2>Контекст · по неделям</h2><p class="sub">'+period(start,end)+' · недели пн–вс внутри месяца · нажмите на месяц, чтобы свернуть или развернуть</p>'+wk_html
    for field,title in [('LocationOfPresenceName','Гео · местонахождение'),('Device','Устройства')]:
        data=direct(cfg,start,end,field);data.sort(key=lambda r:r['Cost'],reverse=True)
        names={'DESKTOP':'Компьютеры','MOBILE':'Телефоны','TABLET':'Планшеты','SMART_TV':'Телевизоры'}
        items=[(names.get(r[field],r[field]),r) for r in data[:10]]
        if len(data)>10:items.append(('Остальные', {k:sum(r[k] for r in data[10:]) for k in ['Cost','Impressions','Clicks']}))
        s+='<h2>'+title+'</h2><p class="sub">Накопительно · '+period(start,end)+'</p>'+bars('Расход с НДС',[(k,r['Cost']) for k,r in items],True)+table([title,'Расход с НДС','Показы','Клики'],[[k,n(r['Cost'],True),n(r['Impressions']),n(r['Clicks'])] for k,r in items],['Итого']+[n(sum(r[k] for r in data),k=='Cost') for k in ['Cost','Impressions','Clicks']])
    s+='<h2>Весь сайт · визиты и доход год к году</h2><p class="sub">Все источники · одинаковые даты обоих лет</p>'
    years=[]
    for year in [end.year-1,end.year]:
        cutoff=dt.date(year,end.month,min(end.day,calendar.monthrange(year,end.month)[1]))
        result=metrika(cfg,dt.date(year,1,1),cutoff,'ym:s:month',True)
        months={int(r['dimensions'][0]['name'][-2:]):r['metrics'] for r in result['data']}
        years.append((months,result['totals']))
    month_names=MONTHS
    rows=[]
    for m in range(1,end.month+1):
        x=years[0][0].get(m,[0]*4); y=years[1][0].get(m,[0]*4)
        rows.append([month_names[m-1]+(f' · 1–{end.day}' if m==end.month else ''),n(x[0]),n(y[0]),delta(x[0],y[0]),n(x[1],True),n(y[1],True),delta(x[1],y[1])])
    x=years[0][1];y=years[1][1]
    s+='<div class="trend-grid">'
    for title,metric,money in [('Визиты',0,False),('Доход',1,True)]:
        mx=max([v[metric] for months,_ in years for v in months.values()]+[1])
        s+='<section class="trend-card"><h3>'+title+'</h3><p class="sub">Серый — '+str(end.year-1)+' · синий — '+str(end.year)+'</p>'
        for m in range(1,end.month+1):
            s+='<div class="trend-month"><strong>'+month_names[m-1]+'</strong><div>'
            for i,(months,_) in enumerate(years):
                value=months.get(m,[0]*4)[metric]
                s+='<div class="trend-line"><span class="trend-track"><i class="year-'+('25' if i==0 else '26')+'" style="width:'+str(value/mx*100)+'%"></i></span><span>'+n(value,money)+'</span></div>'
            s+='</div></div>'
        s+='</section>'
    s+='</div>'
    s+=table(['Месяц','Визиты '+str(end.year-1),'Визиты '+str(end.year),'Изменение','Доход '+str(end.year-1),'Доход '+str(end.year),'Изменение'],rows,['Итого',n(x[0]),n(y[0]),delta(x[0],y[0]),n(x[1],True),n(y[1],True),delta(x[1],y[1])])
    s+='<details class="method"><summary>Источники и методика</summary><p>Расход с НДС, показы, клики, гео и устройства — API Директа. Визиты, доход электронной коммерции, посетители по цели «Корзина» (ID 297400586) и достижения цели «Ecommerce: покупка» (ID 3090348201) — API Метрики, счётчик 90348201, без семплирования. Контекст — источник «Директ» по последнему переходу. Люди в итогах считаются за весь период: сумма дневных значений может отличаться. Ноль по корзинам означает отсутствие данных по этой цели, а не подтверждённое отсутствие корзин. Данные CRM в этот отчёт автоматически не поступают.</p></details>'
    css=(ROOT/'style.txt').read_text().replace('</style>','')
    css+=' .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px}.kpi{background:white;padding:20px;border-radius:12px}.kpi-value{font-size:26px;font-weight:650;margin:10px 0}.kpi-label,.kpi-delta{font-size:12px;color:#6b6a65}td:first-child,th:first-child{text-align:left}tfoot{font-weight:700;background:#e9e7dd}.table-wrap{overflow:auto}body{max-width:1400px;margin:auto}@media(max-width:600px){body{padding:12px}.trend-grid{grid-template-columns:1fr}}'
    css+=' .wk-month td{font-weight:700;background:#e9edf3}.wk-month label{cursor:pointer;user-select:none;display:block}.tog::before{content:"⊟";display:inline-block;width:1.1em;color:#5f5e5a}.wk-week td:first-child{padding-left:34px}.wk-table th{white-space:normal;vertical-align:bottom;min-width:70px}'+wk_css
    return '<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow"><title>РЕСУРС · отчёт</title><style>'+css+'</style></head><body>'+s+'</body></html>'

def main():
    with (ROOT/'update.lock').open('w') as lock:
        fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
        cfg=json.loads((ROOT/'config.json').read_text()); result=render(cfg)
        assert len(result)>10000 and 'Креативы' not in result
        temp=ROOT/'report.tmp';temp.write_text(result);os.chmod(temp,0o600);os.replace(temp,ROOT/'report.html')
        (ROOT/'status.json').write_text(json.dumps({'updated_at':dt.datetime.now(dt.timezone.utc).isoformat(),'ok':True}))
        print('RESURS: report updated successfully')
if __name__=='__main__':main()
