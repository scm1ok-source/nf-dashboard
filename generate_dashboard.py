"""
НФ Dashboard Generator — Multi-Month
"""
import sys, os, math, warnings
from datetime import datetime
warnings.filterwarnings('ignore')

try:
    import pandas as pd
except ImportError:
    print("ERROR: pip install pandas openpyxl"); sys.exit(1)

XLSX = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Нф_2.xlsx')
if not os.path.exists(XLSX): print(f"ERROR: not found: {XLSX}"); sys.exit(1)
print(f"Reading: {XLSX}")
xl = pd.read_excel(XLSX, sheet_name=None)

PASSWORD = 'nf2026'  # ← change this to your desired password

MONTH_ORDER = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь']
MONTHS_SHORT= ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек']
CLEAN_LONG  = ['ОТК','Б2','Б3','Б4','Крашение','СШМ','Лабдип','Мембрана','Ворсование','Тумблирование','Вязание','Ленты','Склад ГП','Склад сырья','Лаборатория','Печать(ЦП+ВП)','Техническое обслуживание чистоты']
CLEAN_SHORT = ['ОТК','Б2','Б3','Б4','Кр.','СШМ','Лабдип','Мемб.','Ворс.','Тумб.','Вязан.','Ленты','Скл.ГП','Скл.С','Лаб.','Печать','ТО']

def sf(v, d=None):
    try: f=float(v); return None if math.isnan(f) else round(f,4)
    except: return d

def fmt(v, dec=1, suf='', d='—'):
    if v is None: return d
    try: f=float(v); return f"{int(round(f))}{suf}" if dec==0 else f"{f:.{dec}f}{suf}"
    except: return d

def fmt_time(s):
    s=str(s)
    for a,b in [(' час ',' ч '),(' часа ',' ч '),(' часов ',' ч '),(' минут','м'),(' минуты','м'),(' минута','м')]: s=s.replace(a,b)
    return s.strip()

def mclean(s): return str(s).replace(' 2026','').replace(' 2025','').strip().capitalize()

# ─────────────────────────────────────────────
data = {}
def md(mn):
    if mn not in data:
        data[mn]={'month':mn,'nir':None,'tr':None,'kar_n':None,'kar_t':'—','t48':None,
                  'reestr':None,'syr':None,'prom':None,'gott':None,'tmc_dev_d':None,'tmc_dev_pct':None,
                  'clean':None,'ship_vol':0,'ship_val':0,'rec_plan':0,'rec_fact':0,
                  'issues':[],'tmc_rows':[],'qual_weeks':[],'week_tasks':[]}
    return data[mn]

# НИР/ТР
nir_tr=xl['НИР_ТР'].copy(); nir_tr.columns=['Месяц','НИР','ТР']
for _,r in nir_tr.iterrows():
    mn=str(r['Месяц'])
    if mn in MONTH_ORDER: d=md(mn); d['nir']=sf(r['НИР'],0); d['tr']=sf(r['ТР'],0)

# KPI summary
kpi=xl['ДЛЯ КАЮ']
cur_month=str(kpi.iloc[2,0]); d=md(cur_month)
d['kar_n']=sf(kpi.iloc[2,4]); d['kar_t']=fmt_time(str(kpi.iloc[2,5]))
d['t48']=sf(kpi.iloc[2,6],0); d['reestr']=sf(kpi.iloc[2,8])
d['syr']=sf(kpi.iloc[2,10]); d['prom']=sf(kpi.iloc[2,11]); d['gott']=sf(kpi.iloc[2,12])
d['tmc_dev_d']=sf(kpi.iloc[2,14]); d['tmc_dev_pct']=sf(kpi.iloc[2,15])
ytd={'kar_n':sf(kpi.iloc[3,4]),'kar_t':fmt_time(str(kpi.iloc[3,5])),
     'reestr':sf(kpi.iloc[3,8]),'tmc_dev_d':sf(kpi.iloc[3,14]),'tmc_dev_pct':sf(kpi.iloc[3,15])}

# Cleanliness
clean_df=xl['Чистота_Порядок']
for _,r in clean_df.iterrows():
    mn=str(r.iloc[0])
    if mn in MONTH_ORDER:
        vals=[sf(r[dep]) for dep in CLEAN_LONG]
        if any(v is not None for v in vals): md(mn)['clean']=vals

# Сортность monthly + weekly
sort_df=xl['Сортность']
for _,r in sort_df.iterrows():
    mn=str(r['Месяц']) if pd.notna(r['Месяц']) else None
    if mn and mn in MONTH_ORDER:
        d=md(mn)
        if d['syr'] is None: d['syr']=sf(r['Сырье Месяц,%'])
        if d['prom'] is None: d['prom']=sf(r['Промежуточный контроль Месяц,%'])
        if d['gott'] is None: d['gott']=sf(r['Готовая ткань Месяц,%'])
    wn=sf(r['Недели']); sw=sf(r['Сырье Неделя,%']); pw=sf(r['Промежуточный контроль Неделя,%']); gw=sf(r['Готовая ткань Неделя,%'])
    if wn and (sw or pw or gw):
        try: mn2=MONTH_ORDER[pd.to_datetime(r.get('Дата')).month-1]
        except: mn2=cur_month
        md(mn2)['qual_weeks'].append({'week':f"Нед {int(wn)}",'syr':sw,'prom':pw,'gott':gw})

# Shipments
for _,r in xl['Отгрузки_Тех'].iterrows():
    mn=str(r['Месяц']).strip()
    if mn in MONTH_ORDER: md(mn)['ship_vol']=sf(r['shipment_volume_m'],0); md(mn)['ship_val']=sf(r['shipment_value_k_rub'],0)

# Receipts
for _,r in xl['Поступления_Тех'].iterrows():
    mn=str(r['Месяц']).strip()
    if mn in MONTH_ORDER: md(mn)['rec_plan']=sf(r['plan_k_rub'],0); md(mn)['rec_fact']=sf(r['fact_k_rub'],0)

# Quarantine
for _,r in xl['Карантин'].iterrows():
    try:
        mn=mclean(r['МесяцГод'])
        if mn not in MONTH_ORDER: continue
        hrs=sf(r['Часы до решения Число'],0)
        md(mn)['issues'].append({'date':pd.to_datetime(r['Дата и время создания']).strftime('%d.%m.%Y'),
            'task':str(int(r['№ задачи Bitrix24'])),'reason':str(r['Причина(кратко)']),
            'hrs':hrs,'time':fmt_time(str(r['Часы до Решения'])),'month':mn})
    except: pass

# Fill kar stats from raw issues for months missing KPI
for mn,d in data.items():
    if d['issues'] and d['kar_n'] is None:
        d['kar_n']=len(d['issues']); avg=sum(i['hrs'] for i in d['issues'])/len(d['issues'])
        h=int(avg); mins=int((avg-h)*60); d['kar_t']=f"{h}ч {mins}м"
        d['t48']=sum(1 for i in d['issues'] if (i['hrs'] or 0)>48)

# TMC
for _,r in xl['ВЭД_ТМЦ'].iterrows():
    raw=r.get('Месяц',None)
    if pd.isna(raw): continue
    mn=mclean(str(raw))
    if mn not in MONTH_ORDER: continue
    fact=r.get('Дата прихода (факт)'); plan_d=r.get('Дата прихода (план)')
    try: ps=pd.to_datetime(plan_d).strftime('%d.%m') if pd.notna(plan_d) else '—'
    except: ps='—'
    try: fs=pd.to_datetime(fact).strftime('%d.%m') if pd.notna(fact) else '—'
    except: fs='—'
    if fs=='—' and ps=='—': continue
    dev=sf(r.get('Отклонение (дни)'))
    md(mn)['tmc_rows'].append({'num':str(r.iloc[0]) if pd.notna(r.iloc[0]) else '—',
        'fabric':(str(r.get('Ткань',''))[:45] if pd.notna(r.get('Ткань',None)) else '—'),
        'qty':str(r.get('Кол-во факт','')) if pd.notna(r.get('Кол-во факт',None)) else '—',
        'unit':str(r.get('Unnamed: 7','')) if pd.notna(r.get('Unnamed: 7',None)) else '',
        'plan':ps,'fact':fs,'dev':int(dev) if dev is not None else None})

# Weekly tasks → assign to months
all_wt=[]
try:
    v=xl['Реестр задач'].iloc[1,1]
    if pd.notna(v): all_wt.append({'week':'1 неделя','pct':float(v)})
except: pass
df_rz=xl['Реестр задач']; i=0
while i<len(df_rz):
    lbl=str(df_rz.iloc[i,0])
    if 'неделя' in lbl.lower():
        for j in range(i+1,min(i+4,len(df_rz))):
            if str(df_rz.iloc[j,0])=='Задачи' and pd.notna(df_rz.iloc[j,1]):
                try: all_wt.append({'week':lbl,'pct':float(df_rz.iloc[j,1])}); break
                except: break
    i+=1
for idx,wt in enumerate(all_wt): md(MONTH_ORDER[min(idx//4,11)])['week_tasks'].append(wt)

def has_data(d):
    return any([d['kar_n'],d['syr'],d['gott'],d['clean'],(d['ship_vol'] or 0)>0,(d['rec_fact'] or 0)>0,d['issues'],d['tmc_rows']])

months_with_data=[mn for mn in MONTH_ORDER if mn in data and has_data(data[mn])]
print(f"✓ Months with data: {months_with_data}")

# ─── SVG helpers ────────────────────────────────────────────────────────────
def svg_bar(vals,labels,color_fn,height=180,w=340):
    vals=[v if v is not None else 0 for v in vals]
    has_neg=any(v<0 for v in vals); abs_max=max(abs(v) for v in vals) or 1
    hi=max(vals) if not has_neg else max(vals); lo=min(vals) if has_neg else 0
    rng=(hi-lo) or 1
    pad_l,pad_r,pad_t,pad_b=50,10,10,34; cw=w-pad_l-pad_r; ch=height-pad_t-pad_b
    bw=max(4,cw/max(len(vals),1)*0.55); sp=cw/max(len(vals),1)
    zero_y=pad_t+ch-((0-lo)/rng)*ch
    grid=''
    for step in range(5):
        tv=lo+(hi-lo)*step/4; y=pad_t+ch-((tv-lo)/rng)*ch
        lv=f"{tv/1000:.0f}K" if abs_max>5000 else (f"{tv:.0f}" if tv==int(tv) else f"{tv:.1f}")
        grid+=f'<line x1="{pad_l}" y1="{y:.1f}" x2="{w-pad_r}" y2="{y:.1f}" stroke="#252d3d" stroke-width="1"/>'
        grid+=f'<text x="{pad_l-4}" y="{y+4:.1f}" text-anchor="end" fill="#6b7a99" font-size="9">{lv}</text>'
    if has_neg: grid+=f'<line x1="{pad_l}" y1="{zero_y:.1f}" x2="{w-pad_r}" y2="{zero_y:.1f}" stroke="#6b7a99" stroke-width="1" stroke-dasharray="3,3"/>'
    bars=''
    for i,(v,lbl) in enumerate(zip(vals,labels)):
        x=pad_l+i*sp+sp/2-bw/2
        if has_neg:
            bh=max(2,abs(v/rng*ch)); by=zero_y if v<0 else pad_t+ch-((v-lo)/rng)*ch
        else:
            bh=max(2,(v/rng)*ch); by=pad_t+ch-bh
        c=color_fn(v,i)
        bars+=f'<rect x="{x:.1f}" y="{by:.1f}" width="{bw:.1f}" height="{bh:.1f}" fill="{c}" rx="3"/>'
        if v!=0:
            lv=f"{v/1000:.0f}K" if abs_max>5000 else (f"{v:.0f}" if v==int(v) else f"{v:.1f}")
            ty=by-3 if v>=0 else by+bh+10
            bars+=f'<text x="{x+bw/2:.1f}" y="{ty:.1f}" text-anchor="middle" fill="#e8edf5" font-size="8" font-weight="700">{lv}</text>'
        bars+=f'<text x="{x+bw/2:.1f}" y="{height-2}" text-anchor="middle" fill="#6b7a99" font-size="9">{lbl}</text>'
    return f'<svg viewBox="0 0 {w} {height}" width="100%" style="display:block;overflow:visible" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{height}" fill="none"/>{grid}{bars}</svg>'

def svg_line(datasets,x_labels,height=180,w=340,y_min=None,y_max=None):
    pad_l,pad_r,pad_t=42,10,10; ch=height-pad_t-34-16; cw=w-pad_l-pad_r
    all_v=[v for ds in datasets for v in ds['data'] if v is not None]
    if not all_v: return f'<svg viewBox="0 0 {w} {height}" width="100%"><text x="{w//2}" y="{height//2}" text-anchor="middle" fill="#6b7a99" font-size="11">Нет данных</text></svg>'
    lo=y_min if y_min is not None else min(all_v)-1; hi=y_max if y_max is not None else max(all_v)+1; rng=(hi-lo) or 1
    n=max(len(x_labels),1); xs=[pad_l+i*cw/max(n-1,1) for i in range(n)]
    grid=''.join(f'<line x1="{pad_l}" y1="{pad_t+ch-((lo+(hi-lo)*s/4)-lo)/rng*ch:.1f}" x2="{w-pad_r}" y2="{pad_t+ch-((lo+(hi-lo)*s/4)-lo)/rng*ch:.1f}" stroke="#252d3d" stroke-width="1"/><text x="{pad_l-4}" y="{pad_t+ch-((lo+(hi-lo)*s/4)-lo)/rng*ch+4:.1f}" text-anchor="end" fill="#6b7a99" font-size="9">{lo+(hi-lo)*s/4:.0f}</text>' for s in range(5))
    xlbls=''.join(f'<text x="{xs[i]:.1f}" y="{pad_t+ch+12}" text-anchor="middle" fill="#6b7a99" font-size="9">{lbl}</text>' for i,lbl in enumerate(x_labels))
    lines=''
    for ds in datasets:
        pts=[(xs[i],pad_t+ch-((v-lo)/rng)*ch) for i,v in enumerate(ds['data']) if v is not None]
        if len(pts)>1:
            path=' '.join(f'{"M" if j==0 else "L"}{x:.1f},{y:.1f}' for j,(x,y) in enumerate(pts))
            lines+=f'<path d="{path}" fill="none" stroke="{ds["color"]}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>'
        for x,y in pts: lines+=f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="{ds["color"]}" stroke="#0b0e14" stroke-width="1.5"/>'
    leg=''; lx=pad_l
    for ds in datasets:
        leg+=f'<rect x="{lx}" y="{height-14}" width="10" height="4" fill="{ds["color"]}" rx="2"/><text x="{lx+13}" y="{height-10}" fill="#6b7a99" font-size="9">{ds["label"]}</text>'; lx+=max(60,len(ds['label'])*6+20)
    return f'<svg viewBox="0 0 {w} {height}" width="100%" style="display:block;overflow:visible" xmlns="http://www.w3.org/2000/svg">{grid}{xlbls}{lines}{leg}</svg>'

def svg_donut(slices,height=160,w=300):
    total=sum(s[1] for s in slices)
    if not total: return f'<svg viewBox="0 0 {w} {height}"><text x="{w//2}" y="{height//2}" text-anchor="middle" fill="#6b7a99">Нет данных</text></svg>'
    cx,cy,R,r=80,80,70,40; angle=-90; paths=''
    for label,val,color in slices:
        sweep=360*val/total; a1,a2=math.radians(angle),math.radians(angle+sweep); lf=1 if sweep>180 else 0
        paths+=f'<path d="M{cx+R*math.cos(a1):.1f},{cy+R*math.sin(a1):.1f} A{R},{R} 0 {lf},1 {cx+R*math.cos(a2):.1f},{cy+R*math.sin(a2):.1f} L{cx+r*math.cos(a2):.1f},{cy+r*math.sin(a2):.1f} A{r},{r} 0 {lf},0 {cx+r*math.cos(a1):.1f},{cy+r*math.sin(a1):.1f} Z" fill="{color}"/>'; angle+=sweep
    leg=''.join(f'<rect x="165" y="{15+i*22}" width="10" height="10" fill="{color}" rx="2"/><text x="179" y="{24+i*22}" fill="#e8edf5" font-size="10">{label} ({val})</text>' for i,(label,val,color) in enumerate(slices))
    return f'<svg viewBox="0 0 {w} {height}" width="100%" style="display:block" xmlns="http://www.w3.org/2000/svg">{paths}{leg}</svg>'

# ─── charts per month ────────────────────────────────────────────────────────
DC=['#ff4757','#ffd23f','#ff6b35','#00e5ff','#7eff6a','#a55eea','#ff8800']
def defect_counts(issues):
    c={}
    for i in issues:
        r=i['reason'].lower()
        if 'расслаи' in r: k='Расслаивается'
        elif 'не соот' in r or 'несоот' in r: k='Не соотв. НД'
        elif 'помар' in r or 'подтёк' in r: k='Помарки/подтёки'
        elif 'адгез' in r: k='Адгезия'
        elif 'неравн' in r: k='Неравн. отделка'
        elif 'скрутк' in r: k='Скрутка тип С'
        else: k=i['reason'][:20]
        c[k]=c.get(k,0)+1
    return c

ship_vols_all=[(data[mn].get('ship_vol') or 0)/1000 if mn in data else 0 for mn in MONTH_ORDER]
ship_vals_all=[(data[mn].get('ship_val') or 0) if mn in data else 0 for mn in MONTH_ORDER]
chart_ship_vol=svg_bar(ship_vols_all,MONTHS_SHORT,lambda v,i:'rgba(0,229,255,0.75)' if ship_vols_all[i]==max(ship_vols_all) else ('rgba(0,229,255,0.45)' if v>0 else 'rgba(37,45,61,0.4)'))
chart_ship_val=svg_bar(ship_vals_all,MONTHS_SHORT,lambda v,i:'rgba(255,107,53,0.8)' if ship_vals_all[i]==max(ship_vals_all) else ('rgba(255,107,53,0.55)' if v>0 else 'rgba(37,45,61,0.4)'))

mcharts={}
for mn in months_with_data:
    d=data[mn]; c={}
    qw=d['qual_weeks']
    c['qual_week']=svg_line([{'label':'Готовая ткань','color':'#7eff6a','data':[w['gott'] for w in qw]},{'label':'Сырьё','color':'#00e5ff','data':[w['syr'] for w in qw]},{'label':'Пром. контроль','color':'#ffd23f','data':[w['prom'] for w in qw]}],[w['week'] for w in qw],y_min=85,y_max=101) if qw else f'<div style="padding:20px;text-align:center;color:#6b7a99">Нет данных</div>'
    dc=defect_counts(d['issues'])
    c['defect']=svg_donut([(k,v,DC[i%len(DC)]) for i,(k,v) in enumerate(dc.items())]) if dc else f'<div style="padding:20px;text-align:center;color:#6b7a99">Нет задач за {mn}</div>'
    if d['clean']:
        cv=[v if v is not None else 0 for v in d['clean']]
        c['clean']=svg_bar(cv,[s[:5] for s in CLEAN_SHORT],lambda v,i:'rgba(255,71,87,0.7)' if v<4.3 else ('rgba(126,255,106,0.7)' if v>=4.8 else 'rgba(255,210,63,0.6)'))
    else: c['clean']=f'<div style="padding:20px;text-align:center;color:#6b7a99">Нет данных</div>'
    wt=d['week_tasks']
    if wt:
        pcts=[w['pct']*100 for w in wt]; lbls=[w['week'].replace(' неделя','')+'н' for w in wt]
        c['tasks']=svg_bar(pcts,lbls,lambda v,i:'rgba(126,255,106,0.7)' if v==100 else ('rgba(255,210,63,0.7)' if v>=95 else 'rgba(255,71,87,0.7)'))
    else: c['tasks']=f'<div style="padding:20px;text-align:center;color:#6b7a99">Нет данных</div>'
    tr=d['tmc_rows']
    if tr:
        devs=[r['dev'] if r['dev'] is not None else 0 for r in tr]; lbls2=[f"#{r['num']}" for r in tr]
        c['supply']=svg_bar(devs,lbls2,lambda v,i:'rgba(255,71,87,0.7)' if v>0 else ('rgba(126,255,106,0.7)' if v==0 else 'rgba(255,210,63,0.6)'))
    else: c['supply']=f'<div style="padding:20px;text-align:center;color:#6b7a99">Нет данных</div>'
    mcharts[mn]=c

print(f"✓ Charts done")

# ─── HTML section per month ──────────────────────────────────────────────────
def pct_bar(val,color):
    w=min(100,max(0,float(val or 0)))
    return f'<div class="qbg"><div class="qbar" style="width:{w:.1f}%;background:{color}"></div></div>'

def kc(label,value,sub,badge,bcls,color,stripe):
    return f'<div class="kpi {stripe}"><div class="kl">{label}</div><div class="kv" style="color:var({color})">{value}</div><div class="ks">{sub}</div><span class="b {bcls}">{badge}</span></div>'

def ks(label,ll,lv,ls,rl,rv,rs,color,stripe):
    return f'<div class="kpi {stripe}"><div class="kl">{label}</div><div style="display:flex;gap:10px;align-items:flex-start"><div><div style="font-size:10px;color:var(--mu);font-weight:700;text-transform:uppercase">{ll}</div><div class="kv" style="color:var({color});font-size:26px">{lv}</div><div class="ks">{ls}</div></div><div style="width:1px;background:var(--bd);align-self:stretch;margin:2px 0"></div><div><div style="font-size:10px;color:var(--mu);font-weight:700;text-transform:uppercase">{rl}</div><div class="kv" style="color:var({color});font-size:26px">{rv}</div><div class="ks">{rs}</div></div></div></div>'

def build_section(mn):
    d=data[mn]; c=mcharts[mn]
    syr=d.get('syr'); prom=d.get('prom'); gott=d.get('gott')
    kar_n=d.get('kar_n'); kar_t=d.get('kar_t','—'); t48=d.get('t48')
    reestr=d.get('reestr'); tmc_d=d.get('tmc_dev_d'); tmc_p=d.get('tmc_dev_pct')
    sv=d.get('ship_vol') or 0; vv=d.get('ship_val') or 0
    rp=d.get('rec_plan') or 0; rf=d.get('rec_fact') or 0
    rpct=rf/rp*100 if rp else 0; tmc_good=tmc_d is not None and tmc_d<=0
    tmc_d_s=fmt(tmc_d,1,'д'); tmc_p_s=f"{(tmc_p or 0)*100:.1f}% {mn}" if tmc_p else '—'

    # clean cells
    cc_html=''
    if d.get('clean'):
        for dep,val in zip(CLEAN_SHORT,d['clean']):
            if val is None: col='var(--mu)';t='—'
            elif val<4.3: col='var(--dan)';t=f'{val:.1f}'
            elif val>=4.8: col='var(--ac3)';t=f'{val:.1f}'
            else: col='var(--ac4)';t=f'{val:.1f}'
            cc_html+=f'<div class="cc"><div class="cd">{dep}</div><div class="cv" style="color:{col}">{t}</div><div class="cu">/5.0</div></div>'
    else: cc_html=f'<div style="color:var(--mu);font-size:12px;padding:8px">Нет данных за {mn}</div>'

    iss_html=''.join(f'<div class="iss"><div class="issh"><div><div class="issn">#{i["task"]} · {i["month"]}</div><div class="issr">{i["reason"]}</div></div><div class="isst" style="color:{"var(--dan)" if (i["hrs"] or 0)>5 else "var(--ac2)"}">{i["time"]}</div></div><div class="issd">{i["date"]}</div></div>' for i in d.get('issues',[]))
    if not iss_html: iss_html=f'<div style="color:var(--mu);font-size:12px;padding:8px">Нет задач за {mn}</div>'

    tmc_html=''
    for r in d.get('tmc_rows',[]):
        dev=r['dev']
        if dev is None: pill='<span class="pill pw">— нет данных</span>'
        elif dev>0: pill=f'<span class="pill pl">+{dev}д задержка</span>'
        elif dev==0: pill='<span class="pill po">✓ В срок</span>'
        else: pill=f'<span class="pill pw">{dev}д раньше</span>'
        tmc_html+=f'<div class="sup"><div class="sn">{r["num"]}</div><div class="si"><div class="sname">{r["fabric"]}</div><div class="sd"><span>{r["qty"]} {r["unit"]}</span><span>план: {r["plan"]}</span><span>факт: {r["fact"]}</span></div></div>{pill}</div>'
    if not tmc_html: tmc_html=f'<div style="color:var(--mu);font-size:12px;padding:8px">Нет данных за {mn}</div>'

    wt_html=''
    for wt in d.get('week_tasks',[]):
        pct=int(wt['pct']*100); wcls='wg' if pct==100 else ('ww' if pct>=95 else 'wb')
        wt_html+=f'<div class="wc"><div class="wn">{wt["week"].capitalize()}</div><div class="wv {wcls}">{pct}%</div></div>'
    if not wt_html: wt_html=f'<div style="color:var(--mu);font-size:12px;padding:8px">Нет данных за {mn}</div>'

    pvf=''
    if rf>0:
        bw2=min(100,rpct); bc='linear-gradient(90deg,var(--ac3),#00ff88)' if rpct>=100 else 'linear-gradient(90deg,var(--dan),#ff6b6b)'
        nc='var(--ac3)' if rpct>=100 else 'var(--dan)'
        pvf=f'<div class="pvf"><div class="pvfh"><span>{mn}</span><span style="color:var(--mu)">{rf:,.0f} / {rp:,.0f} тыс.₽</span></div><div class="pvfbg"><div class="pvff" style="width:{bw2:.1f}%;background:{bc}"></div></div><div class="pvfn" style="color:{nc}">{rpct:.1f}%</div></div>'
    else: pvf=f'<div style="color:var(--mu);font-size:12px;padding:4px">Нет данных за {mn}</div>'

    tmc_devs=[r['dev'] for r in d.get('tmc_rows',[]) if r['dev'] is not None]
    avg_dev=sum(tmc_devs)/len(tmc_devs) if tmc_devs else 0
    on_time=sum(1 for v in tmc_devs if v==0); max_late=max((v for v in tmc_devs if v>0),default=0)
    pb='bg' if (prom or 0)>=97 else 'bw'; pt='✓ Норма' if (prom or 0)>=97 else '⚠ Внимание'

    return f'''<div class="month-view" id="mv-{mn}" style="display:none">
<section class="sec" id="{mn}-overview"><div class="stitle">СВОДКА <span>/ {mn.upper()}</span></div>
<div class="kg">
  {ks('Карантин',mn,fmt(kar_n,0),f'Ср. {kar_t}','YTD',fmt(ytd["kar_n"],0),f'Ср. {ytd["kar_t"]}','--ac','c1')}
  {kc('НИР / ТР',f'{fmt(d.get("nir"),0)} / {fmt(d.get("tr"),0)}',mn,'✓ Норма','bg','--ac3','c3')}
  {kc('Т48+',fmt(t48,0),mn,'✓ Норма' if fmt(t48,0)=='0' else '⚠ Есть','bg' if fmt(t48,0)=='0' else 'bb','--ac3','c3')}
  {kc('Реестр среднее',fmt(reestr,3),mn,'↑ Выше 1' if (reestr or 0)>=1 else '↓','bg' if (reestr or 0)>=1 else 'bb','--ac4','c4')}
  {kc('ТМЦ откл.',tmc_d_s,tmc_p_s,'↑ Раньше' if tmc_good else '↓ Задержка','bg' if tmc_good else 'bb','--ac3' if tmc_good else '--dan','c2')}
</div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac3)"></span>Сортность — {mn}</div>
  <div class="qrow"><div class="ql">Сырьё</div>{pct_bar(syr,"linear-gradient(90deg,#7eff6a,#2ecc71)")}<div class="qp" style="color:var(--ac3)">{fmt(syr,1,'%')}</div></div>
  <div class="qrow"><div class="ql">Пром. контроль</div>{pct_bar(prom,"linear-gradient(90deg,#ffd23f,#f39c12)")}<div class="qp" style="color:var(--ac4)">{fmt(prom,1,'%')}</div></div>
  <div class="qrow"><div class="ql">Готовая ткань</div>{pct_bar(gott,"linear-gradient(90deg,#7eff6a,#2ecc71)")}<div class="qp" style="color:var(--ac3)">{fmt(gott,1,'%')}</div></div>
</div>
<div class="two">
  <div class="card"><div class="ct"><span class="d" style="background:var(--ac)"></span>Отгрузки {mn}</div><div style="font-family:'Bebas Neue';font-size:24px;color:var(--ac)">{vv:,.0f} тыс.₽</div><div style="font-size:11px;color:var(--mu);margin-top:4px">{sv/1000:.1f}K м.п.</div></div>
  <div class="card"><div class="ct"><span class="d" style="background:var(--ac3)"></span>Поступления {mn}</div><div style="font-family:'Bebas Neue';font-size:24px;color:var(--ac3)">{rf:,.0f} тыс.₽</div><div style="font-size:11px;color:{"var(--ok)" if rpct>=100 else "var(--dan)"};margin-top:4px">{"▲" if rpct>=100 else "▼"} {rpct:.1f}% от плана</div></div>
</div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac4)"></span>Чистота — {mn}</div><div class="cg">{cc_html}</div></div>
</section><div class="div"></div>
<section class="sec" id="{mn}-finance"><div class="stitle">ОТГРУЗКИ <span>/ ПОСТУПЛЕНИЯ</span></div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac)"></span>Объём отгрузки (тыс. м.п.) — год</div><div class="chart-box">{chart_ship_vol}</div></div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac3)"></span>Поступления — {mn}</div>{pvf}</div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac4)"></span>Стоимость отгрузки (тыс.₽) — год</div><div class="chart-box">{chart_ship_val}</div></div>
</section><div class="div"></div>
<section class="sec" id="{mn}-quality"><div class="stitle">КАЧЕСТВО <span>/ СОРТНОСТЬ</span></div>
<div class="kg">
  <div class="kpi c3"><div class="kl">Сырьё {mn}</div><div class="kv" style="color:var(--ac3)">{fmt(syr,1,'%')}</div><span class="b bg">✓</span></div>
  <div class="kpi c4"><div class="kl">Пром. контроль</div><div class="kv" style="color:var(--ac4)">{fmt(prom,1,'%')}</div><span class="b {pb}">{pt}</span></div>
  <div class="kpi c3"><div class="kl">Готовая ткань</div><div class="kv" style="color:var(--ac3)">{fmt(gott,1,'%')}</div><span class="b bg">✓</span></div>
</div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac3)"></span>Сортность по неделям</div><div class="chart-box">{c["qual_week"]}</div></div>
</section><div class="div"></div>
<section class="sec" id="{mn}-quarantine"><div class="stitle">КАРАНТИН <span>/ ДЕФЕКТЫ</span></div>
<div class="kg">
  {ks('Задачи',mn,fmt(kar_n,0),'','YTD',fmt(ytd["kar_n"],0),'','--ac2','c2')}
  {ks('Ср. время',mn,kar_t,'','YTD',ytd["kar_t"],'','--ac','c1')}
  {kc('Т48+',fmt(t48,0),'просрочек','✓ Нет' if fmt(t48,0)=='0' else '⚠ Есть','bg' if fmt(t48,0)=='0' else 'bb','--ac3','c3')}
</div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac2)"></span>Причины дефектов</div><div class="chart-box">{c["defect"]}</div></div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac2)"></span>Задачи — {mn}</div>{iss_html}</div>
</section><div class="div"></div>
<section class="sec" id="{mn}-cleanliness"><div class="stitle">ЧИСТОТА <span>/ ПОРЯДОК</span></div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac4)"></span>Оценки — {mn} (из 5.0)</div><div class="chart-box">{c["clean"]}</div></div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac4)"></span>Детали — {mn}</div><div class="cg">{cc_html}</div></div>
</section><div class="div"></div>
<section class="sec" id="{mn}-tasks"><div class="stitle">РЕЕСТР <span>/ ЗАДАЧИ</span></div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac3)"></span>% выполнения по неделям</div><div class="chart-box">{c["tasks"]}</div></div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac)"></span>Недельные показатели</div><div class="wg2">{wt_html}</div></div>
</section><div class="div"></div>
<section class="sec" id="{mn}-supply"><div class="stitle">ТМЦ <span>/ ПОСТАВКИ</span></div>
<div class="kg">
  <div class="kpi c1"><div class="kl">Позиций</div><div class="kv" style="color:var(--ac)">{len(d.get("tmc_rows",[]))}</div></div>
  <div class="kpi c3"><div class="kl">В срок</div><div class="kv" style="color:var(--ac3)">{on_time}</div></div>
  <div class="kpi c2"><div class="kl">Ср. откл.</div><div class="kv" style="color:var(--{"ac3" if avg_dev<=0 else "dan"})">{avg_dev:.1f}д</div></div>
  <div class="kpi c4"><div class="kl">Макс. задержка</div><div class="kv" style="color:var(--{"dan" if max_late>0 else "ac3"})">{f"+{max_late}д" if max_late>0 else "—"}</div></div>
</div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac)"></span>Поставки — {mn}</div>{tmc_html}</div>
<div class="card"><div class="ct"><span class="d" style="background:var(--ac4)"></span>Отклонение (дни)</div><div class="chart-box">{c["supply"]}</div></div>
</section>
</div>'''

all_sec='\n'.join(build_section(mn) for mn in months_with_data)
sel_btns='\n    '.join(f'<button class="mbtn" data-month="{mn}">{mn[:3]}</button>' for mn in months_with_data)
default_mn=months_with_data[-1] if months_with_data else 'Март'
gen_at=datetime.now().strftime('%d.%m.%Y %H:%M')

CSS='''
:root{--bg:#0b0e14;--sf:#131720;--sf2:#1a2030;--bd:#252d3d;--ac:#00e5ff;--ac2:#ff6b35;--ac3:#7eff6a;--ac4:#ffd23f;--tx:#e8edf5;--mu:#6b7a99;--dan:#ff4757;--ok:#2ecc71;--wa:#f1c40f;--r:14px;--g:14px;}
*{box-sizing:border-box;margin:0;padding:0;}
html{scroll-behavior:smooth;}
body{font-family:'Manrope',sans-serif;background:var(--bg);color:var(--tx);padding-bottom:90px;}
.hdr{background:#0d1520;border-bottom:1px solid var(--bd);padding:12px 18px;display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;}
.hdr h1{font-family:'Bebas Neue';font-size:clamp(18px,5vw,24px);letter-spacing:2px;display:flex;align-items:center;gap:10px;}
.dot{width:9px;height:9px;border-radius:50%;background:var(--ac);box-shadow:0 0 10px var(--ac);animation:blink 2s infinite;flex-shrink:0;}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
.mselector{display:flex;gap:6px;flex-wrap:wrap;}
.mbtn{padding:7px 16px;border-radius:8px;border:1px solid var(--bd);background:var(--sf2);color:var(--mu);font-family:'Manrope',sans-serif;font-size:13px;font-weight:700;cursor:pointer;transition:all .15s;-webkit-tap-highlight-color:transparent;touch-action:manipulation;-webkit-appearance:none;}
.mbtn.active{background:rgba(0,229,255,.15);border-color:rgba(0,229,255,.5);color:var(--ac);}
.nav{position:sticky;top:0;z-index:100;background:var(--sf);border-bottom:1px solid var(--bd);display:flex;overflow-x:auto;gap:4px;padding:10px 14px;scrollbar-width:none;}
.nav::-webkit-scrollbar{display:none;}
.nav a{flex-shrink:0;padding:8px 14px;border-radius:8px;border:1px solid transparent;background:transparent;color:var(--mu);font-size:13px;font-weight:600;text-decoration:none;white-space:nowrap;-webkit-tap-highlight-color:transparent;}
.nav a:active{background:var(--sf2);color:var(--tx);}
.sec{padding:16px;scroll-margin-top:52px;}
.stitle{font-family:'Bebas Neue';font-size:clamp(20px,5vw,28px);letter-spacing:2px;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--bd);}
.stitle span{color:var(--ac);}
.div{height:8px;background:var(--sf2);border-top:1px solid var(--bd);border-bottom:1px solid var(--bd);}
.kg{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:var(--g);margin-bottom:var(--g);}
.kpi{background:var(--sf);border:1px solid var(--bd);border-radius:var(--r);padding:14px;position:relative;overflow:hidden;}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.c1::before{background:linear-gradient(90deg,var(--ac),transparent);}
.c2::before{background:linear-gradient(90deg,var(--ac2),transparent);}
.c3::before{background:linear-gradient(90deg,var(--ac3),transparent);}
.c4::before{background:linear-gradient(90deg,var(--ac4),transparent);}
.kl{font-size:11px;color:var(--mu);font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:7px;}
.kv{font-family:'Bebas Neue';font-size:34px;line-height:1;}
.ks{font-size:11px;color:var(--mu);margin-top:5px;}
.b{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;margin-top:5px;}
.bg{background:rgba(46,204,113,.15);color:var(--ok);}
.bw{background:rgba(241,196,15,.15);color:var(--wa);}
.bb{background:rgba(255,71,87,.15);color:var(--dan);}
.card{background:var(--sf);border:1px solid var(--bd);border-radius:var(--r);padding:16px;margin-bottom:var(--g);}
.ct{font-size:12px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px;display:flex;align-items:center;gap:8px;}
.d{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.chart-box{width:100%;overflow:hidden;}
.two{display:grid;grid-template-columns:1fr 1fr;gap:var(--g);margin-bottom:var(--g);}
@media(max-width:540px){.two{grid-template-columns:1fr;}}
.qrow{display:flex;align-items:center;gap:8px;margin-bottom:10px;}
.ql{font-size:12px;color:var(--mu);width:110px;flex-shrink:0;}
.qbg{flex:1;height:9px;background:var(--bd);border-radius:5px;overflow:hidden;}
.qbar{height:100%;border-radius:5px;}
.qp{font-size:12px;font-weight:700;width:50px;text-align:right;flex-shrink:0;}
.pvf{margin-bottom:12px;}
.pvfh{display:flex;justify-content:space-between;margin-bottom:4px;font-size:12px;}
.pvfbg{height:8px;background:var(--bd);border-radius:4px;position:relative;}
.pvff{height:100%;border-radius:4px;position:absolute;top:0;}
.pvfn{font-size:11px;font-weight:700;margin-top:3px;}
.cg{display:grid;grid-template-columns:repeat(auto-fill,minmax(90px,1fr));gap:8px;}
.cc{background:var(--sf2);border:1px solid var(--bd);border-radius:10px;padding:10px;text-align:center;}
.cd{font-size:11px;color:var(--mu);margin-bottom:4px;}
.cv{font-family:'Bebas Neue';font-size:26px;line-height:1;}
.cu{font-size:10px;color:var(--mu);}
.wg2{display:grid;grid-template-columns:repeat(auto-fill,minmax(90px,1fr));gap:8px;}
.wc{background:var(--sf2);border:1px solid var(--bd);border-radius:10px;padding:12px 8px;text-align:center;}
.wn{font-size:11px;color:var(--mu);margin-bottom:4px;font-weight:600;}
.wv{font-family:'Bebas Neue';font-size:28px;line-height:1;}
.wg{color:var(--ac3);}.ww{color:var(--ac4);}.wb{color:var(--dan);}
.iss{background:var(--sf2);border:1px solid var(--bd);border-left:3px solid var(--ac2);border-radius:10px;padding:11px 13px;margin-bottom:7px;}
.issh{display:flex;justify-content:space-between;gap:8px;margin-bottom:4px;}
.issn{font-size:11px;color:var(--mu);font-weight:600;}
.issr{font-size:13px;font-weight:700;}
.isst{font-family:'Bebas Neue';font-size:18px;white-space:nowrap;}
.issd{font-size:11px;color:var(--mu);margin-top:2px;}
.sup{display:flex;gap:10px;padding:11px;background:var(--sf2);border:1px solid var(--bd);border-radius:10px;margin-bottom:7px;align-items:flex-start;}
.sn{font-family:'Bebas Neue';font-size:20px;color:var(--mu);width:26px;flex-shrink:0;}
.si{flex:1;}
.sname{font-size:13px;font-weight:700;margin-bottom:3px;line-height:1.3;}
.sd{font-size:11px;color:var(--mu);display:flex;gap:8px;flex-wrap:wrap;}
.pill{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600;}
.pl{background:rgba(255,71,87,.15);color:var(--dan);}
.po{background:rgba(46,204,113,.15);color:var(--ok);}
.pw{background:rgba(255,210,63,.15);color:var(--ac4);}
.back{position:fixed;bottom:22px;right:18px;z-index:999;background:var(--ac);color:#000;font-weight:800;font-size:13px;padding:11px 18px;border-radius:50px;border:none;cursor:pointer;font-family:'Manrope',sans-serif;box-shadow:0 4px 18px rgba(0,229,255,.45);-webkit-tap-highlight-color:transparent;touch-action:manipulation;-webkit-appearance:none;}

#pw-screen{position:fixed;inset:0;background:var(--bg);z-index:9999;display:flex;align-items:center;justify-content:center;padding:24px;}
.pw-box{background:var(--sf);border:1px solid var(--bd);border-radius:20px;padding:36px 28px;width:100%;max-width:340px;text-align:center;}
.pw-logo{font-family:'Bebas Neue';font-size:36px;letter-spacing:3px;color:var(--tx);margin-bottom:4px;}
.pw-sub{font-size:12px;color:var(--mu);margin-bottom:28px;letter-spacing:.5px;}
.pw-input{width:100%;background:var(--sf2);border:1px solid var(--bd);border-radius:10px;padding:14px 16px;font-family:'Manrope',sans-serif;font-size:16px;color:var(--tx);text-align:center;letter-spacing:3px;outline:none;-webkit-appearance:none;}
.pw-input:focus{border-color:rgba(0,229,255,.5);}
.pw-btn{margin-top:14px;width:100%;background:var(--ac);color:#000;font-family:'Manrope',sans-serif;font-weight:800;font-size:15px;padding:14px;border-radius:10px;border:none;cursor:pointer;-webkit-appearance:none;touch-action:manipulation;-webkit-tap-highlight-color:transparent;}
.pw-btn:active{opacity:.85;}
.pw-err{margin-top:12px;font-size:12px;color:var(--dan);min-height:18px;font-weight:600;}
.pw-dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--ac);box-shadow:0 0 8px var(--ac);animation:blink 2s infinite;margin-right:8px;}

.footer{text-align:center;padding:18px;font-size:11px;color:var(--mu);border-top:1px solid var(--bd);}
'''

import base64 as _b64
pw_b64 = _b64.b64encode(PASSWORD.encode()).decode()
JS = f'''
// Password gate
(function() {{
  var _k = '{pw_b64}';
  window.checkPw = function() {{
    var val = document.getElementById('pw-input').value;
    var ok = false; try {{ ok = (val === atob(_k)); }} catch(e) {{}}
    if (ok) {{
      document.getElementById('pw-screen').style.display = 'none';
      document.getElementById('main-content').style.display = 'block';
      try {{ sessionStorage.setItem('nf_auth','1'); }} catch(e) {{}}
    }} else {{
      var err = document.getElementById('pw-err');
      err.textContent = '\u041d\u0435\u0432\u0435\u0440\u043d\u044b\u0439 \u043f\u0430\u0440\u043e\u043b\u044c';
      document.getElementById('pw-input').value = '';
      document.getElementById('pw-input').focus();
      setTimeout(function(){{ err.textContent=''; }}, 2000);
    }}
  }};
}})();
var cur = '{default_mn}';

function switchMonth(mn) {{
  var views = document.querySelectorAll('.month-view');
  for (var i = 0; i < views.length; i++) {{ views[i].style.display = 'none'; }}
  var target = document.getElementById('mv-' + mn);
  if (target) target.style.display = 'block';
  var btns = document.querySelectorAll('.mbtn');
  for (var i = 0; i < btns.length; i++) {{
    btns[i].classList.remove('active');
    if (btns[i].getAttribute('data-month') === mn) btns[i].classList.add('active');
  }}
  cur = mn;
  window.scrollTo({{top: 0, behavior: 'smooth'}});
}}

function scrollToSection(sec) {{
  var el = document.getElementById(cur + '-' + sec);
  if (el) el.scrollIntoView({{behavior: 'smooth', block: 'start'}});
}}


// Password gate
var _k = 'bmYyMDI2';
function checkPw() {{
  var val = document.getElementById('pw-input').value;
  var ok = false;
  try {{ ok = (val === atob(_k)); }} catch(e) {{}}
  if (ok) {{
    document.getElementById('pw-screen').style.display = 'none';
    document.getElementById('main-content').style.display = 'block';
    try {{ sessionStorage.setItem('nf_auth','1'); }} catch(e) {{}}
  }} else {{
    var err = document.getElementById('pw-err');
    err.textContent = 'Неверный пароль';
    document.getElementById('pw-input').value = '';
    document.getElementById('pw-input').focus();
    setTimeout(function(){{ err.textContent=''; }}, 2000);
  }}
}}

window.addEventListener('load', function() {{
  // Session: skip password if already authenticated
  try {{
    if (sessionStorage.getItem('nf_auth') === '1') {{
      document.getElementById('pw-screen').style.display = 'none';
      document.getElementById('main-content').style.display = 'block';
    }}
  }} catch(e) {{}}
  // Password button
  var pb = document.getElementById('pw-btn');
  if (pb) pb.addEventListener('click', checkPw);
  // Enter key
  var pi = document.getElementById('pw-input');
  if (pi) pi.addEventListener('keydown', function(e) {{
    if (e.key === 'Enter' || e.keyCode === 13) checkPw();
  }});

  var btns = document.querySelectorAll('.mbtn');
  for (var i = 0; i < btns.length; i++) {{
    (function(btn) {{
      btn.addEventListener('click', function() {{
        switchMonth(btn.getAttribute('data-month'));
      }});
    }})(btns[i]);
  }}
  var navLinks = document.querySelectorAll('a[data-sec]');
  for (var i = 0; i < navLinks.length; i++) {{
    (function(a) {{
      a.addEventListener('click', function(e) {{
        e.preventDefault();
        scrollToSection(a.getAttribute('data-sec'));
      }});
    }})(navLinks[i]);
  }}
  switchMonth('{default_mn}');
  // Password session + listeners
  try {{
    if (sessionStorage.getItem('nf_auth') === '1') {{
      document.getElementById('pw-screen').style.display = 'none';
      document.getElementById('main-content').style.display = 'block';
    }}
  }} catch(e) {{}}
  var pb = document.getElementById('pw-btn');
  if (pb) pb.addEventListener('click', checkPw);
  var pi = document.getElementById('pw-input');
  if (pi) pi.addEventListener('keydown', function(e) {{
    if (e.key === 'Enter' || e.keyCode === 13) checkPw();
  }});
}});
'''

HTML = f'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>НФ Дашборд 2026</title>
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Manrope:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<div id="pw-screen">
  <div class="pw-box">
    <div class="pw-logo"><span class="pw-dot"></span>НФ ДАШБОРД</div>
    <div class="pw-sub">ПРОИЗВОДСТВЕННЫЙ ДАШБОРД 2026</div>
    <input class="pw-input" id="pw-input" type="password" placeholder="••••••" maxlength="32" autocomplete="current-password">
    <button class="pw-btn" id="pw-btn">ВОЙТИ</button>
    <div class="pw-err" id="pw-err"></div>
  </div>
</div>
<div id="main-content" style="display:none">
<div class="hdr">
  <h1><div class="dot"></div>НФ ДАШБОРД</h1>
  <div class="mselector">{sel_btns}</div>
</div>
<nav class="nav" id="navtop">
  <a href="#" data-sec="overview">📊 Обзор</a>
  <a href="#" data-sec="finance">💰 Отгрузки</a>
  <a href="#" data-sec="quality">🎯 Качество</a>
  <a href="#" data-sec="quarantine">⚠️ Карантин</a>
  <a href="#" data-sec="cleanliness">🧹 Чистота</a>
  <a href="#" data-sec="tasks">✅ Задачи</a>
  <a href="#" data-sec="supply">📦 ТМЦ</a>
</nav>
{all_sec}
<div class="footer">НФ Производственный Дашборд 2026 · Сгенерировано: {gen_at}</div>
<button class="back" onclick="window.scrollTo({{top:0,behavior:'smooth'}})">&#8679; Меню</button>
<script>{JS}</script>
</div><!-- /main-content -->
</body>
</html>'''

out=os.path.join(os.path.dirname(os.path.abspath(__file__)),'index.html')
with open(out,'w',encoding='utf-8') as f: f.write(HTML)
print(f"✓ Saved: {out}  ({len(HTML)//1024}KB)  Months: {months_with_data}")
