"""TikTok analysis: builds analysis/tiktok/metrics.json (deck input). Run from anywhere."""
import os, sys, json
HERE=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0,HERE)
from tt_load import *
OUT=os.environ.get('OUT_DIR', os.path.join(HERE,'..','analysis','tiktok')); os.makedirs(OUT,exist_ok=True)
def f(x):
    if isinstance(x,(bool,np.bool_)): return bool(x)
    if isinstance(x,(np.floating,float)): return None if (np.isnan(x) or np.isinf(x)) else float(x)
    if isinstance(x,np.integer): return int(x)
    return x
def ag(g): return {k:f(v) for k,v in agg(g).items()}
M={'periods':{'BEFORE':'1 Apr – 30 Jun 2026','AFTER':'1 Jul – 30 Sep 2026'}}
M['total']={}
for p,act,al,tot in [('BEFORE',B,B_all,B_tot),('AFTER',A,A_all,A_tot)]:
    d=ag(act)
    d.update(unique_reach=f(tot.reach), freq_unique=f(tot.impr/tot.reach), cost_1k_unique=f(tot.spend/tot.reach*1000),
             rows=int(len(al)), zero_spend_rows=int((~al.active).sum()), campaigns=int(act.campaign.nunique()),
             adgroups=int(act.groupby(['campaign','adgroup']).ngroups), total_row_spend=f(tot.spend), total_row_impr=f(tot.impr),
             total_row_clicks=f(tot.clicks), total_row_v2s=f(tot.v2s),
             zero_click_ads=int(((act.clicks==0)&(act.objective!='Community interaction')).sum()),
             zero_click_spend=f(act[(act.clicks==0)&(act.objective!='Community interaction')].spend.sum()))
    M['total'][p]=d
M['objective']={}
for p,act in [('BEFORE',B),('AFTER',A)]:
    for o,g in act.groupby('objective'):
        d=ag(g); d['results']=f(g.results.sum()); d['cost_per_result']=f(g.spend.sum()/g.results.sum()) if g.results.sum() else None
        d['share']=f(g.spend.sum()/act.spend.sum()*100); M['objective'].setdefault(o,{})[p]=d
M['theme_reach']={}
for p,act in [('BEFORE',B),('AFTER',A)]:
    for t,g in act[act.objective=='Reach'].groupby('theme'): M['theme_reach'].setdefault(t,{})[p]=ag(g)
def rows(act,keys):
    out=[]
    for k,g in act.groupby(keys):
        d=ag(g); 
        if isinstance(keys,list): d.update(dict(zip(keys,k)))
        else: d[keys]=k
        d['objective']='/'.join(sorted(set(g.objective))); d['results']=f(g.results.sum()); d['share']=f(g.spend.sum()/act.spend.sum()*100)
        d['adgroups']=int(g.adgroup.nunique()); out.append(d)
    return sorted(out,key=lambda x:-x['spend'])
M['campaigns']={p:rows(a,'campaign') for p,a in [('BEFORE',B),('AFTER',A)]}
M['adgroups']={p:rows(a,['campaign','adgroup']) for p,a in [('BEFORE',B),('AFTER',A)]}
def ads(act):
    d=act.copy(); d['ctr_pct']=d.clicks/d.impr*100; d['v2s_rate']=d.v2s/d.impr*100; d['cost_1k_reach']=d.spend/d.reach*1000
    d['cost_1k_v2s']=d.spend/d.v2s*1000; d['cpc_calc']=np.where(d.clicks>0,d.spend/d.clicks.replace(0,np.nan),np.nan)
    cols=['campaign','adgroup','ad','objective','theme','spend','reach','impr','freq','cpm','clicks','ctr_pct','cpc_calc','v2s','v2s_rate','cost_1k_v2s','cost_1k_reach','results']
    return [{k:f(v) for k,v in r.items()} for r in d[cols].to_dict('records')]
M['ads']={p:ads(a) for p,a in [('BEFORE',B),('AFTER',A)]}
M['zero_spend']={p:al[~al.active][['campaign','adgroup','ad']].to_dict('records') for p,al in [('BEFORE',B_all),('AFTER',A_all)]}
json.dump(M,open(os.path.join(OUT,'metrics.json'),'w'),indent=1,ensure_ascii=False,default=f)
T=M['total']
for k in ['spend','impr','unique_reach','freq_unique','cost_1k_unique','cpm','clicks','ctr','cpc','v2s','v2s_rate','cost_1k_v2s','reach_sum','cost_1k_reach']:
    b,a=T['BEFORE'][k],T['AFTER'][k]; print(f"{k:16s} {b:16,.4f} {a:16,.4f} {(a-b)/b*100:+7.1f}%")
for o,d in M['objective'].items(): print(o,{p:(v['ads'],round(v['spend'],2),round(v['share'],1),round(v['cpm'],3),round(v['cost_1k_reach'],3),round(v['ctr'],3),round(v['v2s_rate'],1)) for p,v in d.items()})
for t,d in M['theme_reach'].items(): print(t,{p:(v['ads'],round(v['spend']),round(v['cost_1k_reach'],3),round(v['cpm'],3),round(v['ctr'],3),round(v['v2s_rate'],1),round(v['freq_adlevel'],2)) for p,v in d.items()})
