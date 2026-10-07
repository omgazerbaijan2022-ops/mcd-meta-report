"""Final analysis: builds metrics.json (deck input) and analysis_tables.xlsx (audit trail)."""
import json, re, sys
import os
HERE=os.path.dirname(os.path.abspath(__file__))
exec(open(os.path.join(HERE,'load.py')).read())
OUT=os.environ.get('OUT_DIR', os.path.join(HERE,'..','analysis'))
OBJ={'Reach':'Reach','Instagram profile visits':'Traffic','Link clicks':'Traffic','Post engagements':'Engagement','Interactions':'Engagement','ThruPlay':'Video views','2-Second Continuous Video View':'Video views','':'App installs'}
CREATOR=r'collab|_inf_|influencer|blogger'
def theme(r):
    s=(r.campaign+' '+r.adset+' '+r.ad).lower()
    for n,p in [('FIFA / World Cup',r'fifa|world_cup'),('Mixology',r'mixology'),('Chicken',r'toyuq|chicken'),('McCafé',r'cafe'),('Grimace',r'grimace'),('Restaurant openings',r'opening')]:
        if re.search(p,s): return n
    return 'Other'
for df in (A,B):
    df['objective']=df.result_type.map(OBJ); df['theme']=df.apply(theme,axis=1)
    df['is_video']=df.v3s.fillna(0)>0
    df['creator']=(df.campaign+' '+df.adset+' '+df.ad).str.lower().str.contains(CREATOR)
def f(x):
    if isinstance(x,(np.floating,float)): return None if (np.isnan(x) or np.isinf(x)) else float(x)
    if isinstance(x,(np.integer,)): return int(x)
    return x
def ag(g): return {k:f(v) for k,v in agg(g).items()}
def pct(a,b): return (b-a)/a*100
M={}
M['total']={'BEFORE':ag(B),'AFTER':ag(A)}
M['total']['BEFORE']['campaigns']=B.campaign.nunique(); M['total']['AFTER']['campaigns']=A.campaign.nunique()
M['total']['BEFORE']['adsets']=B.groupby(['campaign','adset']).ngroups; M['total']['AFTER']['adsets']=A.groupby(['campaign','adset']).ngroups
for p,df in [('BEFORE',B),('AFTER',A)]:
    M['total'][p]['v3s_share_of_eng']=f(df.v3s.sum()/df.post_eng.sum()*100)
    v=df[df.is_video]; M['total'][p]['video_spend_share']=f(v.spend.sum()/df.spend.sum()*100)
    M['total'][p]['video_ads']=int(len(v)); M['total'][p]['video_v3s_rate']=f(v.v3s.sum()/v.impr.sum()*100)
    M['total'][p]['video_cost_1k_v3s']=f(v.spend.sum()/v.v3s.sum()*1000)
    M['total'][p]['spend_share_freq2']=f(df[df.freq>=2].spend.sum()/df.spend.sum()*100)
    M['total'][p]['objectives']=sorted(df.objective.unique().tolist())
# objective
M['objective']={}
for p,df in [('BEFORE',B),('AFTER',A)]:
    for o,g in df.groupby('objective'):
        d=ag(g); d['results']=f(g.results.sum()); d['result_types']=sorted(set(g.result_type)); M['objective'].setdefault(o,{})[p]=d
# result types
M['result_type']={}
for p,df in [('BEFORE',B),('AFTER',A)]:
    for rt,g in df.groupby('result_type'):
        M['result_type'].setdefault(rt or '(blank – app install, no results reported)',{})[p]={'ads':int(len(g)),'results':f(g.results.sum()),'spend':f(g.spend.sum()),'cpr':f(g.spend.sum()/g.results.sum()) if g.results.sum()>0 else None,'post_eng':f(g.post_eng.sum()),'cost_per_post_eng':f(g.spend.sum()/g.post_eng.sum())}
# link clicks by objective
M['clicks_by_obj']={p:{o:f(g.link_clicks.sum()) for o,g in df.groupby('objective')} for p,df in [('BEFORE',B),('AFTER',A)]}
# themes reach-only
M['theme_reach']={}
for p,df in [('BEFORE',B),('AFTER',A)]:
    for t,g in df[df.objective=='Reach'].groupby('theme'):
        M['theme_reach'].setdefault(t,{})[p]=ag(g)
# campaigns
def camp(df):
    out=[]
    for c,g in df.groupby('campaign'):
        d=ag(g); d['campaign']=c; d['objective']='/'.join(sorted(set(g.objective))); d['results']=f(g.results.sum()); d['result_types']='/'.join(sorted(set(g.result_type)))
        d['share']=f(g.spend.sum()/df.spend.sum()*100); d['adsets']=g.adset.nunique(); out.append(d)
    return sorted(out,key=lambda x:-x['spend'])
M['campaigns']={'BEFORE':camp(B),'AFTER':camp(A)}
# adsets
def adsets(df):
    out=[]
    for (c,a),g in df.groupby(['campaign','adset']):
        d=ag(g); d.update(campaign=c,adset=a,objective='/'.join(sorted(set(g.objective)))); out.append(d)
    return sorted(out,key=lambda x:-x['spend'])
M['adsets']={'BEFORE':adsets(B),'AFTER':adsets(A)}
# ads
def ads(df):
    d=df.copy()
    d['link_ctr']=d.link_clicks.fillna(0)/d.impr*100; d['eng_rate']=d.post_eng/d.impr*100; d['v3_rate']=d.v3s.fillna(0)/d.impr*100
    d['c1kr']=d.spend/d.reach*1000; d['cpc_calc']=d.spend/d.link_clicks
    cols=['campaign','adset','ad','objective','result_type','results','spend','impr','reach','freq','cpm','link_clicks','link_ctr','cpc_calc','post_eng','eng_rate','v3s','v3_rate','c1kr','is_video','creator','theme']
    return [{k:f(v) if not isinstance(v,(bool,np.bool_)) else bool(v) for k,v in r.items()} for r in d[cols].to_dict('records')]
M['ads']={'BEFORE':ads(B),'AFTER':ads(A)}
# platform (BEFORE names only)
M['platform_before']={}
for suf in ['ig','fb']:
    g=B[B.ad.str.lower().str.endswith('_'+suf)&(B.objective=='Reach')]; M['platform_before'][suf]=ag(g)
# creator
M['creator_reach']={}
for p,df in [('BEFORE',B),('AFTER',A)]:
    r=df[df.objective=='Reach']
    M['creator_reach'][p]={'creator':ag(r[r.creator]),'other':ag(r[~r.creator]),'ads':r[r.creator].ad.tolist()}
# carry-over and sensitivity
co=A.campaign.str.contains("June'26")
M['carryover']={'ag':ag(A[co]),'ads':A[co].ad.tolist(),'after_ex':ag(A[~co])}
M['march_in_before']=ag(B[B.campaign.str.contains("March")])
r=A[A.objective=='Reach']
M['reach_after_ex_sept']=ag(r[~r.campaign.isin(['EntryLevel_Reach_Sep_26',"HM_Reach_Sep'26"])])
M['dupe_after']=A[A.duplicated(['campaign','adset','ad'],keep=False)][['campaign','adset','ad','spend']].to_dict('records')
M['missing']={p:{'link_clicks_blank':int(df.link_clicks.isna().sum()),'v3s_blank':int(df.v3s.isna().sum()),'results_blank':int(df.results.isna().sum())} for p,df in [('BEFORE',B),('AFTER',A)]}
json.dump(M,open(os.path.join(OUT,'metrics.json'),'w'),indent=1,ensure_ascii=False,default=f)

# ---- KPI comparison table
K=[('Amount spent (USD)','spend','scale',None),('Impressions','impr','scale',None),('Reach (sum of ad-level reach)','reach_sum','scale',None),
   ('Frequency (impr ÷ reach)','freq','eff','lower'),('CPM (USD)','cpm','eff','lower'),('Cost per 1,000 reached (USD)','cost_per_1k_reach','eff','lower'),
   ('Link clicks','link_clicks','scale',None),('Link CTR (%)','link_ctr','eff','higher'),('CTR (all) (%)','ctr_all','eff','higher'),('CPC – link (USD)','cpc','eff','lower'),
   ('Post engagements','post_eng','scale',None),('Engagement rate (% of impr)','eng_rate','eff','higher'),('Cost per post engagement (USD)','cpe','eff','lower'),
   ('3-second video plays','v3s','scale',None),('3s plays per impression (%)','v3s_rate','eff','higher'),('Cost per 1,000 3s plays (USD)','cost_per_1k_v3s','eff','lower')]
rows=[]
for lab,k,typ,good in K:
    b=M['total']['BEFORE'][k]; a=M['total']['AFTER'][k]; ch=pct(b,a)
    verdict='Scale' if typ=='scale' else ('Better' if (ch<0 if good=='lower' else ch>0) else 'Worse')
    rows.append(dict(KPI=lab,BEFORE=b,AFTER=a,Abs_change=a-b,Pct_change=ch,Type=typ,Verdict=verdict))
kpi=pd.DataFrame(rows)
with pd.ExcelWriter(os.path.join(OUT,'analysis_tables.xlsx')) as w:
    kpi.to_excel(w,sheet_name='KPI summary',index=False)
    pd.DataFrame([dict(objective=o,period=p,**{k:v for k,v in d.items() if k!='result_types'}) for o,dd in M['objective'].items() for p,d in dd.items()]).to_excel(w,sheet_name='By objective',index=False)
    pd.DataFrame([dict(result_type=o,period=p,**d) for o,dd in M['result_type'].items() for p,d in dd.items()]).to_excel(w,sheet_name='By result type',index=False)
    for p in ['BEFORE','AFTER']:
        pd.DataFrame(M['campaigns'][p]).to_excel(w,sheet_name=f'Campaigns {p}',index=False)
        pd.DataFrame(M['adsets'][p]).to_excel(w,sheet_name=f'Ad sets {p}',index=False)
        pd.DataFrame(M['ads'][p]).to_excel(w,sheet_name=f'Ads {p}',index=False)
    pd.DataFrame([dict(theme=t,period=p,**d) for t,dd in M['theme_reach'].items() for p,d in dd.items()]).to_excel(w,sheet_name='Initiatives (Reach obj)',index=False)
pd.set_option('display.width',200); pd.set_option('display.float_format',lambda x:f'{x:,.4f}')
print(kpi.to_string())
