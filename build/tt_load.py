import pandas as pd, numpy as np, re, os, warnings
warnings.filterwarnings("ignore")
TT_DIR=os.environ.get("TT_DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data"))
FILES={"BEFORE":os.path.join(TT_DIR,"tiktok-april-june.xlsx"),"AFTER":os.path.join(TT_DIR,"tiktok-july-sept.xlsx")}
COLS=['campaign','adgroup','ad','spend','reach','impr','cpm','clicks','ctr','cpc','results','cpr','freq','v2s','currency']
NUM=['spend','reach','impr','cpm','clicks','ctr','cpc','results','cpr','freq','v2s']
def load_raw(p):
    df=pd.read_excel(FILES[p],header=0,dtype=str)
    df.columns=COLS
    for c in ['campaign','adgroup','ad']: df[c]=df[c].astype(str).str.strip()
    tot=df[df.campaign.str.startswith('Total of')].copy()
    df=df[~df.campaign.str.startswith('Total of')].copy()
    for c in NUM:
        df[c]=pd.to_numeric(df[c].replace({'-':np.nan,'':np.nan}),errors='coerce')
        tot[c]=pd.to_numeric(tot[c].replace({'-':np.nan,'':np.nan}),errors='coerce')
    df['period']=p
    return df.reset_index(drop=True), tot.iloc[0]
def objective(r):
    c=r.campaign.lower()
    if 'appinstall' in c or 'apppromo' in c or 'app_promo' in c or 'app install' in c: return 'App promotion'
    if 'community' in c or 'взаимодействие' in c: return 'Community interaction'
    if c.startswith('video views'): return 'Video views'
    if r.results==r.reach and r.reach>0: return 'Reach'
    if 'reach' in c: return 'Reach'
    return 'Unlabelled (results ≠ reach)'
THEMES=[('FIFA / World Cup',r'fifa|world.?cup'),('Mixology',r'mixology|mcxcola'),('Happy Meal',r'happy.?meal|\bhm_|_hm_|^hm_'),('Chicken',r'chicken|toyuq|mcchicken'),('McCafé',r'cafe|cofe'),('Grimace',r'grimace'),('McFlurry',r'mcflurry'),('Spiderman GenZ',r'spiderman'),('Restaurant opening',r'opening'),('Stranger Things',r'stranger'),('Bizim Burger / Roll',r'bizim'),('Squishmallows',r'squishmallow'),('Generic',r'generic')]
GENERIC_CAMP=r"^(reach_(april|may|june|feb|march|jan)|community|взаимодействие|video views_)"
def theme(r):
    c=r.campaign.lower()
    src=r.ad.lower() if re.search(GENERIC_CAMP,c) else c+' '+r.ad.lower()
    if not re.search(GENERIC_CAMP,c):
        for n,p in THEMES:
            if re.search(p,c): return n
    for n,p in THEMES:
        if re.search(p,src): return n
    return 'Other'
def agg(g):
    s=g.spend.sum(); i=g.impr.sum(); r=g.reach.sum(); c=g.clicks.sum(); v=g.v2s.sum()
    d=lambda a,b,k=1: a/b*k if b else np.nan
    return dict(ads=int(len(g)),spend=s,impr=i,reach_sum=r,freq_adlevel=d(i,r),cpm=d(s,i,1000),clicks=c,ctr=d(c,i,100),cpc=d(s,c),
                v2s=v,v2s_rate=d(v,i,100),cpv2s=d(s,v),cost_1k_v2s=d(s,v,1000),cost_1k_reach=d(s,r,1000))
B_all,B_tot=load_raw("BEFORE"); A_all,A_tot=load_raw("AFTER")
for df in (B_all,A_all):
    df['active']=df.spend>0
    df['objective']=df.apply(objective,axis=1)
    df['theme']=df.apply(theme,axis=1)
B=B_all[B_all.active].copy(); A=A_all[A_all.active].copy()
