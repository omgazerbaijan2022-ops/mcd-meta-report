import pandas as pd, numpy as np, warnings
warnings.filterwarnings("ignore")
import os
D=os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data"))+"/"
FILES={"BEFORE":D+"mcd-april-june.xlsx","AFTER":D+"mcd-july-sept.xlsx"}
COLS=['campaign','adset','ad','result_type','results','spend','impr','reach','freq','cpm','link_clicks','adset2','ctr_all','cpc','post_eng','v3s','start','end','results_init']
def load(p):
    df=pd.read_excel(FILES[p],header=0)
    df.columns=COLS
    for c in ['results','spend','impr','reach','freq','cpm','link_clicks','ctr_all','cpc','post_eng','v3s']:
        df[c]=pd.to_numeric(df[c].replace('',np.nan),errors='coerce')
    df['result_type']=df['result_type'].fillna('').astype(str).str.strip()
    df['clicks_all']=df.ctr_all*df.impr/100
    df['period']=p
    return df
B=load("BEFORE"); A=load("AFTER")
def agg(df):
    s=df.spend.sum(); i=df.impr.sum(); r=df.reach.sum(); lc=df.link_clicks.sum(); ca=df.clicks_all.sum()
    return dict(ads=len(df),spend=s,impr=i,reach_sum=r,freq=i/r if r else np.nan,cpm=s/i*1000 if i else np.nan,
        link_clicks=lc,link_ctr=lc/i*100 if i else np.nan,ctr_all=ca/i*100 if i else np.nan,cpc=s/lc if lc else np.nan,
        post_eng=df.post_eng.sum(),cpe=s/df.post_eng.sum() if df.post_eng.sum() else np.nan,v3s=df.v3s.sum(),
        cost_per_1k_reach=s/r*1000 if r else np.nan, eng_rate=df.post_eng.sum()/i*100, v3s_rate=df.v3s.sum()/i*100,
        cost_per_1k_v3s=s/df.v3s.sum()*1000 if df.v3s.sum() else np.nan)
