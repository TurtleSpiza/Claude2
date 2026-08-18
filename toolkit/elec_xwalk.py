import pandas as pd, re, collections, json
W='/mnt/user-data/outputs/PS_WP_Transaction_Register_3FY_v43_HANDOVER.xlsx'
df=pd.read_excel(W,sheet_name='Register',header=3,engine='calamine').iloc[:21469]
amt=pd.to_numeric(df.iloc[:,19],errors='coerce').fillna(0)
basis=df.iloc[:,129].astype(str)
STY={'street':'st','road':'rd','drive':'dr','court':'ct','avenue':'ave','circuit':'cct','parkway':'pkwy','crescent':'cres','place':'pl','boulevard':'bvd','terrace':'tce','lane':'ln','close':'cl','way':'way','highway':'hwy','esplanade':'esp'}
SUB_ALIAS={'beenle':'beenleigh','cedar grov':'cedar grove','waterford w':'waterford west','prie':'priestdale','logan cent':'logan central','slacks':'slacks creek','jimboom':'jimboomba','browns pl':'browns plains','waterford q':'waterford','marsde':'marsden'}
def norm(s):
    s=str(s).lower(); s=re.sub(r'\b(qld|queensland)\b',' ',s); s=re.sub(r'\b4\d{3}\b',' ',s)
    s=re.sub(r'[^a-z0-9 ]',' ',s)
    for k,v in STY.items(): s=re.sub(r'\b'+k+r'\b',v,s)
    return re.sub(r'\s+',' ',s).strip()
def parse(s):
    """-> (number or '', street tokens, suburb guess)"""
    n=norm(s)
    m=re.search(r'\b(\d{1,5}(?:\s*-\s*\d{1,5})?)\s+([a-z][a-z ]+?\b(?:st|rd|dr|ct|ave|cct|pkwy|cres|pl|bvd|tce|ln|cl|way|hwy|esp))\b\s*(.*)$',n)
    if m: return m.group(1).replace(' ',''), m.group(2).strip(), m.group(3).strip()
    m=re.search(r'\b([a-z][a-z ]+?\b(?:st|rd|dr|ct|ave|cct|pkwy|cres|pl|bvd|tce|ln|cl|way|hwy|esp))\b\s*(.*)$',n)
    if m: return '', m.group(1).strip(), m.group(2).strip()
    return None
def subfix(t):
    t=t.strip()
    for k,v in SUB_ALIAS.items():
        if t.startswith(k) or t.endswith(k): return v
    return t.split(' lot ')[0].strip()

# ---- index 1: rates parcel narrations (Council internal): "parcelid-street-park"
idx=collections.defaultdict(set); src=collections.defaultdict(set)
for v in df[basis=='Rates parcel narration'].iloc[:,21].astype(str).unique():
    m=re.search(r'S07 - \d+-(.+?)-(.+)$',v)
    if not m: continue
    addr,park=m.group(1),m.group(2).strip()
    p=parse(addr)
    if not p: continue
    num,street,_=p
    key=(num,street); idx[key].add(park); src[key].add('rates parcel')
    idx[('',street)].add(park); src[('',street)].add('rates parcel')
# ---- index 2: strategy layer House_Add + Park_Name
sl=pd.read_excel('/mnt/user-data/uploads/Parks_Strategy_Layer.xlsx',header=5,engine='calamine')
sl=sl[sl.Park_Name.notna()&sl.House_Add.notna()]
for a,pk,sb in zip(sl.House_Add,sl.Park_Name,sl.Suburb):
    if str(a).strip() in ('<Null>','nan') or str(pk).strip() in ('<Null>','nan'): continue
    p=parse(str(a)+' '+str(sb))
    if not p: continue
    num,street,_=p; pk=str(pk).strip()
    idx[(num,street)].add(pk); src[(num,street)].add('strategy layer')
    idx[('',street)].add(pk); src[('',street)].add('strategy layer')
# ---- index 3: parks directory
dr=pd.read_csv('/mnt/user-data/uploads/Logan_City_Council_Parks_-6649927587222555696.csv')
for a,pk in zip(dr.ParkAddress,dr.ParkName):
    p=parse(str(a))
    if not p: continue
    num,street,_=p
    idx[(num,street)].add(pk); src[(num,street)].add('directory')
    idx[('',street)].add(pk); src[('',street)].add('directory')
print('index keys',len(idx))

# ---- electricity rows
el=df[basis=='Supply-point address only (address route pending)']
sup=el.iloc[:,41].astype(str); narr=el.iloc[:,21].astype(str)
res=[]
for i,(s,n,a) in enumerate(zip(sup,narr,amt[el.index])):
    # use the fuller narration segment after the account number when supply string is truncated
    seg=n.split('-',2)[-1] if n.count('-')>=2 else s
    txt=seg if len(seg)>len(s) else s
    p=parse(txt)
    if not p: res.append(dict(row=el.index[i]+5,text=s,amt=a,flag='RED',park='',why='address not parseable')); continue
    num,street,tail=p
    hit=None
    if num and (num,street) in idx and len(idx[(num,street)])==1:
        hit=list(idx[(num,street)])[0]; res.append(dict(row=el.index[i]+5,text=s,amt=a,flag='GREEN',park=hit,why=f'number+street exact in {"/".join(sorted(src[(num,street)]))}')); continue
    if num and (num,street) in idx: res.append(dict(row=el.index[i]+5,text=s,amt=a,flag='RED',park='; '.join(sorted(idx[(num,street)])),why='number+street resolves to several parks')); continue
    if ('',street) in idx and len(idx[('',street)])==1:
        hit=list(idx[('',street)])[0]; res.append(dict(row=el.index[i]+5,text=s,amt=a,flag='AMBER',park=hit,why=f'street only, unique across {"/".join(sorted(src[("",street)]))}')); continue
    if ('',street) in idx: res.append(dict(row=el.index[i]+5,text=s,amt=a,flag='RED',park='; '.join(sorted(idx[('',street)])),why='street serves several parks')); continue
    res.append(dict(row=el.index[i]+5,text=s,amt=a,flag='RED',park='',why='street not in any Council address record'))
R=pd.DataFrame(res)
print(R.groupby('flag').agg(lines=('amt','size'),dollars=('amt','sum')).round(0))
print('\nby distinct supply string:')
g=R.groupby(['text','flag','park','why']).agg(lines=('amt','size'),d=('amt','sum')).reset_index().sort_values('lines',ascending=False)
print(g.head(25).to_string(index=False,max_colwidth=42))
R.to_pickle('elec_res.pkl')
