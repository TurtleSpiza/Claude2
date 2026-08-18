"""b45_prepare.py - gate + screen + match the qpower_brizsouth_20260818 corpus."""
import json, re, collections, hashlib
import pandas as pd
from python_calamine import CalamineWorkbook

W='/mnt/user-data/outputs/PS_WP_Transaction_Register_3FY_v44_HANDOVER.xlsx'
CORPUS='/mnt/user-data/uploads/Download_the_JSON_invoice_corpus.json'
raw=open(CORPUS,'rb').read()
SHA=hashlib.sha256(raw).hexdigest(); MD5=hashlib.md5(raw).hexdigest()
assert SHA=='e3a4a0fbaf05e644f144523c53ba5b0aeb750163ff9dce0d08902e92f8999416'
d=json.loads(raw); DOCS=d['documents']
prim=[x for x in DOCS if not x.get('duplicate_of')]; dups=[x for x in DOCS if x.get('duplicate_of')]
assert len(prim)==44 and len(dups)==2

wb=CalamineWorkbook.from_path(W)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python()
e0,e1=[int(x) for x in str(cfg['EI_DATA']).split(':')]
EVIDS={str(r[0]).strip() for r in ei[e0-1:e1] if r and r[0] not in (None,'')}
PREFIX=sorted({e.split('-')[0] for e in EVIDS if '-' in e and not e.split('-')[0].isdigit()})
pa0,pa1=[int(x) for x in str(cfg['SITES_PANELA']).split(':')]
GAZ=[str(x[0]).strip() for x in wb.get_sheet_by_name('Sites').to_python()[pa0-1:pa1]]
CANON={g.lower():g for g in GAZ}
GP=re.compile('('+'|'.join(re.escape(g) for g in sorted(GAZ,key=len,reverse=True) if len(g)>5)+')',re.I)

# ---- gates
def rounds(x): 
    from decimal import Decimal,ROUND_HALF_UP
    return float(Decimal(str(x)).quantize(Decimal('0.01'),rounding=ROUND_HALF_UP))
G=collections.Counter()
for x in prim:
    lines=[l for l in x['lines'] if l['line_type']=='PRICED']
    s=round(sum(float(l['line_ex_gst']) for l in lines),2)
    assert abs(s-float(x['printed_subtotal_ex_gst']))<0.005, (x['invoice_no'],s)
    assert abs(rounds(float(x['printed_subtotal_ex_gst'])/10)-float(x['printed_gst']))<0.005,(x['invoice_no'],'gst')
    # F1 (v45): printed_total_incl_gst holds the SUBTOTAL on all 31 BrizSouth invoices. Restate = subtotal + GST
    sub_,gst_=float(x['printed_subtotal_ex_gst']),float(x['printed_gst'])
    if abs(float(x['printed_total_incl_gst'])-sub_)<0.005 and gst_>0:
        x['printed_total_incl_gst']=round(sub_+gst_,2); x['_f1_total_restated']=True; G['f1_total']+=1
    assert abs(sub_+gst_-float(x['printed_total_incl_gst']))<0.005, x['invoice_no']
    assert all(l['source']=='TEXT' for l in lines), x['invoice_no']
    G['self_tie']+=1
# rule 12 prefix-aware screen (bare id + every prefix on the sheet)
resight=[]
for x in prim:
    idn=str(x['invoice_no']).strip()
    for c in [idn]+[p+'-'+idn for p in PREFIX]:
        if c in EVIDS: resight.append((idn,c))
assert not resight, resight
# corpus-internal duplicate screen (v40 rule)
norm=lambda t:re.sub(r'\s+',' ',t).strip().lower()
h=collections.defaultdict(list)
for x in prim: h[hashlib.md5(norm(' '.join(str(l.get('line_text') or '') for l in x['lines'])).encode()).hexdigest()].append(x['invoice_no'])
assert not [v for v in h.values() if len(v)>1], [v for v in h.values() if len(v)>1]

# ---- header restatement from the retained line_text
def txt(x): return [' '.join(str(l.get('line_text') or '').split()) for l in x['lines']]
def find(rows,pat,grp=1,after=0):
    for i,t in enumerate(rows):
        m=re.search(pat,t,re.I)
        if m:
            if after and i+after<len(rows): return rows[i+after]
            return m.group(grp).strip()
    return None
FACTS={}
for x in prim:
    T=txt(x); joined='\n'.join(T)
    qp = x['supplier'].startswith('Q Power')
    site = find(T,r'^Site:\s*(.+)$') or find(T,r'Job Location:\s*(.+)$')
    if site: site=re.sub(r'\s+',' ',site).strip(' .')
    addr = find(T,r'^Site Address:\s*(.+)$')
    order= find(T,r'^Ord?der No\.?:?\s*(.+)$') or find(T,r'^Order No\.?:?\s*(.+)$')
    cr   = find(T,r'Customer Request #\s*(\d+)') or (re.search(r'CR(\d{6,8})',order).group(0) if order and re.search(r'CR(\d{6,8})',order) else None)
    pk   = find(T,r'\b(PK\s?\d{6,7})\b')
    tech = find(T,r'^Technician:\s*$',after=1) if qp else None
    offr = find(T,r'^Ordered By:\s*$',after=1) if qp else find(T,r'^Ordered By:\s*(.+)$')
    if not offr: offr=find(T,r'^Ordered By:\s*(.+)$')
    works= None
    if qp:
        try:
            i=T.index('Works Completed:'); j=next(k for k in range(i+1,len(T)) if T[k].endswith(':') and len(T[k])<40)
            works=' | '.join(T[i+1:j])
        except Exception: works=None
    gsites=sorted({CANON.get(m.lower(),m) for m in GP.findall(joined)})
    FACTS[x['invoice_no']]=dict(site_printed=site,site_address=addr,order=order,cr=cr,pk=pk,tech=tech,officer=offr,
        works=works,gaz_sites=gsites,vendor=x['supplier'],abn=x['supplier_abn'],date=x['invoice_date'],
        sub=float(x['printed_subtotal_ex_gst']),gst=float(x['printed_gst']),tot=float(x['printed_total_incl_gst']),
        pages=f"pages {x['page_range'][0]} to {x['page_range'][1]} of {x['source_file']}", f1_total=bool(x.get('_f1_total_restated')),
        lines=[dict(desc=' '.join(str(l['line_text']).split()).rsplit('$',1)[0].strip() if False else None) for l in []])
    # verbatim spot-check (rule 19.2): every restated token present in the retained text
    for k in ('site_printed','order','pk','tech','officer'):
        v=FACTS[x['invoice_no']][k]
        if v: assert v.split(' | ')[0][:24].lower() in joined.lower(), (x['invoice_no'],k,v)

# ---- priced line capture (verbatim description = printed item text before the numeric tail)
LINES={}
for x in prim:
    out=[]
    for n,l in enumerate([l for l in x['lines'] if l['line_type']=='PRICED'],1):
        t=' '.join(str(l['line_text']).split())
        desc=re.sub(r'\s*[\d.,]+\s*\$[\d.,]+\s*10%\s*\$[\d.,]+$','',t)      # Q Power table
        desc=re.sub(r'\s*\$[\d.,]+$','',desc).strip()                        # BrizSouth amount tail
        assert desc and desc.lower() in ' '.join(txt(x)).lower(), (x['invoice_no'],desc)
        out.append(dict(n=n,desc=desc,qty=l.get('qty'),unit=l.get('unit_price_ex_gst'),
                        gst='10%' if x['supplier'].startswith('Q Power') else '(not printed)',
                        amt=round(float(l['line_ex_gst']),2)))
    assert abs(round(sum(o['amt'] for o in out),2)-FACTS[x['invoice_no']]['sub'])<0.005
    LINES[x['invoice_no']]=out

# ---- register match (amount corroboration mandatory)
df=pd.read_excel(W,sheet_name='Register',header=3,engine='calamine').iloc[:21469]
ref=df.iloc[:,7].astype(str).str.strip().str.replace(r'\.0$','',regex=True)
amt=pd.to_numeric(df.iloc[:,19],errors='coerce').fillna(0)
MATCH={}
for x in prim:
    idn=str(x['invoice_no']).strip(); sub=FACTS[idn]['sub']
    m=df[ref==idn]; assert len(m)>0, idn
    nz=[i for i in m.index if abs(float(amt[i])-sub)<0.02]
    assert len(nz)==1, (idn,[float(amt[i]) for i in m.index])
    zero=[i for i in m.index if abs(float(amt[i]))<0.005]
    assert len(nz)+len(zero)==len(m), idn
    MATCH[idn]=dict(target_row=int(nz[0])+5, target_lk=str(df.iloc[nz[0],0]),
                    companions=[dict(row=int(i)+5,lk=str(df.iloc[i,0])) for i in zero],
                    cred=str(df.iloc[nz[0],10]), name=str(df.iloc[nz[0],11]),
                    pk=str(df.iloc[nz[0],16]), status=str(df.iloc[nz[0],32]), amount=round(float(amt[nz[0]]),2))
print('F1 totals restated:',G['f1_total'])
print('gates PASS | primaries',len(prim),'| re-sightings 0 | matched',len(MATCH),'| companions',sum(len(v['companions']) for v in MATCH.values()))
json.dump(dict(sha=SHA,md5=MD5,facts=FACTS,lines=LINES,match=MATCH,
               dups=[dict(inv=x['invoice_no'],of=x['duplicate_of'],pages=x['page_range']) for x in dups]),
          open('/home/claude/b45/b45_parse.json','w'),indent=1)
print('written b45_parse.json')
