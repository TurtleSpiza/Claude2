"""b51_prepare.py - gate + screen + restate + match the qpower_historical_20260818 corpus (Batch 42).
Also restates the v45 Q Power invoice/due dates from the retained Batch 40 corpus text (F1 repair)."""
import json, re, collections, hashlib
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
from python_calamine import CalamineWorkbook
W='/mnt/user-data/uploads/PS_WP_Transaction_Register_3FY_v50_HANDOVER.xlsx'
CORPUS='/mnt/user-data/uploads/Download_the_Q_Power_JSON_corpus.json'
raw=open(CORPUS,'rb').read(); SHA=hashlib.sha256(raw).hexdigest(); MD5=hashlib.md5(raw).hexdigest()
assert SHA=='296dbe390a720a1bce5b7f90e88b5b9324fe4aca59093552c9bd2a90b98b6103', SHA
d=json.loads(raw); DOCS=d['documents']; MAN=d['manifest']
prim=[x for x in DOCS if not x.get('duplicate_of')]; dups=[x for x in DOCS if x.get('duplicate_of')]
assert len(prim)==51 and len(dups)==0
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
def rh(x): return float(Decimal(str(x)).quantize(Decimal('0.01'),rounding=ROUND_HALF_UP))
def txt(x): return [' '.join(str(l.get('line_text') or '').split()) for l in x['lines'] if l['source']=='TEXT']
def ocr(x): return [' '.join(str(l.get('line_text') or '').split()) for l in x['lines'] if l['source']=='OCR']
G=collections.Counter()
# ---- gates
for x in prim:
    sub,gst,tot=float(x['printed_subtotal_ex_gst']),float(x['printed_gst']),float(x['printed_total_incl_gst'])
    assert abs(rh(sub/10)-gst)<0.005,(x['invoice_no'],'gst')
    assert abs(sub+gst-tot)<0.005,(x['invoice_no'],'tot')
    T=txt(x); J='\n'.join(T)
    assert re.search(r'Sub-?Total ex GST\s*\$?%s'%re.escape(f'{sub:.2f}'),J) or f'${sub:.2f}' in J, (x['invoice_no'],'subtotal token')
    assert f'${tot:.2f}' in J, (x['invoice_no'],'total token')
    lines=[l for l in x['lines'] if l['line_type']=='PRICED']
    if x['self_tie']=='TIE':
        assert lines and abs(round(sum(float(l['line_ex_gst']) for l in lines),2)-sub)<0.005,x['invoice_no']
        assert all(l['source']=='TEXT' for l in lines); G['tie']+=1
    else:
        assert not lines; assert 'Part # Item' not in J, x['invoice_no']   # template genuinely prints no line table
        G['no_table']+=1
# rule 12 screen (bare id + every prefix on the sheet)
res=[(x['invoice_no'],c) for x in prim for c in [x['invoice_no']]+[p+'-'+x['invoice_no'] for p in PREFIX] if c in EVIDS]
assert not res,res
norm=lambda t:re.sub(r'\s+',' ',t).strip().lower()
h=collections.defaultdict(list)
for x in prim: h[hashlib.md5(norm(' '.join(txt(x))).encode()).hexdigest()].append(x['invoice_no'])
assert not [v for v in h.values() if len(v)>1]
# ---- header restatement from the retained TEXT rows
def dmy(s):
    dd,mm,yy=s.split('/'); return f'{int(dd)}-{["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][int(mm)-1]}-{yy}', f'{yy}-{mm}-{dd}'
LETTER={'8/27 Allgas Street','Slacks Creek QLD 4127','Tel. 07 3155 4225','qpower.com.au','ABN 82 067 507 591','Licence # 69807','PLEASE PAY BY AMOUNT INVOICE DATE','Logan City Council','PO Box 3226','Logan City DC QLD 4114','Description'}
ADD_AS_PRINTED=['Flindersia Riverside Park']   # printed name absent from Panel A; added as printed (v44 precedent), parent/alias flagged
FACTS={}; LINES={}
for x in prim:
    k=x['invoice_no']; T=txt(x); J='\n'.join(T); tot=float(x['printed_total_incl_gst'])
    m=next((re.search(r'(\d\d/\d\d/\d{4})\s+\$([\d,]+\.\d\d)\s+(\d\d/\d\d/\d{4})',t) for t in T if re.search(r'(\d\d/\d\d/\d{4})\s+\$([\d,]+\.\d\d)\s+(\d\d/\d\d/\d{4})',t)),None)
    assert m,k; due,amt,inv=m.groups(); assert abs(float(amt.replace(',',''))-tot)<0.005,(k,'hdr amount')
    inv_lbl,inv_iso=dmy(inv); due_lbl,due_iso=dmy(due)
    f1_date = (x['invoice_date']!=inv_lbl)
    assert f'TAX INVOICE NO. {k}' in J,k
    # Site / Site Address / Order No blocks (rows between labels)
    idx={lab:next((i for i,t in enumerate(T) if t.startswith(lab)),None) for lab in ('Site:','Site Address:','Order No.:','Description','Ordered By:')}
    def block(a,b):
        if idx[a] is None: return None
        i=idx[a]; j=idx[b] if idx[b] is not None else i+1
        rows=[T[i][len(a):].strip()]+[T[r] for r in range(i+1,j) if T[r]]
        for r_ in rows:
            if r_: assert r_.lower() in J.lower(),(k,a,r_)
        return ' '.join(r for r in rows if r).strip() or None
    site=block('Site:','Site Address:'); addr=block('Site Address:','Order No.:'); order=block('Order No.:','Description')
    if order and idx['Description'] is None: order=T[idx['Order No.:']][len('Order No.:'):].strip()
    cr=None; po=None
    if order:
        mm=re.search(r'CR\s*#?\s*(\d{6,8})',order); cr=('CR'+mm.group(1)) if mm else None
        mm=re.search(r'PO\s*(\d{6,7})',order) or re.search(r'\b(7\d{5})\b',order); po=('PO'+mm.group(1)) if mm else None
    if not cr:
        mm=re.search(r'CR\s*#?\s*(\d{6,8})',J); cr=('CR'+mm.group(1)) if mm else None
    # Description block = Work Requested rows down to Ordered By / first Work Note / table header
    dstart=idx['Description']; desc=None
    if dstart is not None:
        rows=[]
        for r in range(dstart+1,len(T)):
            t=T[r]
            if t.startswith('Ordered By:') or ' - Work Note' in t or t.startswith('Part # Item') or t.startswith('Service - ') or t.startswith('Page ') or t.startswith('Sub-Total') or t.startswith('For electrical'): break
            if not t or t in LETTER or re.match(r'^(\d\d/\d\d/\d{4})\s+\$',t) or t.startswith('TAX INVOICE NO'): continue
            rows.append(t)
        desc=' | '.join(rows) if rows else None
    offr=None
    if idx['Ordered By:'] is not None: offr=T[idx['Ordered By:']][len('Ordered By:'):].strip() or None
    # Work notes: every row from the first "- Work Note" header down to the table/totals, letterhead rows dropped
    wn=[i for i,t in enumerate(T) if re.search(r'\(\d\d/\d\d/\d{2,4}\)\s*-\s*Work Note',t)]
    works=None; techs=[]
    if wn:
        rows=[]
        for r in range(wn[0],len(T)):
            t=T[r]
            if t.startswith('Part # Item') or t.startswith('Service - ') or t.startswith('Sub-Total') or t.startswith('For electrical'): break
            if t in LETTER or re.match(r'^(\d\d/\d\d/\d{4})\s+\$',t) or t.startswith('TAX INVOICE NO') or t.startswith('Page ') or not t: continue
            rows.append(t)
            mm=re.match(r'^(.+?)\s*\(\d\d/\d\d/\d{2,4}\)\s*-\s*Work Note',t)
            if mm and mm.group(1).strip() not in techs: techs.append(mm.group(1).strip())
        works=' | '.join(rows)
    pk=None; mm=re.search(r'\b(PK\s?\d{6,7})\b',J); pk=mm.group(1).replace(' ','') if mm else None
    svc=next((t for t in T if t.startswith('Service - ')),None)
    gsites=[]
    if site:
        for mm2 in GP.finditer(site):
            pre=site[:mm2.start()]
            if pre=='' or re.search(r'(/|,|-|\bat|\()\s*$',pre): gsites.append(CANON.get(mm2.group(1).lower(),mm2.group(1)))
        for add in ADD_AS_PRINTED:
            if re.search(r'(?<![A-Za-z])'+re.escape(add)+r'(?![A-Za-z])',site,re.I): gsites=[add]
        gsites=sorted(set(gsites))
    pages=x['page_range']
    FACTS[k]=dict(vendor=x['supplier'],abn=x['supplier_abn'],inv_lbl=inv_lbl,inv_iso=inv_iso,due_lbl=due_lbl,due_iso=due_iso,
        corpus_date=x['invoice_date'],f1_date=f1_date,site=site,addr=addr,order=order,cr=cr,po=po,officer=offr,desc=desc,works=works,
        techs=techs,pk=pk,svc=svc,gaz_sites=gsites,sub=float(x['printed_subtotal_ex_gst']),gst=float(x['printed_gst']),tot=tot,
        pages=f"pages {pages[0]} to {pages[1]} of {x['source_file']}",npages=pages[1]-pages[0]+1,tie=x['self_tie'],notes=x.get('notes'))
    for fld in ('officer',):
        v=FACTS[k][fld]
        if v: assert v[:20].lower() in J.lower(),(k,fld,v)
    # ---- lines
    if x['self_tie']=='TIE':
        # table blocks: rows after 'Part # Item ...' until totals/footer; F12 rule = NARRA above + PRICED + NARRA below
        seq=[]; inblk=False
        for l in x['lines']:
            if l['source']!='TEXT' or l['line_type']=='BLANK': continue
            t=' '.join(str(l['line_text']).split())
            if t.startswith('Part # Item'): inblk=True; continue
            if inblk and (t.startswith('Sub-Total') or t.startswith('For electrical') or t.startswith('Page ') or t.startswith('GST $') or t.startswith('Total $')): inblk=False; continue
            if inblk and t: seq.append((l['line_type'],t,l))
        out=[]; used=set(); i=0
        pi=[j for j,s in enumerate(seq) if s[0]=='PRICED']
        for n,j in enumerate(pi,1):
            ty,t,l=seq[j]
            core=re.sub(r'\s*[\d.,]+\s*\$[\d.,]+\s*10%\s*\$[\d.,]+$','',t).strip()
            rows=[]
            if j-1>=0 and seq[j-1][0]!='PRICED' and (j-1) not in used: rows.append(seq[j-1][1]); used.add(j-1)
            if core: rows.append(core)
            if rows and rows[0]!=core and j+1<len(seq) and seq[j+1][0]!='PRICED' and (j+1) not in used: rows.append(seq[j+1][1]); used.add(j+1)
            elif not core and j+1<len(seq) and seq[j+1][0]!='PRICED' and (j+1) not in used: rows.append(seq[j+1][1]); used.add(j+1)
            for r_ in rows: assert r_.lower() in J.lower(),(k,r_)
            desc_=' | '.join(rows); assert desc_,(k,'empty desc',t)
            up=float(l['unit_price_ex_gst']); q=float(l['qty']); a=round(float(l['line_ex_gst']),2)
            assert abs(round(q*up,2)-a)<0.011,(k,q,up,a)
            out.append(dict(n=n,desc=desc_,desc_rows=rows,qty=q,unit=up,gst='10%',amt=a,src='TEXT'))
        assert len(used)==len([s for s in seq if s[0]!='PRICED']),(k,'unconsumed narrative rows in table',[s[1] for j,s in enumerate(seq) if s[0]!='PRICED' and j not in used])
        assert abs(round(sum(o['amt'] for o in out),2)-FACTS[k]['sub'])<0.005,k
    else:
        assert desc,(k,'no description block')
        out=[dict(n=1,desc=desc,desc_rows=[desc],qty=None,unit=None,gst='(not printed)',amt=FACTS[k]['sub'],src='SINGLE_SCOPE')]
    LINES[k]=out
    if f1_date: G['f1_date']+=1
# ---- register match (amount corroboration mandatory; unmatched -> rule 16 only)
df=pd.read_excel(W,sheet_name='Register',header=3,engine='calamine').iloc[:21469]
ref=df.iloc[:,7].astype(str).str.strip().str.replace(r'\.0$','',regex=True)
amt=pd.to_numeric(df.iloc[:,19],errors='coerce').fillna(0)
MATCH={}; UNM=[]
for x in prim:
    k=x['invoice_no']; sub=FACTS[k]['sub']
    m=df[ref==k]
    if len(m)==0: UNM.append(k); continue
    nz=[i for i in m.index if abs(float(amt[i])-sub)<0.02]; assert len(nz)==1,(k,[float(amt[i]) for i in m.index])
    zero=[i for i in m.index if abs(float(amt[i]))<0.005]; assert len(nz)+len(zero)==len(m),k
    assert str(df.iloc[nz[0],87]) in ('nan','None',''),(k,'target already carries an EvID')
    assert str(df.iloc[nz[0],32])!='Confirmed',(k,'target Confirmed')
    assert str(df.iloc[nz[0],8]) in ('PUR Cred Invoice','DIR Cred Invoice','Creditors invoices'),k
    docd=str(df.iloc[nz[0],3])[:10]; assert docd==FACTS[k]['inv_iso'],(k,docd,FACTS[k]['inv_iso'])
    MATCH[k]=dict(target_row=int(nz[0])+5,target_lk=str(df.iloc[nz[0],0]),companions=[dict(row=int(i)+5,lk=str(df.iloc[i,0])) for i in zero],
        cred=str(df.iloc[nz[0],10]),name=str(df.iloc[nz[0],11]),pk=str(df.iloc[nz[0],16]),status=str(df.iloc[nz[0],32]),fy=str(df.iloc[nz[0],4]),
        docdate=docd,amount=round(float(amt[nz[0]]),2),ttype=str(df.iloc[nz[0],8]))
# ---- v45 date repair from the Batch 40 retained text
b40=json.load(open('/home/claude/up/left/toolkit/corpus_qpower_brizsouth_20260818.json'))
b40raw=open('/home/claude/up/left/toolkit/corpus_qpower_brizsouth_20260818.json','rb').read()
assert hashlib.sha256(b40raw).hexdigest()=='e3a4a0fbaf05e644f144523c53ba5b0aeb750163ff9dce0d08902e92f8999416'
REPAIR={}
for x in b40['documents']:
    if not x['supplier'].startswith('Q Power') or x.get('duplicate_of'): continue
    T=[' '.join(str(l.get('line_text') or '').split()) for l in x['lines'] if l['source']=='TEXT']
    m=next((re.search(r'(\d\d/\d\d/\d{4})\s+\$([\d,]+\.\d\d)\s+(\d\d/\d\d/\d{4})',t) for t in T if re.search(r'(\d\d/\d\d/\d{4})\s+\$([\d,]+\.\d\d)\s+(\d\d/\d\d/\d{4})',t)),None)
    assert m; due,a,inv=m.groups(); assert abs(float(a.replace(',',''))-float(x['printed_total_incl_gst']))<0.005
    il,ii=dmy(inv); dl,di=dmy(due)
    if x['invoice_date']!=il: REPAIR[x['invoice_no']]=dict(inv_iso=ii,inv_lbl=il,due_iso=di,due_lbl=dl,corpus_date=x['invoice_date'])
print('gates PASS | primaries',len(prim),'| tie',G['tie'],'| no-table (single-scope)',G['no_table'],'| f1 dates',G['f1_date'],'| re-sightings 0 | matched',len(MATCH),'| unmatched',UNM,'| companions',sum(len(v['companions']) for v in MATCH.values()))
print('sites (gaz):',sum(1 for k in FACTS if FACTS[k]['gaz_sites']),'| no PK printed:',sum(1 for k in FACTS if not FACTS[k]['pk']),'| v45 date repairs',len(REPAIR))
print('sites not in gaz:',sorted({FACTS[k]['site'] for k in FACTS if not FACTS[k]['gaz_sites'] and FACTS[k]['site']}))
json.dump(dict(sha=SHA,md5=MD5,manifest=MAN,facts=FACTS,lines=LINES,match=MATCH,unmatched=UNM,repair=REPAIR),open('/home/claude/b51/b51_parse.json','w'),indent=1)
