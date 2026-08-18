"""b51_elec.py - Origin collective account summaries A-FE35E476-022..-027 (FY2025/26 Sep-25 to Feb-26): parse the
CURRENT CHARGES - SUB ACCOUNT SUMMARY, reconcile every register (statement, NMI) group to a statement sub-account by
supply address prefix + amount, and report ties. Feeds b51_build (Data_Acquisition/Open_Items) and the analysis workbook."""
import subprocess, re, json, collections, hashlib
UP='/home/claude/up/elec/'
FILES={'A-FE35E476-022':'A-FE35E476-022-summary.pdf','A-FE35E476-023':'A-FE35E476-023-summary.pdf','A-FE35E476-024':'A-FE35E476-024-summary.pdf',
       'A-FE35E476-025':'A-FE35E476-025-summary.pdf','A-FE35E476-026':'A-FE35E476-026-summary (1).pdf','A-FE35E476-027':'A-FE35E476-027-summary.pdf'}
def money(s): return float(s.replace(',',''))
ROW=re.compile(r'^\s*(A-[0-9A-F]{8})\s+(\d{6,10})\s+(.*?)\s+(-?)\$([\d,]+\.\d\d)\s+(?:CR\s+)?(-?)\$([\d,]+\.\d\d)(\s+CR)?\s*$')
STMT={}
for stmt,f in FILES.items():
    raw=open(UP+f,'rb').read()
    txt=subprocess.run(['pdftotext','-layout',UP+f,'-'],capture_output=True,text=True).stdout
    lines=txt.split('\n')
    hdr=next(i for i,l in enumerate(lines) if 'CURRENT CHARGES - SUB ACCOUNT SUMMARY' in l)
    issue=re.search(r'Issue date\s*\n\s*(\d\d \w{3} \d\d)',txt).group(1)
    newch=re.search(r'Current charges \(incl GST of \$([\d,]+\.\d\d)\)[^\n]*?\$([\d,]+\.\d\d)',txt)
    tot_gst,tot_incl=money(newch.group(1)),money(newch.group(2))
    rows=[]; pend=[]
    for l in lines[hdr+1:]:
        if 'Total current charges' in l:
            m=re.search(r'\$([\d,]+\.\d\d)\s+\$([\d,]+\.\d\d)',l); tc_gst,tc_incl=money(m.group(1)),money(m.group(2)); break
        l=l.replace('\f',''); s=l.strip()
        if not s or s in ('Current','charges') or s.startswith('Sub account') or s.startswith('CURRENT CHARGES'):
            if not s: pend=[]
            continue
        m=ROW.match(l.replace('\f',''))
        if m:
            sub,inv,addr,ng,gst,nv,incl,cr=m.groups()
            addr=' '.join(pend+[addr.strip()]); pend=[]
            g=money(gst); v=money(incl)
            if cr or ng: g=-g
            if cr or nv: v=-v
            rows.append(dict(sub=sub,inv=inv,addr=' '.join(addr.split()),gst=g,incl=v,ex=round(v-g,2)))
        else:
            pend.append(s)
    n_acc=int(re.search(r'Number of accounts\s+(\d+)',txt).group(1))
    assert len(rows)==n_acc,(stmt,len(rows),n_acc)
    assert abs(sum(r['incl'] for r in rows)-tc_incl)<0.011,(stmt,sum(r['incl'] for r in rows),tc_incl)
    assert abs(sum(r['gst'] for r in rows)-tc_gst)<0.011,(stmt,'gst')
    assert abs(tc_incl-tot_incl)<0.011 and abs(tc_gst-tot_gst)<0.011,(stmt,'summary vs table')
    STMT[stmt]=dict(file=f,md5=hashlib.md5(raw).hexdigest(),sha256=hashlib.sha256(raw).hexdigest(),issue=issue,accounts=n_acc,total_incl=tc_incl,total_gst=tc_gst,rows=rows)
    print(stmt,issue,n_acc,'accounts, incl',f'{tc_incl:,.2f}','ex',f'{tc_incl-tc_gst:,.2f}')
# ---- register groups
data=json.load(open('/home/claude/reg_cache.json'))[:21469]
def norm(s):
    s=str(s).lower().replace('&','and'); s=re.sub(r'[^a-z0-9 ]',' ',s); s=re.sub(r'\b(qld|queensland)\b',' ',s)
    for a,b in {'street':'st','road':'rd','drive':'dr','court':'ct','avenue':'ave','circuit':'cct','parkway':'pkwy','crescent':'cres','place':'pl','boulevard':'bvd','terrace':'tce','lane':'ln','close':'cl','highway':'hwy','esplanade':'esp','parade':'pde'}.items(): s=re.sub(r'\b'+a+r'\b',b,s)
    return re.sub(r'\s+',' ',s).strip()
grp=collections.defaultdict(lambda:dict(amt=0.0,rows=[],addr=set(),pk=set(),na=set(),dz=collections.Counter(),dy=collections.Counter()))
for i,r in enumerate(data):
    st_=str(r[7])
    if st_ not in STMT: continue
    nmi=str(r[40]).strip()
    g=grp[(st_,nmi)]; g['amt']+=float(r[19] or 0); g['rows'].append(i+5); g['addr'].add(str(r[41]).strip()); g['pk'].add(str(r[16])); g['na'].add(str(r[13]))
    g['dz'][str(r[129])]+=1; g['dy'][str(r[128])]+=1
print('register groups',len(grp),'lines',sum(len(g['rows']) for g in grp.values()))
# ---- match: address prefix (register supply string is truncated) then amount
recon=[]; used=collections.Counter()
for (st_,nmi),g in grp.items():
    cands=[]
    for row in STMT[st_]['rows']:
        na=norm(row['addr'])
        for a in g['addr']:
            an=norm(a)
            if not an: continue
            toks=an.split(' ')
            if len(toks)>1 and len(toks[-1])<=2: an=' '.join(toks[:-1])   # the register supply string is cut at 40 chars: drop a truncated tail token
            if na.startswith(an) or an.startswith(na): cands.append(row); break
    amt_ok=[c for c in cands if abs(c['ex']-round(g['amt'],2))<=0.02]
    if len(amt_ok)==1: pick=amt_ok[0]; how='address prefix + amount tie'
    elif len(cands)==1: pick=cands[0]; how='address prefix (amount differs)'
    elif len(amt_ok)>1: pick=amt_ok[0]; how='address prefix + amount tie (several sub-accounts share the address, first taken)'
    else:
        # amount-only fallback where the truncated string does not match (e.g. an abbreviated register label)
        am=[row for row in STMT[st_]['rows'] if abs(row['ex']-round(g['amt'],2))<=0.02 and used[(st_,row['sub'])]==0]
        pick=am[0] if len(am)==1 else None; how='amount only (unique)' if pick else 'NO MATCH'
    if pick: used[(st_,pick['sub'])]+=1
    recon.append(dict(stmt=st_,nmi=nmi,reg_ex=round(g['amt'],2),lines=len(g['rows']),rows=g['rows'],reg_addr=' / '.join(sorted(g['addr'])),pk='/'.join(sorted(g['pk'])),na='/'.join(sorted(g['na'])),
        sub=pick['sub'] if pick else '',inv=pick['inv'] if pick else '',stmt_addr=pick['addr'] if pick else '(no statement sub-account matched)',stmt_ex=pick['ex'] if pick else None,stmt_incl=pick['incl'] if pick else None,
        tie=('TRUE' if pick and abs(pick['ex']-round(g['amt'],2))<=0.02 else 'FALSE'),how=how,dz=g['dz'].most_common(1)[0][0],dy=g['dy'].most_common(1)[0][0]))
ties=sum(1 for r in recon if r['tie']=='TRUE'); print('groups tied to a statement sub-account:',ties,'/',len(recon),'| $ tied',f"{sum(r['reg_ex'] for r in recon if r['tie']=='TRUE'):,.2f}",'| register total',f"{sum(r['reg_ex'] for r in recon):,.2f}")
print(collections.Counter(r['how'] for r in recon))
for r in recon:
    if r['tie']!='TRUE': print('  X',r['stmt'],r['nmi'],r['reg_ex'],r['reg_addr'][:60],'|',r['stmt_addr'][:50],r['stmt_ex'],r['how'])
# statement sub-accounts NOT on the Parks register (Council-wide account: expected)
for st_,S in STMT.items():
    on=sum(1 for r in recon if r['stmt']==st_ and r['tie']=='TRUE'); S['parks_groups']=on
    S['parks_ex']=round(sum(r['reg_ex'] for r in recon if r['stmt']==st_),2)
    print(st_,'parks share ex GST',f"{S['parks_ex']:,.2f}",'of statement ex',f"{S['total_incl']-S['total_gst']:,.2f}",'| sub-accounts on register',on,'of',S['accounts'])
json.dump(dict(stmt=STMT,recon=recon),open('/home/claude/b51/b51_elec.json','w'),indent=0,default=str)
