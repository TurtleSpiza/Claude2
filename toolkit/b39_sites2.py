from python_calamine import CalamineWorkbook
import collections, re, json

wb=CalamineWorkbook.from_path('zipx/PS_WP_Transaction_Register_3FY_v41_HANDOVER.xlsx')
sh=wb.get_sheet_by_name('Register').to_python(); data=sh[4:21473]

RATES=re.compile(r'Internal Rates - S\d+ - (\d{6,8})-(.+)$')
FEAT=re.compile(r'(park|reserve|gardens?|wetlands?|forest|green|square|oval|sportsfield|conservation|bushland|parklands?)\s*$',re.I)
BLOCK={'general park','skate park','dog park','water park','car park','the park','memorial park','sports park',
       'district park','adventure park','park park','industrial park','business park','logan park','city park',
       'sporting complex','bmx park','all abilities park','local park','district sports park','neighbourhood park',
       'linear park','pump track park','park reserve','environmental reserve','bushland reserve','conservation reserve',
       'nature reserve','open space reserve','sports field','sports fields','logan reserve','park green'}

def canon(s):
    s=re.sub(r'\s*\([^)]*\)\s*',' ',str(s))
    s=re.sub(r'[^A-Za-z0-9 ]',' ',s)
    return re.sub(r'\s+',' ',s).strip().lower()

parcel={}
for r in data:
    m=RATES.search(str(r[21]))
    if not m: continue
    pid,rest=m.group(1),m.group(2).strip()
    parts=[p.strip() for p in rest.split('-') if p.strip()]
    if len(parts)>=2 and FEAT.search(parts[-1]) and canon(parts[-1]) not in BLOCK:
        parcel.setdefault(pid,(parts[-1],' '.join(parts[:-1])))
    else: parcel.setdefault(pid,(None,rest))

gaz={}
for pid,(pk,addr) in parcel.items():
    if pk:
        c=canon(pk)
        if len(c.split())>=2 and c not in BLOCK: gaz.setdefault(c,pk.strip())

NAMEPAT=re.compile(r'\b([A-Z][A-Za-z\'\-]+(?:\s+[A-Z][A-Za-z\'\-]+){0,3}\s+(?:Park|Reserve|Gardens?|Wetlands?|Forest|Parklands?|Oval|Bushland|Green|Square|Common))\b')
extra=collections.Counter()
for r in data:
    for idx in (105,107,41,21,22):
        t=str(r[idx])
        if not t or t=='None': continue
        for m in NAMEPAT.finditer(t.title() if t.isupper() else t): extra[m.group(1).strip()]+=1
for nm,ct in extra.items():
    c=canon(nm)
    if len(c.split())>=2 and c not in gaz and c not in BLOCK and ct>=3: gaz[c]=nm

# canonical label per canon key: prefer rates-gazetteer form, else best Title Case
def titlecase(s):
    small={'of','the','and','de','at','on','in'}
    ws=s.split()
    return ' '.join(w if (w.isupper() and len(w)<=3) else (w.lower() if (i>0 and w.lower() in small) else w.capitalize())
                    for i,w in enumerate(ws))
RATESNAMES={canon(v[0]) for v in parcel.values() if v[0]}
for c in list(gaz):
    if c not in RATESNAMES or gaz[c].isupper(): gaz[c]=titlecase(gaz[c])
_seen={}
for c,l in gaz.items():
    k=l.lower()
    assert k not in _seen or _seen[k]==c, f'case-collision {l}'
    _seen[k]=c
GAZRE=[(re.compile(r'(?<![a-z0-9])'+re.escape(c)+r'(?![a-z0-9])'),lbl)
       for c,lbl in sorted(gaz.items(),key=lambda x:-len(x[0]))]

def hits(text):
    if not text or text in ('None','(not printed)','(blank as printed)'): return []
    c=canon(text); found=[]; used=[]
    for rx,lbl in GAZRE:
        m=rx.search(c)
        if m and not any(m.start()<e and s<m.end() for s,e in used):
            found.append(lbl); used.append((m.start(),m.end()))
    return found

SRC=[(105,'Sighted invoice site (printed)'),(107,'Sighted invoice work description'),
     (41,'Electricity supply point'),(21,'GL narration'),(22,'Enquiry narration')]
out=[]; basis_ct=collections.Counter()
park_fy=collections.defaultdict(lambda:collections.defaultdict(float))
park_ln=collections.defaultdict(lambda:collections.defaultdict(int))
park_bas=collections.defaultdict(collections.Counter)
park_acct=collections.defaultdict(collections.Counter)
park_pk=collections.defaultdict(collections.Counter)
multi_amt=0.0; multi_n=0
for i,r in enumerate(data):
    rn=i+5; amt=float(r[19] or 0); fy=str(r[4])
    park=''; basis=''; allp=[]
    m=RATES.search(str(r[21]))
    if m and parcel.get(m.group(1),(None,))[0]:
        raw=parcel[m.group(1)][0]
        park=gaz.get(canon(raw),raw); basis='Rates parcel narration'; allp=[park]
    else:
        for idx,lbl in SRC:
            h=hits(str(r[idx]))
            if h:
                allp=h; basis=lbl
                park = h[0] if len(h)==1 else f'Multiple sites ({len(h)})'
                break
    if park.startswith('Multiple'): multi_amt+=amt; multi_n+=1
    basis_ct[basis or '(unattributed)']+=1
    if park and not park.startswith('Multiple'):
        park_fy[park][fy]+=amt; park_ln[park][fy]+=1; park_bas[park][basis]+=1
        park_acct[park][f"{r[13]} {r[14]}"]+=1; park_pk[park][str(r[16])]+=1
    out.append({'row':rn,'park':park,'basis':basis,'all':allp,'fy':fy,'amt':round(amt,2),
                'acct':str(r[13]),'acctname':str(r[14]),'pk':str(r[16]),'sec':str(r[2])})

tot=len(data); totamt=sum(float(r[19] or 0) for r in data)
single=[o for o in out if o['park'] and not o['park'].startswith('Multiple')]
sa=sum(o['amt'] for o in single)
print(f'gazetteer parks {len(gaz)}   distinct parks with spend {len(park_fy)}')
print(f'single-site lines  {len(single):>6,} ({len(single)/tot:5.1%})   $ {sa:>13,.2f} ({sa/totamt:5.1%})')
print(f'multi-site lines   {multi_n:>6,} ({multi_n/tot:5.1%})   $ {multi_amt:>13,.2f} ({multi_amt/totamt:5.1%})  -- not allocated')
un=tot-len(single)-multi_n; ua=totamt-sa-multi_amt
print(f'unattributed       {un:>6,} ({un/tot:5.1%})   $ {ua:>13,.2f} ({ua/totamt:5.1%})')
print()
for k,v in basis_ct.most_common(): print(f'  {k:38} {v:>7,}')
print()
print('TOP 30 PARKS BY SPEND (single-site attribution only)')
for p,d in sorted(park_fy.items(),key=lambda x:-sum(x[1].values()))[:30]:
    print(f'  {p[:40]:42} ${sum(d.values()):>11,.2f}  lines {sum(park_ln[p].values()):>5,}  PK {park_pk[p].most_common(1)[0][0]}  {park_bas[p].most_common(1)[0][0][:26]}')
json.dump({'rows':out,'park_fy':{k:dict(v) for k,v in park_fy.items()},'park_ln':{k:dict(v) for k,v in park_ln.items()},
           'park_bas':{k:dict(v) for k,v in park_bas.items()},'park_acct':{k:dict(v) for k,v in park_acct.items()},
           'park_pk':{k:dict(v) for k,v in park_pk.items()},'gaz':gaz,
           'parcel':{k:list(v) for k,v in parcel.items()}},open('b39_sites2.json','w'))
