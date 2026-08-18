from python_calamine import CalamineWorkbook
import json, collections
import os
P=json.load(open(os.environ.get('B48_POS','/home/claude/b48/b48_pos.json'))); wb=CalamineWorkbook.from_path(P['final']); CT=P['CT']
fails=[]
def ok(n,c,x=''):
    print(f'  {"PASS" if c else "FAIL"}  {n}{("  -> "+str(x)[:250]) if (x!="" and not c) else ""}')
    if not c: fails.append(n)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
c0,c1=[int(x) for x in str(cfg['EI_CONTROL_ROWS']).split(':')]
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python(); vals=[]
for row in ei[c0-1:c1]:
    v=[c for c in row if c not in (None,'')]; vals.append(v[-1] if v else None)
tru=[v for v in vals if isinstance(v,str)]
ok('EI 14 controls',len(vals)==14); ok('EI 12 exact TRUE',len(tru)==12 and all(v.strip()=='TRUE' for v in tru),vals)
rg=wb.get_sheet_by_name('Register').to_python(); data=rg[4:21473]
ok('control total unchanged',any(isinstance(c,float) and abs(c-CT)<0.005 for c in rg[21474]))
ok('21,469 rows',len(data)==21469)
ok('green blocks unchanged',sum(1 for r in data if r[87] not in (None,''))==P['gb'])
sv=[str(r[128]) if r[128] not in (None,'') else '' for r in data]
bv=[str(r[129]) if r[129] not in (None,'') else '' for r in data]
SUB=set(P['suburbs'])
MAB=P['MA']
ok('every row carries a basis',all(bv))
ok('no suburb written as a site',not [s for s in sv if s and ';' not in s and s.lower() in SUB],[s for s in sv if s and ';' not in s and s.lower() in SUB][:3])
bad=[s for s,b in zip(sv,bv) if ';' in s and b!=MAB and any(p.strip().lower() in SUB for p in s.split(';'))]
ok('no joined list contains a suburb (allocated-multi rows exempt: their split is fixed by printed lines)',not bad,bad[:3])
ok('no "(not printed)" written as a site',not any('not printed' in s.lower() for s in sv))
ok('no joined list reduces to one park',not [s for s in sv if ';' in s and len({p.strip() for p in s.split(';') if p.strip()})==1])
CL=set(P['CLASS']); MA=P['MA']
fold=collections.defaultdict(set)
for s,b in zip(sv,bv):
    if s and ';' not in s and b not in CL and b!=MA: fold[s.lower()].add(s)
ok('no case-folded site collision',all(len(v)==1 for v in fold.values()),{k:v for k,v in fold.items() if len(v)>1})
st=wb.get_sheet_by_name('Sites').to_python()
chk=[c for c in st[P['CHK']-1] if isinstance(c,str) and c in ('TRUE','FALSE')]
ok('Sites tie 3x TRUE',len(chk)==3 and all(v=='TRUE' for v in chk),st[P['CHK']-1][:5])
tot=st[P['TOT']-1]; print(f'   Sites TOTAL lines {tot[1]:,.0f} $ {tot[2]:,.2f}')
pa=st[P['PA'][0]-1:P['PA'][1]]; labels=[str(r[0]) for r in pa]
ok(f'Panel A {P["parks"]}',len(pa)==P['parks']); ok('labels unique',len(set(labels))==len(labels))
ok('no Panel A label is a suburb',not [l for l in labels if l.lower() in SUB])
eil=wb.get_sheet_by_name('Evidence_Invoice_Lines').to_python(); e=int(float(cfg['EIL_END']))
ok('every EIL data row has a site basis',all(len(r)>12 and r[12] not in (None,'') for r in eil[4:e]))
th=wb.get_sheet_by_name('Themes').to_python(); a,b=[int(x) for x in str(cfg['THEMES_CONTROL_ROWS']).split(':')]
ok('Themes controls TRUE',all(str(r[1]).strip()=='TRUE' for r in th[a-1:b]))
cv=wb.get_sheet_by_name('Coverage').to_python(); ok('Coverage tie',any(isinstance(c,float) and abs(c-CT)<0.02 for c in cv[141]))
sa=wb.get_sheet_by_name('Site_Allocation').to_python(); ok('Site_Allocation tie TRUE',str(sa[int(float(cfg['SITE_ALLOC_TOTAL_ROW']))][7]).strip()=='TRUE')
ok('Config version',str(cfg.get('WORKBOOK_VERSION')) in ('v48','v49','v50'))
for k in ('SUBURB_BLOCKLIST','GAZ_UNIVERSE','STORES_CATALOGUE_MATCHES','EIL_SITE_LINES'): ok(f'Config {k}',k in cfg)
err=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and c[:4] in ('#REF','#NAM','#VAL','#DIV','#N/A','#NUM','#NUL'): err[s]+=1
ok('no formula errors',not err,dict(err))
print(); print('FAILS:',fails if fails else 'none')
