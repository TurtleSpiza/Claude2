from python_calamine import CalamineWorkbook
import json, collections
P=json.load(open('/home/claude/b47/b47_pos.json')); wb=CalamineWorkbook.from_path(P['final']); CT=P['CT']
fails=[]
def ok(n,c,x=''):
    print(f'  {"PASS" if c else "FAIL"}  {n}{("  -> "+str(x)[:280]) if (x!="" and not c) else ""}')
    if not c: fails.append(n)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
c0,c1=[int(x) for x in str(cfg['EI_CONTROL_ROWS']).split(':')]
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python(); vals=[]
for row in ei[c0-1:c1]:
    v=[c for c in row if c not in (None,'')]; vals.append(v[-1] if v else None)
tru=[v for v in vals if isinstance(v,str)]; counts=[v for v in vals if isinstance(v,float)]
ok('EI 14 controls',len(vals)==14); ok('EI counts unchanged',counts[:2]==[float(cfg['SIGHTED_COUNT'])]*2,counts[:2])
ok('EI 12 exact TRUE',len(tru)==12 and all(v.strip()=='TRUE' for v in tru),vals)
rg=wb.get_sheet_by_name('Register').to_python(); data=rg[4:21473]
ok('control total unchanged',any(isinstance(c,float) and abs(c-CT)<0.005 for c in rg[21474]))
ok('21,469 rows',len(data)==21469)
ok('green blocks unchanged',sum(1 for r in data if r[87] not in (None,''))==P['gb'])
sv=[str(r[128]) if r[128] not in (None,'') else '' for r in data]
bv=[str(r[129]) if r[129] not in (None,'') else '' for r in data]
CL=set(P['CLASS']); MA=P['MA']
ok('every row carries a basis',all(bv))
ok('no (no site named)',not any(b=='(no site named)' for b in bv))
nw=sum(1 for b in bv if b==P['CR_BASIS'])
ok(f'CR-written rows = {P["n_write"]}',nw==P['n_write'],nw)
ok('every CR-written row carries a site',all(s for s,b in zip(sv,bv) if b==P['CR_BASIS']))
nf=sum(1 for b in bv if '[CR DISAGREES' in b)
ok(f'CR disagreement flags = {P["n_flag"]}',nf==P['n_flag'],nf)
ok('flagged rows keep their ORIGINAL invoice-derived basis',all(b.split(' [CR DISAGREES')[0] not in CL for b in bv if '[CR DISAGREES' in b))
ok('flagged rows keep a site',all(s for s,b in zip(sv,bv) if '[CR DISAGREES' in b))
fold=collections.defaultdict(set)
for s,b in zip(sv,bv):
    bb=b.split(' [CR DISAGREES')[0]
    if s and bb not in CL and bb!=MA: fold[s.lower()].add(s)
ok('no case-folded site collision',all(len(v)==1 for v in fold.values()),{k:v for k,v in fold.items() if len(v)>1})
st=wb.get_sheet_by_name('Sites').to_python()
chk=[c for c in st[P['CHK']-1] if isinstance(c,str) and c in ('TRUE','FALSE')]
ok('Sites tie 3x TRUE',len(chk)==3 and all(v=='TRUE' for v in chk),st[P['CHK']-1][:5])
tot=st[P['TOT']-1]; print(f'   Sites TOTAL lines {tot[1]:,.0f} $ {tot[2]:,.2f}')
pa=st[P['PA'][0]-1:P['PA'][1]]; labels=[str(r[0]) for r in pa]
ok(f'Panel A {P["parks"]} sites',len(pa)==P['parks']); ok('labels unique',len(set(labels))==len(labels))
regset=set(s for s,b in zip(sv,bv) if s and b.split(' [CR DISAGREES')[0] not in CL and b!=MA)
ok('Panel A covers every register site',regset<=set(labels),regset-set(labels))
ok('Panel B 394',len(st[P['PB'][0]-1:P['PB'][1]])==394)
xw=wb.get_sheet_by_name('Site_Crosswalk').to_python()
ok('Site_Crosswalk grew',len([r for r in xw[4:P['XW']] if r and r[0]])>7)
vs=wb.get_sheet_by_name('Vendor_Series').to_python()
ok('F82 alias on Vendor_Series',any('F82 Landscaping' in str(r[0]) for r in vs if r and r[0]))
th=wb.get_sheet_by_name('Themes').to_python(); a,b=[int(x) for x in str(cfg['THEMES_CONTROL_ROWS']).split(':')]
ok('Themes controls TRUE',all(str(r[1]).strip()=='TRUE' for r in th[a-1:b]))
cv=wb.get_sheet_by_name('Coverage').to_python(); ok('Coverage tie',any(isinstance(c,float) and abs(c-CT)<0.02 for c in cv[141]))
sa=wb.get_sheet_by_name('Site_Allocation').to_python(); ok('Site_Allocation tie TRUE',str(sa[int(float(cfg['SITE_ALLOC_TOTAL_ROW']))][7]).strip()=='TRUE')
ok('Config v47',str(cfg.get('WORKBOOK_VERSION'))=='v47')
err=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and c[:4] in ('#REF','#NAM','#VAL','#DIV','#N/A','#NUM','#NUL'): err[s]+=1
ok('no formula errors',not err,dict(err))
print(); print('FAILS:',fails if fails else 'none')
