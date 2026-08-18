from python_calamine import CalamineWorkbook
import json, collections
P=json.load(open('/home/claude/b43/b43_pos.json'))
wb=CalamineWorkbook.from_path(P['final'])
fails=[]
def ok(n,c,extra=''):
    print(f'  {"PASS" if c else "FAIL"}  {n}{("  -> "+str(extra)) if (extra and not c) else ""}')
    if not c: fails.append(n)
CT=P['CT']
# --- EI controls exact TRUE (positions from Config)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
c0,c1=[int(x) for x in str(cfg['EI_CONTROL_ROWS']).split(':')]
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python()
vals=[]
for row in ei[c0-1:c1]:
    v=[c for c in row if c not in (None,'')]; vals.append(v[-1] if v else None)
counts=[v for v in vals if isinstance(v,float)]; tru=[v for v in vals if isinstance(v,str)]
ok('EI controls: 14 rows',len(vals)==14)
ok('EI controls: two count rows = SIGHTED_COUNT',counts[:2]==[float(cfg['SIGHTED_COUNT'])]*2,counts[:2])
ok('EI controls: remaining 12 exactly "TRUE"',len(tru)==12 and all(str(v).strip()=='TRUE' for v in tru),vals)
# --- register
rg=wb.get_sheet_by_name('Register').to_python()
tot=rg[21474]; ok('Register control total unchanged',any(isinstance(c,float) and abs(c-CT)<0.005 for c in tot))
data=rg[4:21473]; ok('Register 21,469 data rows',len(data)==21469)
S=128;B=129
sv=[str(r[S]) if r[S] not in (None,'') else '' for r in data]
bv=[str(r[B]) if r[B] not in (None,'') else '' for r in data]
CL=set(P['CLASS_LABELS']); MULTI='Multiple sites named (not allocated)'
ok('every row carries a basis',all(bv))
ok('no row reads "(no site named)"',not any(b=='(no site named)' for b in bv))
n_single=sum(1 for s,b in zip(sv,bv) if s and b not in CL)
ok(f'single-site lines = {P["single"]}',n_single==P['single'],n_single)
ok(f'multi lines = {P["multi"]}',sum(1 for b in bv if b==MULTI)==P['multi'])
ok('every multi row names >1 park',all(';' in s for s,b in zip(sv,bv) if b==MULTI))
ok('class rows carry no park (except depot label)',all((not s) or b.startswith('Non-park') or b==MULTI for s,b in zip(sv,bv) if b in CL))
for lbl,n in P['classes'].items():
    ok(f'class "{lbl[:40]}" = {n}',sum(1 for b in bv if b==lbl)==n)
fold=collections.defaultdict(set)
for s,b in zip(sv,bv):
    if s and b not in CL: fold[s.lower()].add(s)
ok('no case-folded park label collision',all(len(v)==1 for v in fold.values()),{k:v for k,v in fold.items() if len(v)>1})
# untouched-columns spot check: sum of col T and count of green blocks
ok('green-block count unchanged vs v42 source',sum(1 for r in data if r[87] not in (None,''))==P['gb_count'])
ok('leaked label absent from register',not any(s=='Undertake General Park' for s in sv))
# --- Sites
st=wb.get_sheet_by_name('Sites').to_python()
chk=st[P['CHK_ROW']-1]; chkv=[c for c in chk if isinstance(c,str) and c in ('TRUE','FALSE')]
ok('Sites tie check 3x TRUE',len(chkv)==3 and all(v=='TRUE' for v in chkv),chk[:5])
totr=st[P['TOT_ROW']-1]; print(f'   Sites TOTAL lines {totr[1]:,.0f} $ {totr[2]:,.2f}')
pa=st[P['PA'][0]-1:P['PA'][1]]; labels=[str(r[0]) for r in pa]
ok(f'Panel A {P["parks"]} rows',len(pa)==P['parks']); ok('Panel A labels unique',len(set(labels))==len(labels))
regset=set(s for s,b in zip(sv,bv) if s and b not in CL)
ok('Panel A == register park set',set(labels)==regset,(set(labels)^regset))
ok('no Panel A row computes zero lines',all(isinstance(r[1],float) and r[1]>0 for r in pa))
grp=st[P['GRP'][0]-1:P['GRP'][1]]
ok('group rows = class labels',[str(r[0]) for r in grp]==P['CLASS_LABELS'])
ok('every group row count > 0',all(isinstance(r[1],float) and r[1]>0 for r in grp),[(r[0],r[1]) for r in grp])
pb=st[P['PB'][0]-1:P['PB'][1]]; ok('Panel B relocated (394 rows)',len(pb)==394)
# --- Themes
th=wb.get_sheet_by_name('Themes').to_python()
ctl=th[P['TH_C'][0]-1:P['TH_C'][1]]
ok('Themes: 6 controls exact TRUE',len(ctl)==6 and all(str(r[1]).strip()=='TRUE' for r in ctl),[(r[0][:40],r[1]) for r in ctl])
tt=th[P['TH_TOT']-1]; ok('Themes all-years total = control total',isinstance(tt[5],float) and abs(tt[5]-CT)<0.005,tt[5])
# --- Coverage
cv=wb.get_sheet_by_name('Coverage').to_python(); c142=cv[141]
ok('Coverage all-years tie',any(isinstance(c,float) and abs(c-CT)<0.02 for c in c142))
# --- Config
ok('Config version v43',str(cfg.get('WORKBOOK_VERSION'))=='v43')
for k in ('SITE_CLASS_LABELS','SITE_NONE_LINES','SITE_SOURCE_COLS','SITE_BANNED_COLS','THEMES_TOTAL_ROW','THEMES_CONTROL_ROWS','PARK_COUNT'): ok(f'Config {k}',k in cfg)
ok('Config CONTROL_TOTAL',abs(float(cfg['CONTROL_TOTAL'])-CT)<0.005)
ok('Config PARK_COUNT = Panel A',int(float(cfg['PARK_COUNT']))==len(pa))
# --- formula errors
err=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and c[:4] in ('#REF','#NAM','#VAL','#DIV','#N/A','#NUM','#NUL'): err[s]+=1
ok('no formula errors',not err,dict(err))
print(); print('FAILS:',fails if fails else 'none')
