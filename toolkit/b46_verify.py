from python_calamine import CalamineWorkbook
import json, collections
P=json.load(open('/home/claude/b46/b46_pos.json')); wb=CalamineWorkbook.from_path(P['final']); CT=P['CT']
fails=[]
def ok(n,c,x=''):
    print(f'  {"PASS" if c else "FAIL"}  {n}{("  -> "+str(x)[:280]) if (x!="" and not c) else ""}')
    if not c: fails.append(n)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
c0,c1=[int(x) for x in str(cfg['EI_CONTROL_ROWS']).split(':')]
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python(); vals=[]
for row in ei[c0-1:c1]:
    v=[c for c in row if c not in (None,'')]; vals.append(v[-1] if v else None)
counts=[v for v in vals if isinstance(v,float)]; tru=[v for v in vals if isinstance(v,str)]
ok('EI 14 controls',len(vals)==14)
ok('EI counts = SIGHTED (unchanged)',counts[:2]==[float(cfg['SIGHTED_COUNT'])]*2,counts[:2])
ok('EI 12 exact TRUE',len(tru)==12 and all(v.strip()=='TRUE' for v in tru),vals)
rg=wb.get_sheet_by_name('Register').to_python(); data=rg[4:21473]
ok('control total unchanged',any(isinstance(c,float) and abs(c-CT)<0.005 for c in rg[21474]))
ok('21,469 rows',len(data)==21469)
ok('green-block count unchanged',sum(1 for r in data if r[87] not in (None,''))==P['gb'])
CL=set(P['CLASS']); MA=P['MA']
sv=[str(r[128]) if r[128] not in (None,'') else '' for r in data]
bv=[str(r[129]) if r[129] not in (None,'') else '' for r in data]
ok('every row carries a basis',all(bv))
ok('no (no site named)',not any(b=='(no site named)' for b in bv))
ok('Shailer Park removed from the register',not any(s=='Shailer Park' for s in sv))
# identification rows
idr=set(P['ident_rows'])
bad=[r for r in idr if not(data[r-5][10] and data[r-5][11] and data[r-5][12] and str(data[r-5][28])=='Tier 1')]
ok(f'all {P["n_id"]} identified rows carry code/contractor/ABN/Tier 1',not bad,bad[:5])
ok('no identified row carries an Ev Invoice ID',all(data[r-5][87] in (None,'') for r in idr))
ok('no identified row is Confirmed',all(str(data[r-5][32])!='Confirmed' for r in idr))
ok('no identified row still reads Unidentified',all(not str(data[r-5][11]).startswith('Unidentified (FY') for r in idr))
fold=collections.defaultdict(set)
for s,b in zip(sv,bv):
    if s and b not in CL and b!=MA and b!=P['NEWFAC']: fold[s.lower()].add(s)
ok('no case-folded site label collision',all(len(v)==1 for v in fold.values()),{k:v for k,v in fold.items() if len(v)>1})
ok('crosswalk basis present on the register',any(b==P['XW_BASIS'] for b in bv))
ok('facility basis present',any(b==P['NEWFAC'] for b in bv))
# Sites
st=wb.get_sheet_by_name('Sites').to_python()
chk=[c for c in st[P['CHK']-1] if isinstance(c,str) and c in ('TRUE','FALSE')]
ok('Sites tie 3x TRUE',len(chk)==3 and all(v=='TRUE' for v in chk),st[P['CHK']-1][:5])
tot=st[P['TOT']-1]; print(f'   Sites TOTAL lines {tot[1]:,.0f} $ {tot[2]:,.2f}')
pa=st[P['PA'][0]-1:P['PA'][1]]; labels=[str(r[0]) for r in pa]
ok(f'Panel A {P["parks"]} sites',len(pa)==P['parks']); ok('labels unique',len(set(labels))==len(labels))
grp=st[P['GRP'][0]-1:P['GRP'][1]]; ok('group rows = class list',[str(r[0]) for r in grp]==[c for c in ([x for x in P['CLASS'] if x.startswith('(residue')]+[x for x in P['CLASS'] if not x.startswith('(residue')])])
ok('Panel B 394',len(st[P['PB'][0]-1:P['PB'][1]])==394)
# other sheets
xw=wb.get_sheet_by_name('Site_Crosswalk').to_python()
ok('Site_Crosswalk rows',len([r for r in xw[4:P['XW_END']] if r and r[0]])>=7)
cl=wb.get_sheet_by_name('Creditor_Lines').to_python()
ok(f'Creditor_Lines to {P["CL_END"]}',cl[P['CL_END']-1][0] not in (None,''))
ok('Config CL_DATA',str(cfg['CL_DATA'])==f'5:{P["CL_END"]}',cfg['CL_DATA'])
th=wb.get_sheet_by_name('Themes').to_python(); a,b=[int(x) for x in str(cfg['THEMES_CONTROL_ROWS']).split(':')]
ok('Themes controls TRUE',all(str(r[1]).strip()=='TRUE' for r in th[a-1:b]))
cv=wb.get_sheet_by_name('Coverage').to_python(); ok('Coverage tie',any(isinstance(c,float) and abs(c-CT)<0.02 for c in cv[141]))
sa=wb.get_sheet_by_name('Site_Allocation').to_python(); satot=int(float(cfg['SITE_ALLOC_TOTAL_ROW']))
ok('Site_Allocation tie TRUE',str(sa[satot][7]).strip()=='TRUE',sa[satot][7])
ok('Config v46',str(cfg.get('WORKBOOK_VERSION'))=='v46')
err=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and c[:4] in ('#REF','#NAM','#VAL','#DIV','#N/A','#NUM','#NUL'): err[s]+=1
ok('no formula errors',not err,dict(err))
print(); print('FAILS:',fails if fails else 'none')
