from python_calamine import CalamineWorkbook
import json, collections
P=json.load(open('/home/claude/b44/b44_pos.json')); wb=CalamineWorkbook.from_path(P['final']); CT=P['CT']
fails=[]
def ok(n,c,x=''):
    print(f'  {"PASS" if c else "FAIL"}  {n}{("  -> "+str(x)[:300]) if (x!="" and not c) else ""}')
    if not c: fails.append(n)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
c0,c1=[int(x) for x in str(cfg['EI_CONTROL_ROWS']).split(':')]
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python(); vals=[]
for row in ei[c0-1:c1]:
    v=[c for c in row if c not in (None,'')]; vals.append(v[-1] if v else None)
counts=[v for v in vals if isinstance(v,float)]; tru=[v for v in vals if isinstance(v,str)]
ok('EI 14 controls',len(vals)==14); ok('EI counts = SIGHTED',counts[:2]==[float(cfg['SIGHTED_COUNT'])]*2,counts[:2]); ok('EI 12 exact TRUE',len(tru)==12 and all(v.strip()=='TRUE' for v in tru),vals)
rg=wb.get_sheet_by_name('Register').to_python(); tot=rg[21474]
ok('control total',any(isinstance(c,float) and abs(c-CT)<0.005 for c in tot)); data=rg[4:21473]; ok('21,469 rows',len(data)==21469)
sv=[str(r[128]) if r[128] not in (None,'') else '' for r in data]; bv=[str(r[129]) if r[129] not in (None,'') else '' for r in data]
amt=[float(r[19]) if isinstance(r[19],(int,float)) else 0.0 for r in data]; lk=[str(r[0]) for r in data]
CL=set(P['CLASS_LABELS']); MA=P['MULTI_ALLOC']; MULTI='Multiple sites named (not allocated)'
ok('every row basis',all(bv)); ok('no (no site named)',not any(b=='(no site named)' for b in bv))
ok('green-block count unchanged',sum(1 for r in data if r[87] not in (None,''))==P['gb_count'])
n_single=sum(1 for s,b in zip(sv,bv) if s and b not in CL and b!=MA); ok(f'single {P["single"]}',n_single==P['single'],n_single)
ok(f'multi held {P["multi_held"]}',sum(1 for b in bv if b==MULTI)==P['multi_held'])
ok(f'multi allocated {P["alloc_lk"]}',sum(1 for b in bv if b==MA)==P['alloc_lk'])
ok(f'elec green {P["green"]}',sum(1 for b in bv if b==P['E_GREEN'])==P['green']); ok(f'elec amber {P["amber"]}',sum(1 for b in bv if b==P['E_AMBER'])==P['amber'])
ok('every green/amber elec row has a park',all(s for s,b in zip(sv,bv) if b in (P['E_GREEN'],P['E_AMBER'])))
# Site_Allocation ties per LineKey
sa=wb.get_sheet_by_name('Site_Allocation').to_python(); rows=sa[4:P['SA_END']]
ok(f'Site_Allocation {P["alloc_rows"]} rows',len(rows)==P['alloc_rows'],len(rows))
per=collections.Counter(); regamt={}
for r in rows: per[str(r[0])]+=float(r[7]); regamt[str(r[0])]=float(r[5])
amap=dict(zip(lk,amt))
bad=[k for k in per if abs(per[k]-amap.get(k,1e9))>0.005]
ok('every allocation LineKey ties to Register T',not bad,bad[:5])
ok('every allocation LineKey carries the allocated-multi basis',all(bv[lk.index(k)]==MA for k in per))
ok('every allocated-multi row is on Site_Allocation',all(k in per for k,b in zip(lk,bv) if b==MA))
tie=sa[P['SA_TOT']][7] if len(sa)>P['SA_TOT'] else None; ok('Site_Allocation tie check TRUE',str(tie).strip()=='TRUE',tie)
ok('SA total = allocated',isinstance(sa[P['SA_TOT']-1][7],float) and abs(sa[P['SA_TOT']-1][7]-P['alloc_total'])<0.01,sa[P['SA_TOT']-1][7])
# EIL L/M
eil=wb.get_sheet_by_name('Evidence_Invoice_Lines').to_python()
ok('EIL header L/M',str(eil[3][11]).startswith('Site') and str(eil[3][12]).startswith('Site basis'))
ok('EIL every data row has a site basis',all(len(r)>12 and r[12] not in (None,'') for r in eil[4:P['eil_end']]))
# Sites
st=wb.get_sheet_by_name('Sites').to_python(); chk=st[P['CHK_ROW']-1]; chkv=[c for c in chk if isinstance(c,str) and c in('TRUE','FALSE')]
ok('Sites tie 3x TRUE',len(chkv)==3 and all(v=='TRUE' for v in chkv),chk[:5])
totr=st[P['TOT_ROW']-1]; print(f'   Sites TOTAL lines {totr[1]:,.0f} $ {totr[2]:,.2f}')
pa=st[P['PA'][0]-1:P['PA'][1]]; labels=[str(r[0]) for r in pa]
ok(f'Panel A {P["parks"]}',len(pa)==P['parks']); ok('labels unique',len(set(labels))==len(labels))
ok('no Panel A row computes zero $ AND zero lines',all((isinstance(r[1],float) and r[1]>0) or (isinstance(r[2],float) and abs(r[2])>0) for r in pa))
regset=set(s for s,b in zip(sv,bv) if s and b not in CL and b!=MA)|{str(r[6]) for r in rows if str(r[6])!=P['RESIDUE']}
ok('Panel A == park universe',set(labels)==regset,set(labels)^regset)
grp=st[P['GRP'][0]-1:P['GRP'][1]]; ok('group rows count',len(grp)==1+len(P['CLASS_LABELS']))
ok('Panel B 394',len(st[P['PB'][0]-1:P['PB'][1]])==394)
th=wb.get_sheet_by_name('Themes').to_python(); tc0,tc1=[int(x) for x in str(cfg['THEMES_CONTROL_ROWS']).split(':')]
ok('Themes controls TRUE',all(str(r[1]).strip()=='TRUE' for r in th[tc0-1:tc1]))
cv=wb.get_sheet_by_name('Coverage').to_python(); ok('Coverage tie',any(isinstance(c,float) and abs(c-CT)<0.02 for c in cv[141]))
ok('Config v44',str(cfg.get('WORKBOOK_VERSION'))=='v44')
for k in ('SITE_ALLOC_DATA','SITE_ALLOC_TOTAL_ROW','EIL_SITE_COLS','ELEC_ADDR_GREEN','ELEC_ADDR_AMBER'): ok(f'Config {k}',k in cfg)
err=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and c[:4] in('#REF','#NAM','#VAL','#DIV','#N/A','#NUM','#NUL'): err[s]+=1
ok('no formula errors',not err,dict(err))
print(); print('FAILS:',fails if fails else 'none')
