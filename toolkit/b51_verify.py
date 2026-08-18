from python_calamine import CalamineWorkbook
import json, collections, re, datetime
P=json.load(open('/home/claude/b51/b51_pos.json')); wb=CalamineWorkbook.from_path(P['final']); CT=P['CT']
fails=[]
def ok(n,c,x=''):
    print(f'  {"PASS" if c else "FAIL"}  {n}{("  -> "+str(x)[:300]) if (x!="" and not c) else ""}')
    if not c: fails.append(n)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python(); vals=[]
for row in ei[P['C'][0]-1:P['C'][1]]:
    v=[c for c in row if c not in (None,'')]; vals.append(v[-1] if v else None)
counts=[v for v in vals if isinstance(v,float)]; tru=[v for v in vals if isinstance(v,str)]
ok('EI 14 controls',len(vals)==14)
ok('EI counts = new SIGHTED',counts[:2]==[float(P['SIGHT'])]*2,counts[:2])
ok('EI 12 exact TRUE',len(tru)==12 and all(v.strip()=='TRUE' for v in tru),vals)
rg=wb.get_sheet_by_name('Register').to_python(); data=rg[4:21473]
ok('control total',any(isinstance(c,float) and abs(c-CT)<0.005 for c in rg[21474]))
ok('21,469 rows',len(data)==21469)
gb=[r for r in data if r[87] not in (None,'')]
ok(f'green blocks {P["SIGHT"]}',len(gb)==P['SIGHT'],len(gb))
new=[data[r-5] for r in P['gb_rows']]
ok('every new gb has EvID/status/tier',all(r[87] and str(r[32])=='Confirmed' and str(r[28])=='Tier 1' for r in new))
ok('every new gb 3 checks TRUE',all(str(r[121])=='TRUE' and str(r[122])=='TRUE' and str(r[123])=='TRUE' for r in new),[(r[0],r[121],r[122],r[123]) for r in new if not(str(r[121])==str(r[122])==str(r[123])=='TRUE')][:4])
ok('every sighted line 3 checks TRUE',all(str(r[121])=='TRUE' and str(r[122])=='TRUE' and str(r[123])=='TRUE' for r in gb),[r[0] for r in gb if not(str(r[121])==str(r[122])==str(r[123])=='TRUE')][:4])
ok('Nature Detail not generic',all(len(str(r[24]))>12 for r in new))
ok('unmatched EvIDs on no register row',not [r[0] for r in data if str(r[87]) in P['unm']])
tm={str(r[0]) for r in wb.get_sheet_by_name('Theme_Map').to_python()[4:31]}
ok('new Nature Category in Theme_Map',all(str(r[23]) in tm for r in new)); ok('new rows resolve a Theme',all(str(r[25]).strip() not in ('','None','Unmapped - add to Theme_Map') for r in new))
# date repair
rep=P['rep']; ok('13 date repairs on register',all(str(data[r-5][93])[:10]==rep[str(data[r-5][87])]['inv_iso'] and str(data[r-5][94])[:10]==rep[str(data[r-5][87])]['due_iso'] for r in P['rep_rows']))
eimap={str(r[0]):r for r in ei[4:P['EI_END']] if r and r[0] not in (None,'')}
ok('13 date repairs on EI',all(str(eimap[k][3])[:10]==d['inv_iso'] and str(eimap[k][4])[:10]==d['due_iso'] for k,d in rep.items()))
ok('every new EI date = register Doc Date',all(str(data[r-5][93])[:10]==str(data[r-5][3])[:10] for r in P['gb_rows']))
# EIL / recon
eil=wb.get_sheet_by_name('Evidence_Invoice_Lines').to_python()
ok(f'EIL end {P["EIL_END"]}',eil[P['EIL_END']-1][0] not in (None,'') and (len(eil)<=P['EIL_END'] or eil[P['EIL_END']][0] in (None,'')))
rec=eil[P['R'][0]-1:P['R'][1]]; ok(f'recon rows {P["RECN"]}',len(rec)==P['RECN'])
ok('every recon row TRUE',all(str(r[4]).strip()=='TRUE' for r in rec),[r[2] for r in rec if str(r[4]).strip()!='TRUE'][:5])
ids={str(r[2]) for r in rec}; ok('every new invoice has a recon row',all(k in ids for k in P['invs']))
ok('every EIL data row has a site basis',all(len(r)>12 and r[12] not in (None,'') for r in eil[4:P['EIL_END']]))
# citation limbs
rr={str(r[2]):i+P['R'][0] for i,r in enumerate(rec) if r[2] not in (None,'')}
evrows=collections.defaultdict(list)
for i,r in enumerate(data):
    if r[87] not in (None,''): evrows[str(r[87])].append(i+5)
bad=[]
for r in ei[4:P['EI_END']]:
    if not r or r[0] in (None,''): continue
    e=str(r[0]); v=str(r[12]); m=re.search(r'Evidence_Invoice_Lines row (\d+)',v)
    if not m or int(m.group(1))!=rr.get(e): bad.append((e,'recon',m.group(1) if m else None,rr.get(e)))
    m2=re.findall(r'[Rr]egister row(?:\(s\)|s)? ([\d, and]+|\(no register match\))',v)
    if e in evrows:
        cited={int(x) for s in m2 for x in re.findall(r'\d+',s)}
        if not cited or not cited<=set(evrows[e]): bad.append((e,'reg',m2,evrows[e]))
ok('both citation limbs (EI)',not bad,bad[:4])
bad2=[]
for i,r in enumerate(data):
    if r[87] in (None,''): continue
    m=re.findall(r'(?:reconciliation|recon) row (\d+)',str(r[27]),re.I)
    if not m or any(int(x)!=rr.get(str(r[87])) for x in m): bad2.append((i+5,r[87],m,rr.get(str(r[87]))))
ok('register col 28 recon citations',not bad2,bad2[:4])
# sites
sv=[str(r[128]) if r[128] not in (None,'') else '' for r in data]; bv=[str(r[129]) if r[129] not in (None,'') else '' for r in data]
SUB=set(P['suburbs']); CL=set(P['CLASS']); MA=P['MA']
ok('every row carries a basis',all(bv))
ok('no suburb written as a site',not [s for s in sv if s and ';' not in s and s.lower() in SUB])
AL=set(P['aliases'])-{'mt warren oval park'}
ok('no alias token remains in DY',not [s for s in sv if any(t.strip().lower() in AL for t in s.split(';'))],[s for s in sv if any(t.strip().lower() in AL for t in s.split(';'))][:4])
ok('Logan River Park label retired',not [s for s in sv if s=='Logan River Park' or 'Logan River Park' in [t.strip() for t in s.split(';')]])
ok('no joined list reduces to one park',not [s for s in sv if ';' in s and len({p.strip() for p in s.split(';') if p.strip()})==1])
fold=collections.defaultdict(set)
for s,b in zip(sv,bv):
    if s and ';' not in s and b not in CL and b!=MA: fold[s.lower()].add(s)
ok('no case-folded site collision',all(len(v)==1 for v in fold.values()),{k:v for k,v in fold.items() if len(v)>1})
st=wb.get_sheet_by_name('Sites').to_python(); chk=[c for c in st[P['CHK']-1] if isinstance(c,str) and c in ('TRUE','FALSE')]
ok('Sites tie 3x TRUE',len(chk)==3 and all(v=='TRUE' for v in chk),st[P['CHK']-1][:5])
tot=st[P['TOT']-1]; print(f'   Sites TOTAL lines {tot[1]:,.0f} $ {tot[2]:,.2f}')
pa=st[P['PA'][0]-1:P['PA'][1]]; labels=[str(r[0]) for r in pa]
ok(f'Panel A {P["parks"]}',len(pa)==P['parks']); ok('labels unique',len(set(labels))==len(labels)); ok('no Panel A label is a suburb',not [l for l in labels if l.lower() in SUB])
ok('Logan River Parklands in Panel A, Logan River Park not','Logan River Parklands' in labels and 'Logan River Park' not in labels)
th=wb.get_sheet_by_name('Themes').to_python(); a,b=[int(x) for x in str(cfg['THEMES_CONTROL_ROWS']).split(':')]
ok('Themes controls TRUE',all(str(r[1]).strip()=='TRUE' for r in th[a-1:b]))
cv=wb.get_sheet_by_name('Coverage').to_python(); ok('Coverage tie',any(isinstance(c,float) and abs(c-CT)<0.02 for c in cv[141]))
sa=wb.get_sheet_by_name('Site_Allocation').to_python(); ok('Site_Allocation tie TRUE',str(sa[int(float(cfg['SITE_ALLOC_TOTAL_ROW']))][7]).strip()=='TRUE')
ok('Config v51',str(cfg.get('WORKBOOK_VERSION'))=='v51'); ok('Config SIGHTED_COUNT',int(float(cfg['SIGHTED_COUNT']))==P['SIGHT'])
vb=wb.get_sheet_by_name('Vendor_Boilerplate').to_python(); keys={str(r[0]) for r in vb[4:] if r and r[0]}
cited={str(r[117]) for r in new}|{str(r[118]) for r in new}; ok('cited BP keys exist',cited<=keys,cited-keys)
xw=wb.get_sheet_by_name('Site_Crosswalk').to_python(); ok('Site_Crosswalk extended',len([r for r in xw[4:] if r and r[0]])>=P['xw'][1]-4)
err=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and c[:4] in ('#REF','#NAM','#VAL','#DIV','#N/A','#NUM','#NUL'): err[s]+=1
ok('no formula errors',not err,dict(err))
print(); print('FAILS:',fails if fails else 'none')
