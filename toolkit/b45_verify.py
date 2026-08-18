from python_calamine import CalamineWorkbook
import json, collections
P=json.load(open('/home/claude/b45/b45_pos.json')); wb=CalamineWorkbook.from_path(P['final']); CT=P['CT']
fails=[]
def ok(n,c,x=''):
    print(f'  {"PASS" if c else "FAIL"}  {n}{("  -> "+str(x)[:260]) if (x!="" and not c) else ""}')
    if not c: fails.append(n)
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python()
vals=[]
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
ok('every new gb 3 checks TRUE',all(str(r[121])=='TRUE' and str(r[122])=='TRUE' and str(r[123])=='TRUE' for r in new),
   [(r[0],r[121],r[122],r[123]) for r in new if not(str(r[121])==str(r[122])==str(r[123])=='TRUE')][:4])
ok('Nature Detail not generic',all(len(str(r[24]))>12 for r in new))
ok('companions carry no EvID',all(data[r-5][87] in (None,'') for r in P['companions']))
eil=wb.get_sheet_by_name('Evidence_Invoice_Lines').to_python()
ok(f'EIL end {P["EIL_END"]}',eil[P['EIL_END']-1][0] not in (None,'') and (len(eil)<=P['EIL_END'] or eil[P['EIL_END']][0] in (None,'')))
rec=eil[P['R'][0]-1:P['R'][1]]
ok(f'recon rows {P["R"][1]-P["R"][0]+1}',len(rec)==P['R'][1]-P['R'][0]+1)
ok('every recon row TRUE',all(str(r[4]).strip()=='TRUE' for r in rec),[r[2] for r in rec if str(r[4]).strip()!='TRUE'][:5])
ids={str(r[2]) for r in rec}
ok('every new invoice has a recon row',all(k in ids for k in P['invs']))
st=wb.get_sheet_by_name('Sites').to_python()
ck=int(float(cfg['SITES_CHECK_ROW'])); chk=[c for c in st[ck-1] if isinstance(c,str) and c in ('TRUE','FALSE')]
ok('Sites tie 3x TRUE',len(chk)==3 and all(v=='TRUE' for v in chk),st[ck-1][:5])
th=wb.get_sheet_by_name('Themes').to_python(); a,b=[int(x) for x in str(cfg['THEMES_CONTROL_ROWS']).split(':')]
ok('Themes controls TRUE',all(str(r[1]).strip()=='TRUE' for r in th[a-1:b]))
cv=wb.get_sheet_by_name('Coverage').to_python(); ok('Coverage tie',any(isinstance(c,float) and abs(c-CT)<0.02 for c in cv[141]))
sa=wb.get_sheet_by_name('Site_Allocation').to_python(); satot=int(float(cfg['SITE_ALLOC_TOTAL_ROW']))
ok('Site_Allocation tie TRUE',str(sa[satot][7]).strip()=='TRUE',sa[satot][7])
ok('Config v45',str(cfg.get('WORKBOOK_VERSION'))=='v45')
ok('Config SIGHTED_COUNT',int(float(cfg['SIGHTED_COUNT']))==P['SIGHT'])
tm={str(r[0]) for r in wb.get_sheet_by_name('Theme_Map').to_python()[4:31]}
ok('every new Nature Category exists in Theme_Map',all(str(r[23]) in tm for r in new),{str(r[23]) for r in new}-tm)
ok('every new row has a Theme',all(str(r[25]).strip() not in ('','None') for r in new))
vb=wb.get_sheet_by_name('Vendor_Boilerplate').to_python(); keys={str(r[0]) for r in vb[4:] if r and r[0]}
cited={str(r[117]) for r in new}|{str(r[118]) for r in new}
ok('cited BP keys exist',cited<=keys,cited-keys)
err=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and c[:4] in ('#REF','#NAM','#VAL','#DIV','#N/A','#NUM','#NUL'): err[s]+=1
ok('no formula errors',not err,dict(err))
print(); print('FAILS:',fails if fails else 'none')
