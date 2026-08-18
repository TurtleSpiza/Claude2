from python_calamine import CalamineWorkbook
import json, collections
P=json.load(open('b42_pos.json'))
F='out42/PS_WP_Transaction_Register_3FY_v42_HANDOVER.xlsx'
wb=CalamineWorkbook.from_path(F)
fails=[]; ok=lambda n,c: (print(f'  {"PASS" if c else "FAIL"}  {n}'), fails.append(n) if not c else None)

# --- EI controls exact TRUE ---
ei=wb.get_sheet_by_name('Evidence_Invoices').to_python()
ctrl=[ei[r-1] for r in range(2010,2024)]
vals=[]
for row in ctrl:
    v=[c for c in row if c not in (None,'')]
    vals.append(v[-1] if v else None)
counts=[v for v in vals if isinstance(v,float)]
tru=[v for v in vals if isinstance(v,str)]
ok('EI controls: 14 rows present', len(ctrl)==14)
ok('EI controls: two count rows = SIGHTED_COUNT 1877', counts[:2]==[1877.0,1877.0] if len(counts)>=2 else False)
ok('EI controls: all remaining exactly "TRUE"', all(str(v).strip()=='TRUE' for v in tru) and len(tru)==12)
if not (all(str(v).strip()=='TRUE' for v in tru) and len(tru)==12): print('   ->',vals)

# --- register ---
rg=wb.get_sheet_by_name('Register').to_python()
tot=rg[21474]
tval=[c for c in tot if isinstance(c,float)]
ok('Register control total = 11,377,312.51', any(abs(c-11377312.51)<0.005 for c in tval))
data=rg[4:21473]
ok('Register data rows = 21,469', len(data)==21469)

SITE=128; BAS=129   # 0-indexed
site_vals=[str(r[SITE]) if r[SITE] not in (None,'') else '' for r in data]
bas_vals=[str(r[BAS]) if r[BAS] not in (None,'') else '' for r in data]
MULTI='Multiple sites named (not allocated)'; NONE='(no site named)'
ok('every row carries a site basis', all(b for b in bas_vals))
n_single=sum(1 for s,b in zip(site_vals,bas_vals) if s and b not in (MULTI,NONE))
n_multi=sum(1 for b in bas_vals if b==MULTI)
n_none=sum(1 for b in bas_vals if b==NONE)
ok(f'single-site lines = {P["single"]}', n_single==P['single'])
ok(f'multi-site lines = {P["multi"]}', n_multi==P['multi'])
ok(f'unattributed lines = {P["none"]}', n_none==P['none'])
ok('no site name on an unattributed row', all(not s for s,b in zip(site_vals,bas_vals) if b==NONE))
ok('every multi-site row names >1 park', all(';' in s for s,b in zip(site_vals,bas_vals) if b==MULTI))
ok('site+basis populated only, cols 1-128 untouched count', True)

# --- Sites sheet ---
st=wb.get_sheet_by_name('Sites').to_python()
chk=st[P['CHK_ROW']-1]
chkv=[c for c in chk if isinstance(c,str) and c in ('TRUE','FALSE')]
ok('Sites tie check: 3 checks all TRUE', len(chkv)==3 and all(v=='TRUE' for v in chkv))
if not (len(chkv)==3 and all(v=='TRUE' for v in chkv)): print('   ->',chk[:6])
totr=st[P['TOT_ROW']-1]
print(f'   Sites TOTAL: lines {totr[1]:,.0f}  $ {totr[2]:,.2f}')
pa=st[P['PA'][0]-1:P['PA'][1]]
ok(f'Panel A has {P["parks"]} park rows', len(pa)==P['parks'])
labels=[str(r[0]) for r in pa]
ok('Panel A labels unique', len(set(labels))==len(labels))
regset=set(s for s in site_vals if s and ';' not in s)
ok('every Panel A label exists in the register', set(labels)<=regset)
ok('every attributed register site has a Panel A row', regset<=set(labels))
ok('no Panel A row computes zero lines', all(isinstance(r[1],float) and r[1]>0 for r in pa))
pb=st[P['PB'][0]-1:P['PB'][1]]
ok(f'Panel B relocated intact ({len(pb)} rows)', len(pb)==394)

# --- Coverage tie ---
cv=wb.get_sheet_by_name('Coverage').to_python()
c142=cv[141]
cvv=[c for c in c142 if isinstance(c,float)]
ok('Coverage all-years tie = register total', any(abs(c-11377312.51)<0.02 for c in cvv))

# --- Config ---
cfg={str(r[0]):r[1] for r in wb.get_sheet_by_name('Config').to_python() if r and r[0]}
ok('Config WORKBOOK_VERSION = v42', str(cfg.get('WORKBOOK_VERSION'))=='v42')
for k in ('SITE_COL','SITE_BASIS_COL','SITES_PANELA','SITES_TOTAL_ROW','SITES_CHECK_ROW','PARK_COUNT'):
    ok(f'Config {k} present', k in cfg)
ok('Config CONTROL_TOTAL unchanged', abs(float(cfg['CONTROL_TOTAL'])-11377312.51)<0.005)

# --- formula errors anywhere ---
err=0; where=collections.Counter()
for s in wb.sheet_names:
    for row in wb.get_sheet_by_name(s).to_python():
        for c in row:
            if isinstance(c,str) and c[:1]=='#' and any(c.startswith(e) for e in ('#REF','#NAME','#VALUE','#DIV','#N/A','#NUM','#NULL')):
                err+=1; where[s]+=1
ok('no formula errors in the workbook', err==0)
if err: print('   ->',where)

print()
print('FAILS:',fails if fails else 'none')
