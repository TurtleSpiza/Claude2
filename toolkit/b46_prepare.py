"""b46_prepare.py - gate + match five APLEDGER creditor histories; assemble site corrections."""
import pandas as pd, re, json, hashlib, collections
from python_calamine import CalamineWorkbook
W='/mnt/user-data/outputs/PS_WP_Transaction_Register_3FY_v45_HANDOVER.xlsx'
U='/mnt/user-data/uploads/'
FILES=[('Ledger_Accounts_Transactions_Table_-_2026-08-18T145642_894.xlsx','ELE032','Electrical Data & Security Services Pty Ltd'),
       ('Ledger_Accounts_Transactions_Table_-_2026-08-18T145511_609.xlsx','SEA029','Sea-Crete (per APLEDGER SEA029)'),
       ('Ledger_Accounts_Transactions_Table_-_2026-08-18T150751_527.xlsx','QPO001','Q Power (Qld) Pty Ltd'),
       ('Ledger_Accounts_Transactions_Table_-_2026-08-18T150639_854.xlsx','SEC022','Unidentified (APLEDGER SEC022, ABN 97 613 325 769)'),
       ('Ledger_Accounts_Transactions_Table_-_2026-08-18T151046_941.xlsx','HIB003','Unidentified (APLEDGER HIB003, ABN 30 306 518 188)')]
cw=CalamineWorkbook.from_path(W)
cfg={str(r[0]):r[1] for r in cw.get_sheet_by_name('Config').to_python() if r and r[0]}
G0,G1=[int(x) for x in str(cfg['SITES_GROUP_ROWS']).split(':')]
CL={str(r[0]) for r in cw.get_sheet_by_name('Sites').to_python()[G0-1:G1] if r and r[0]}
PA0,PA1=[int(x) for x in str(cfg['SITES_PANELA']).split(':')]
GAZ=[str(x[0]).strip() for x in cw.get_sheet_by_name('Sites').to_python()[PA0-1:PA1]]

r=pd.read_excel(W,sheet_name='Register',header=3,engine='calamine'); r=r[r.iloc[:,0].notna()]
amt=pd.to_numeric(r.iloc[:,19],errors='coerce').fillna(0)
reg=pd.DataFrame({'row':[i+5 for i in range(len(r))],'lk':r.iloc[:,0].astype(str),
  'ref':r.iloc[:,7].astype(str).str.strip().str.replace(r'\.0$','',regex=True),'a':amt.values,
  'v':r.iloc[:,11].fillna('Unidentified').astype(str).values,'c':r.iloc[:,10].fillna('').astype(str).values,
  'st':r.iloc[:,32].astype(str).values,'evid':r.iloc[:,87].fillna('').astype(str).str.strip().values,
  'bas':r.iloc[:,129].fillna('').astype(str).values,'site':r.iloc[:,128].fillna('').astype(str).values})

CRED=[]; TGT={}; skipped=collections.Counter(); held=[]
for f,acct,label in FILES:
    raw=open(U+f,'rb').read(); md5=hashlib.md5(raw).hexdigest()
    e=pd.read_excel(U+f,header=3,engine='calamine'); e=e[e.iloc[:,0].notna()]; e=e[e.iloc[:,0].astype(str)!='Reference']
    e=e[e.iloc[:,6].astype(str)=='PUR Cred Invoice']                       # payments excluded
    e['ref']=e.iloc[:,0].astype(str).str.strip().str.replace(r'\.0$','',regex=True)
    e['amt']=pd.to_numeric(e.iloc[:,10],errors='coerce')
    e=e[e.iloc[:,0].astype(str).str.strip()!='']                            # grand-total trap: blank reference row
    abn=e.iloc[:,20].astype(str).str.replace(r'\D','',regex=True)
    mode=abn[abn.str.len()==11].mode()
    ABN=' '.join([mode[0][:2],mode[0][2:5],mode[0][5:8],mode[0][8:]]) if len(mode) else '(not stated)'
    for _,x in e.iterrows():
        CRED.append([acct,label,ABN,x['ref'],str(x.iloc[5])[:10],str(x.iloc[6]),' '.join(str(x.iloc[7]).split())[:250],
                     float(x['amt']) if pd.notna(x['amt']) else 0.0,str(x.iloc[11])[:10],str(x.iloc[14]),md5])
    m=reg.merge(e[['ref','amt']],on='ref',how='inner')
    for _,q in m.iterrows():
        if abs(q.amt-q.a*1.1)>=0.02: held.append((acct,q.ref,round(q.a,2),round(q.amt,2))); continue
        if q.st=='Confirmed' and q.evid=='': pass                          # legacy Confirmed w/o rule 17 -> still skip (v35)
        if q.st=='Confirmed': skipped['confirmed']+=1; continue
        if q.evid!='': skipped['already sighted']+=1; continue
        if abs(q.a)<0.005: skipped['zero amount']+=1; continue
        TGT[q.row]=dict(row=int(q.row),lk=q.lk,ref=q.ref,acct=acct,label=label,abn=ABN,
                        amount=round(float(q.a),2),incl=round(float(q.amt),2),prev=q.v,bas=q.bas)
    print(f'{acct}: history {len(e)} invoices, ABN {ABN}; refs matched {len(m)}, targets {sum(1 for t in TGT.values() if t["acct"]==acct)}')
print('TOTAL targets',len(TGT),'$',round(sum(t['amount'] for t in TGT.values()),2),'| skipped',dict(skipped),'| held',len(held))

# ---------------- site corrections (user directives, 18-Aug-2026)
GAZ_REMOVE=['Shailer Park']                       # suburb, not a park (David Robinson's business address)
OVERRIDE={'99 Ewing Rd Woodridge':'Ewing Park','293 Logan St Eagleby':'Olivers Sports Complex',
          '71 FLAGSTONIAN DR':'Flagstone Regional Park','71 Flagstonian Dr':'Flagstone Regional Park'}
FACILITY=['Olivers Sports Complex','West Logan Netball Courts','Logan and District Services Club',
          'Historical Society Amenities (223 Main St)','Little Athletics Club (293 Logan St, Eagleby)']
GAZ_ADD=['Ewing Park','Flagstone Regional Park','Crestmead Park','Stoneleigh Reserve Park','Brough Place Park',
         'East Beaumont Park','Everleigh Park','Kurrajong Park','Kohlmann Park','Sebring Park','Quandong Park']
json.dump(dict(targets=list(TGT.values()),cred=CRED,held=held,skipped=dict(skipped),
               gaz_remove=GAZ_REMOVE,gaz_add=GAZ_ADD,facility=FACILITY,override=OVERRIDE),
          open('/home/claude/b46/b46_parse.json','w'),indent=1)
print('written b46_parse.json | creditor rows',len(CRED))
