import csv, collections, re, json
from python_calamine import CalamineWorkbook

FILES={'A-FE35E476-021':'A-FE35E476-021-breakdown.csv','A-FE35E476-028':'A-FE35E476-028-breakdown.csv',
       'A-FE35E476-029':'A-FE35E476-029-breakdown.csv','A-FE35E476-030':'A-FE35E476-030-breakdown__1_.csv',
       'A-FE35E476-031':'A-FE35E476-031-breakdown.csv'}
UP='/mnt/user-data/uploads/'

# ---- energy statement site totals + detail ----
site={}   # (stmt,nmi) -> dict
detail=collections.defaultdict(list)
for stmt,f in FILES.items():
    for r in csv.DictReader(open(UP+f, newline='', encoding='utf-8-sig')):
        nmi=(r['NMI : MIRN : DPI'] or '').strip()
        if not nmi: continue
        clean=lambda s:(s or '').replace('\n',' ').replace('\r',' ').strip()
        if (r['Charge Type'] or '')=='' and (r['Charge Description'] or '')=='TOTAL':
            v=(r['Total Current Account Charges (GST excl.)'] or '').replace(',','').strip()
            try: x=float(v)
            except: continue
            site[(stmt,nmi)]={'stmt':stmt,'nmi':nmi,'sub':r['Sub Account Number'],'inv':r['Invoice Number'],
                'addr':clean(r['Service Address']),'supplier':clean(r['Supplier']) or '(not printed)',
                'abn':clean(r['Supplier ABN']) or '(not printed)','issue':clean(r['Issue Date']),
                'due':clean(r['Due']),'ex_gst':x}
        if (r['Type'] or '')=='Charges':
            a=(r['Amount(GST incl)'] or '').replace(',','').strip()
            try: amt=float(a)
            except: amt=None
            detail[(stmt,nmi)].append({'start':clean(r['Start Date']),'end':clean(r['End Date']),
                'ctype':clean(r['Charge Type']),'cdesc':clean(r['Charge Description']),
                'meter':clean(r['Meter Number']),'read':clean(r['Read Type']),
                'vol':clean(r['Volume']),'uom':clean(r['Units of Measure']),
                'rate':clean(r['Charge/Rate']),'rate_uom':clean(r['Charge/Rate UOM']),'amt':amt})

# ---- register ----
wb=CalamineWorkbook.from_path('zipx/PS_WP_Transaction_Register_3FY_v41_HANDOVER.xlsx')
sh=wb.get_sheet_by_name('Register').to_python(); data=sh[4:21473]
nmipat=re.compile(r'-((?:QB|31|30)\w{9,11})-')
grp=collections.defaultdict(lambda:{'amt':0.0,'rows':[],'pk':set(),'fy':set(),'periods':[]})
elec_rows=[]
for i,r in enumerate(data):
    if str(r[13])!='73312' or str(r[7]) not in FILES: continue
    rn=i+5; narr=str(r[21])
    m=nmipat.search(narr); nmi=m.group(1) if m else str(r[40]).strip()
    k=(str(r[7]),nmi)
    g=grp[k]; g['amt']+=float(r[19] or 0); g['rows'].append(rn); g['pk'].add(str(r[16])); g['fy'].add(str(r[4]))
    pm=re.search(r'Electricity Charges (\S+) to (\S+)',narr)
    if pm: g['periods'].append(pm.group(1)+' to '+pm.group(2))
    elec_rows.append({'row':rn,'stmt':str(r[7]),'nmi':nmi,'date':str(r[3])[:10],'pk':str(r[16]),
        'amt':round(float(r[19] or 0),2),'narr':narr,'site':str(r[41]),'tier':str(r[28]),'status':str(r[32])})

recon=[]
for k,g in grp.items():
    s=site.get(k)
    recon.append({'stmt':k[0],'nmi':k[1],'addr':(s or {}).get('addr','(site not on statement)'),
        'sub':(s or {}).get('sub',''),'inv':(s or {}).get('inv',''),'issue':(s or {}).get('issue',''),
        'stmt_ex_gst':round((s or {}).get('ex_gst',0.0),2),'reg_ex_gst':round(g['amt'],2),
        'lines':len(g['rows']),'rows':','.join(str(x) for x in sorted(g['rows'])),
        'pk':'/'.join(sorted(g['pk'])),'fy':'/'.join(sorted(g['fy'])),
        'periods':' ; '.join(sorted(set(g['periods'])))[:180],
        'charge_lines':len(detail.get(k,[]))})
recon.sort(key=lambda x:(-x['reg_ex_gst']))

# ---- waste series ----
waste=[]
for i,r in enumerate(data):
    na=str(r[13]); narr=str(r[21]); amt=float(r[19] or 0)
    if na in ('73122','7B115') and 'Accrue' not in narr and 'Accrual' not in narr and abs(amt)>0.005:
        waste.append({'row':i+5,'fy':str(r[4]),'acct':na,'acctname':str(r[14]),'date':str(r[3])[:10],
            'ref':str(r[7]),'pk':str(r[16]),'amt':round(amt,2),'narr':narr,'status':str(r[32]),'tier':str(r[28])})
waste.sort(key=lambda x:x['date'])

json.dump({'site':{f'{k[0]}|{k[1]}':v for k,v in site.items()},'recon':recon,'elec_rows':elec_rows,'waste':waste},
          open('b39_data.json','w'),indent=0)
print('sites',len(site),'recon groups',len(recon),'elec rows',len(elec_rows),'waste rows',len(waste))
print('recon ties:',sum(1 for x in recon if abs(x['stmt_ex_gst']-x['reg_ex_gst'])<=0.02),'/',len(recon))
print('register elec on 5 stmts:',round(sum(x['reg_ex_gst'] for x in recon),2))
