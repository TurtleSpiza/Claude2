import json, re, hashlib, collections
from decimal import Decimal, ROUND_HALF_UP

D = json.load(open('/mnt/user-data/uploads/Download_the_JSON_corpus__1_.json'))
docs = D['documents']
def r2(x): return float(Decimal(str(x)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))

defects = []
def add(fam, ref, msg): defects.append((fam, ref, msg))

# ---- ABN checksum (ATO weighted modulus 89) ----
W = [10,1,3,5,7,9,11,13,15,17,19]
def abn_ok(a):
    dd = re.sub(r'\D','', a or '')
    if len(dd) != 11: return None
    n = [int(c) for c in dd]; n[0] -= 1
    return sum(w*v for w,v in zip(W,n)) % 89 == 0
LCC = '21627796435'

# ---- per-document gates ----
for x in docs:
    ref = x['doc_ref']
    priced = [l for l in x['lines'] if l.get('line_ex_gst') is not None]
    s = r2(sum(l['line_ex_gst'] for l in priced)) if priced else 0.0
    sub, gst, tot = x['printed_subtotal_ex_gst'], x['printed_gst'], x['printed_total_incl_gst']

    # G1 qty x unit = amount
    for l in priced:
        q, u, a = l.get('qty'), l.get('unit_price_ex_gst'), l['line_ex_gst']
        if q is not None and u is not None:
            if abs(r2(q*u) - r2(a)) > 0.005:
                add('G1', ref, f"qty {q} x unit {u} = {r2(q*u)} but line_ex_gst {a}")

    # G2 captured = printed subtotal
    if sub is None: add('F1', ref, 'printed_subtotal_ex_gst is null')
    elif abs(s - r2(sub)) > 0.005:
        add('G2', ref, f"captured {s} vs printed subtotal {sub} (diff {r2(s-sub)}), priced lines {len(priced)}")

    # G3 printed GST = 10% of subtotal
    if sub is not None and gst is not None and abs(r2(sub/10) - r2(gst)) > 0.005:
        add('G3', ref, f"printed GST {gst} vs 10% of {sub} = {r2(sub/10)}")
    if gst is None: add('F1', ref, 'printed_gst is null')

    # G4 subtotal + GST = total
    if None not in (sub, gst, tot) and abs(r2(sub+gst) - r2(tot)) > 0.005:
        add('G4', ref, f"subtotal {sub} + GST {gst} = {r2(sub+gst)} vs printed total {tot}")
    if tot is None: add('F1', ref, 'printed_total_incl_gst is null')

    # G5 ABN
    a = x.get('supplier_abn')
    if a is None: add('G5', ref, f"supplier_abn null ({x['supplier']})")
    else:
        dd = re.sub(r'\D','', a)
        if dd == LCC: add('F8', ref, f"supplier_abn is LCC's own ABN ({x['supplier']})")
        elif abn_ok(a) is False: add('G5', ref, f"ABN {a} fails checksum")

    if x['supplier'] in (None,'','Unclassified supplier'):
        add('G6', ref, 'supplier unclassified')

# ---- duplicate detection (report claims 0) ----
byref = collections.defaultdict(list)
for x in docs: byref[x['doc_ref']].append(x)
for k, v in byref.items():
    if len(v) > 1:
        add('R12', k, 'doc_ref appears %d times: %s' % (len(v), '; '.join(
            f"{y['source_file']} p{y['page_range']} ${y['captured_ex_gst']}" for y in v)))

def norm(x):
    t = ' '.join((l['line_text'] or '') for l in x['lines'] if l['line_type'] != 'OCR_DUPLICATE')
    return hashlib.md5(re.sub(r'\s+',' ',t).strip().encode()).hexdigest()
byh = collections.defaultdict(list)
for x in docs: byh[norm(x)].append(x['doc_ref']+'@'+x['source_file'])
for h, v in byh.items():
    if len(v) > 1: add('R12', v[0], 'identical normalised text: ' + ', '.join(v))

# ---- segmentation: invoice-number tokens in page text vs doc_refs ----
refs = set(x['doc_ref'] for x in docs)
pat = re.compile(r'(?:Invoice\s*(?:number|No\.?|Number)\s*:?\s*)([0-9]{3}\s?[0-9]{3}|INV-[0-9]{4}|[0-9]{6,7})', re.I)
for x in docs:
    found = set()
    for l in x['lines']:
        if l['line_type'] == 'OCR_DUPLICATE': continue
        for m in pat.finditer(l['line_text'] or ''):
            found.add(re.sub(r'\s','', m.group(1)))
    own = re.sub(r'\s','', x['doc_ref'])
    extra = {f for f in found if f != own and f.lstrip('0') != own.lstrip('0')}
    if extra: add('SEG', x['doc_ref'], f"page text carries other invoice numbers {sorted(extra)} over pages {x['page_range']}")

# ---- report ----
fam = collections.Counter(f for f,_,_ in defects)
print('DOCUMENTS', len(docs), '| PRICED LINES', sum(1 for x in docs for l in x['lines'] if l.get('line_ex_gst') is not None))
print('DEFECTS BY FAMILY:', dict(fam))
print()
for f in sorted(fam):
    print('---', f)
    for ff, ref, msg in defects:
        if ff == f: print(f'   {ref:<10} {msg}')
json.dump([{'family':f,'doc':r,'detail':m} for f,r,m in defects], open('/home/claude/defects_b37.json','w'), indent=1)
