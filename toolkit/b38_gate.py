#!/usr/bin/env python3
"""b38_gate.py - repair and gate the levi_bind_3 corpus (Batch 38 candidate)."""
import json, re, collections, hashlib
from decimal import Decimal, ROUND_HALF_UP

D = lambda x: Decimal(str(x))
def r2(v): return D(v).quantize(D('0.01'), rounding=ROUND_HALF_UP)

SRC = '/mnt/user-data/uploads/Download_the_JSON_corpus__2_.json'
C = json.load(open(SRC))
docs = C['documents']
log, flags = [], []
def rec(f, d, m): log.append({'family': f, 'doc': d, 'action': m})

W = [10,1,3,5,7,9,11,13,15,17,19]
def abn_ok(a):
    dd = re.sub(r'\D', '', a or '')
    if len(dd) != 11: return False
    n = [int(c) for c in dd]; n[0] -= 1
    return sum(w*v for w, v in zip(W, n)) % 89 == 0

def vis(x):
    return [l for l in x['lines'] if l['line_type'] != 'OCR_DUPLICATE']
def flat(x):
    return re.sub(r'\s+', ' ', ' '.join((l['line_text'] or '') for l in vis(x)))

AMT = re.compile(r'^\s*(.+?)\s{2,}([\d,]+\.\d{2})\s*$')

# ---- F9: amount-column-only line tables classified NARRATIVE (v40 family) ----
for x in docs:
    sub = x['printed_subtotal_ex_gst']
    priced = [l for l in x['lines'] if l.get('line_ex_gst') is not None]
    if sub is None or priced:
        continue
    cands = []
    for l in x['lines']:
        if l['line_type'] != 'NARRATIVE':
            continue
        t = (l['line_text'] or '').rstrip()
        m = AMT.match(t)
        if not m:
            continue
        desc = m.group(1).strip()
        if re.match(r'^(Subtotal|GST|Total|Balance|Paid|Bank|BSB|Account|Name|Page|DATE|DUE)', desc, re.I):
            continue
        cands.append((l, desc, D(m.group(2).replace(',', ''))))
    if cands and r2(sum(c[2] for c in cands)) == r2(sub):
        for l, desc, amt in cands:
            l['line_type'] = 'PRICED'; l['line_ex_gst'] = float(amt)
            l['qty'] = None; l['unit_price_ex_gst'] = None; l['gst'] = None
            l['note'] = ('Amount column only; no qty or unit price printed. '
                         'Reclassified NARRATIVE -> PRICED on the printed amount (F9).')
        x['captured_ex_gst'] = float(r2(sub)); x['self_tie'] = 'TIE'
        rec('F9', x['doc_ref'], f'{len(cands)} priced line(s) recovered, sum ties printed subtotal {sub}')

# ---- identity / scope screening ----
for x in docs:
    a = x.get('supplier_abn')
    if not abn_ok(a):
        flags.append(('ABN', x['doc_ref'], f'ABN {a} fails checksum'))
    txt = flat(x)
    m = re.search(r'Invoice\s*#?\s*:?\s*(\d{7,8})', txt)
    if str(x['doc_ref']).startswith('UNCLASSIFIED') and m:
        x['printed_invoice_no'] = m.group(1)
        rec('ID', x['doc_ref'], f'printed invoice number {m.group(1)} recovered from the OCR layer')

# ---- duplicates within the corpus ----
byref = collections.Counter(str(x['doc_ref']) for x in docs)
for k, v in byref.items():
    if v > 1: flags.append(('R12', k, f'doc_ref appears {v} times'))
h = collections.defaultdict(list)
for x in docs:
    h[hashlib.md5(flat(x).encode()).hexdigest()].append(x['doc_ref'])
for k, v in h.items():
    if len(v) > 1: flags.append(('R12', v[0], 'identical normalised text: ' + ', '.join(map(str, v))))

# ---- arithmetic gates ----
tie = out = 0
for x in docs:
    priced = [l for l in x['lines'] if l.get('line_ex_gst') is not None]
    s = r2(sum(D(l['line_ex_gst']) for l in priced)) if priced else D('0.00')
    sub, gst, tot = x['printed_subtotal_ex_gst'], x['printed_gst'], x['printed_total_incl_gst']
    if sub is not None and s == r2(sub): tie += 1
    else:
        out += 1
        flags.append(('TIE', x['doc_ref'], f'captured {s} vs printed subtotal {sub}'))
    if None not in (sub, gst) and abs(r2(D(sub) / 10) - r2(gst)) > D('0.01'):
        flags.append(('GST', x['doc_ref'], f'printed GST {gst} vs 10% {r2(D(sub)/10)}'))
    if None not in (sub, gst, tot) and abs(r2(D(sub) + D(gst)) - r2(tot)) > D('0.01'):
        flags.append(('TOT', x['doc_ref'], f'{sub}+{gst} != {tot}'))

C['manifest']['repair_pass'] = 'b38_gate.py, 18-Aug-2026'
json.dump(C, open('/mnt/user-data/outputs/corpus_b38_repaired.json', 'w'), indent=1)
json.dump({'repairs': log, 'flags': [{'family': f, 'doc': d, 'detail': m} for f, d, m in flags]},
          open('/mnt/user-data/outputs/gate_log_b38.json', 'w'), indent=1)

print('docs', len(docs), '| TIE', tie, '| OUT', out)
print('repairs:', dict(collections.Counter(r['family'] for r in log)))
print('flags:', dict(collections.Counter(f for f, _, _ in flags)))
for f in flags: print('   ', f)
