#!/usr/bin/env python3
"""b38_prepare.py - v41 build inputs, cohort A (Batch 38 levi_bind_3)."""
import json, re
from decimal import Decimal, ROUND_HALF_UP

D = lambda x: Decimal(str(x))
def r2(v): return D(v).quantize(D('0.01'), rounding=ROUND_HALF_UP)
def money(v): return f'{r2(v):f}'
NP, BLANK = '(not printed)', '(blank as printed)'

B38 = json.load(open('/mnt/user-data/outputs/corpus_b38_repaired.json'))
OOS = {'UNCLASSIFIED-PAGE-1', '00022061'}

MON = {m: i for i, m in enumerate(
    ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'], 1)}
def isodate(s):
    if not s: return None
    s = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', str(s).strip(), flags=re.I)
    m = re.search(r'(\d{1,2})[- ]([A-Za-z]{3})[a-z]*[- ](\d{2,4})', s)
    if m:
        y = m.group(3); y = ('20' + y) if len(y) == 2 else y
        return f'{y}-{MON[m.group(2).lower()]:02d}-{int(m.group(1)):02d}'
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', s)
    if m: return f'{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}'
    return None

def vis(x):
    return [l for l in x['lines']
            if l['line_type'] != 'OCR_DUPLICATE'
            and not str(l['line_type']).startswith(('IMAGE', 'ANNOT'))]
def rows_of(x): return [(l['line_text'] or '') for l in vis(x)]

def assert_tokens(x, rec):
    f = re.sub(r'\s+', ' ', ' '.join(rows_of(x))); fd = re.sub(r'\D', '', f)
    skip = {'family','id','lines','anomalies','pages','page_count','source_file','invoice_date',
            'due_date','pk_norm','evid','payment_bp','terms_bp','subtotal','gst','total','cohort',
            'nature_category','nature_prefix','nature_detail','terms_text'}
    for k, v in rec.items():
        if k in skip or not isinstance(v, str) or v in (NP, BLANK, ''): continue
        for piece in re.split(r'\s*\|\s*|\s*;\s*|\s*,\s*', v):
            piece = piece.strip().rstrip('.')
            if len(piece) < 4 or piece.startswith('('): continue
            if re.sub(r'\s+', ' ', piece) in f: continue
            dg = re.sub(r'\D', '', piece)
            if len(dg) >= 5 and dg in fd: continue
            return f'{k}: {piece!r}'
    return None


def levai(x):
    rows = rows_of(x); j = ' '.join(rows)
    g = lambda p: (re.search(p, j).group(1).strip() if re.search(p, j) else None)
    priced = [l for l in x['lines'] if l.get('line_ex_gst') is not None]
    ptx = {(l['line_text'] or '').rstrip() for l in priced}
    first_p = next((i for i, t in enumerate(rows) if t.rstrip() in ptx), len(rows))
    hstart = next((i for i, t in enumerate(rows) if 'DESCRIPTION' in t and 'AMOUNT' in t), None)
    hd = [t.strip() for t in rows[(hstart or 0) + 1:first_p] if t.strip()]
    officer = hd[0] if hd else NP
    contract = next((h for h in hd if re.match(r'^(PAR|PSR|BUS)[/A-Z0-9]', h)), NP)
    pk = next((h for h in hd if re.match(r'^PK\s?\d{5,7}$', h)), NP)
    site = next((h for h in hd if re.search(r'\(Ref:', h)), NP)
    comp = next((h for h in hd if h.startswith('Completed')), NP)
    cr = (re.match(r'^(\d{7})', site).group(1) if site != NP and re.match(r'^(\d{7})', site) else NP)
    return dict(
        vendor='T & H LEVAI PTY LTD', abn='65 100 395 480',
        address=(g(r'((?:Lot 2|41) Industrial Ave)') or '41 Industrial Ave') + ', LOGAN VILLAGE QLD 4207',
        phone='Office: (07) 3803 0032', acn=NP,
        purchase_order=g(r'YOUR REF\s+(\S+)') or NP, contract=contract,
        bill_to='Logan City Council, 150 Wembley Rd, PO Box 3226, LOGAN CENTRAL QLD 4114',
        requesting_officer=officer, site=site, pk_printed=pk, cr_wo=cr,
        work_description=site, narrative=' | '.join(hd) if hd else NP,
        due_date=isodate(g(r'DUE DATE\s+(.+?)\s*$')),
        ship_to=NP, request_date=NP, request_via=NP, technician=NP, site_contact=NP,
        payments_made=NP, balance_due=NP,
        anomalies=['Levai (Xero) layout: a DESCRIPTION / AMOUNT table with no qty and no unit-price '
                   'column, so per-line qty, unit price and GST are not printed. Head rows print the '
                   f'requesting officer, contract, PK, the "CR - Park (Ref: n)" job line and "{comp}".'],
        payment_bp='BP:LEV-P1', terms_bp=NP,
        nature_category='Parks maintenance (contract/reactive)',
        nature_prefix='Parks maintenance works')


def fourpark(x):
    j = ' '.join(rows_of(x))
    g = lambda p: (re.search(p, j).group(1).strip() if re.search(p, j) else None)
    return dict(
        vendor='4Park Pty Ltd', abn='56 657 333 296',
        address='36 Adams Drive', phone='P: 08 9472 1788', acn=NP,
        purchase_order=g(r'External Document No\.\s+(\S+)') or NP,
        contract=g(r'Our Reference\s+(\S+)') or NP,
        bill_to='RENE HARREMAN, 177 CHAMBERS FLAT ROAD, MARSDEN QLD 4132',
        requesting_officer='RENE HARREMAN', site=NP, pk_printed=NP, cr_wo=NP,
        work_description='HANDLE GRIP (GYM)', narrative='HANDLE GRIP (GYM)',
        due_date=isodate(g(r'Due Date\s+(\S+)')),
        ship_to=NP, request_date=NP, request_via=NP, technician=NP, site_contact=NP,
        payments_made=NP, balance_due=NP,
        anomalies=['4Park (Forpark) layout: the printed Amount column is GST-INCLUSIVE ($2,970.00 and '
                   '$264.00); the ex-GST values are Quantity x Unit Price (Excl. GST), which is what the '
                   'printed SubTotal sums. Account No. C09284. No PK printed on the face. The extractor '
                   'lumped both printed lines into one $2,940.00 record and both were restored.'],
        payment_bp='BP:4PK-P1', terms_bp='BP:4PK-T1',
        nature_category='Minor equipment', nature_prefix='Playground equipment supply')


def tradetools(x):
    j = ' '.join(rows_of(x))
    g = lambda p: (re.search(p, j).group(1).strip() if re.search(p, j) else None)
    return dict(
        vendor='TradeTools Pty Ltd', abn='67 010 700 053',
        address='109 Park Road', phone='Phone: 07 3299 2150', acn=NP,
        purchase_order='706980', contract=NP,
        bill_to='LOGAN CITY COUNCIL, PO BOX 3226, LOGAN CITY DC',
        requesting_officer='Dominic', site=NP, pk_printed=NP, cr_wo=NP,
        work_description='Milwaukee M18 5.0Ah FUEL 6 Pce Combo Kit W/Bonus Rotary +',
        narrative='Our Order No 16854892', due_date=NP,
        ship_to='LOGAN CITY COUNCIL, PO BOX 3226, LOGAN CITY DC',
        request_date=NP, request_via=NP, technician=NP, site_contact='0433950023',
        payments_made=NP, balance_due=NP,
        anomalies=['TradeTools layout: Item Code / Item Description / Ordered / Shipped / B/Ord / UOM / '
                   'Item Price Ex GST / Line Total Ex GST / Line Total Inc GST. The bill-to block prints '
                   "LCC's own ABN 21627796435 (rule 19.2 F8 trap; the vendor ABN is in the letterhead). "
                   'No PK printed on the face. The extractor attached the line amount to a footer '
                   'warranty clause row; restored to the printed item row.'],
        payment_bp=NP, terms_bp=NP,
        nature_category='Minor equipment', nature_prefix='Power tool purchase')


BUILDERS = {'T & H Levai Pty Ltd': ('LEV', levai), '4Park Pty Ltd': ('4PK', fourpark),
            'TradeTools Pty Ltd': ('TT', tradetools)}

corpus, errs = [], []
for x in B38['documents']:
    ref = str(x['doc_ref'])
    if ref in OOS: continue
    fam, fn = BUILDERS[x['supplier']]
    rec = fn(x)
    priced = [l for l in x['lines'] if l.get('line_ex_gst') is not None]
    lines = []
    for n, l in enumerate(priced, 1):
        raw = (l['line_text'] or '').rstrip()
        drows = l.get('desc_rows') or [raw.strip()]
        if fam == 'LEV':
            desc = re.sub(r'\s{2,}[\d,]+\.\d{2}\s*$', '', raw).strip()
            drows = [desc]
        else:
            desc = re.sub(r'^\S+\s{2,}', '', drows[0]).strip()
            desc = re.sub(r'\s{2,}.*$', '', desc).strip()
            if len(drows) > 1:
                desc = (desc + ' ' + drows[1].strip()).strip()
        lines.append({'n': n, 'desc': desc, 'desc_rows': drows,
                      'qty': (f"{l['qty']:.2f}" if l.get('qty') else NP),
                      'unit': (money(l['unit_price_ex_gst']) if l.get('unit_price_ex_gst') else NP),
                      'gst': (money(l['gst']) if l.get('gst') else NP),
                      'amount': money(l['line_ex_gst']),
                      'pk_printed': rec['pk_printed'],
                      'pk_norm': rec['pk_printed'] if re.match(r'^PK\d{6}$', str(rec['pk_printed'])) else NP,
                      'notes': l.get('note') or ''})
    rec.update({'id': ref, 'family': fam,
                'evid': ('INV-' + ref if fam == 'TT' else ref),
                'source_file': 'levi bind 3.pdf', 'cohort': 'A',
                'pages': f'Page {x["page_range"][0]} of levi bind 3.pdf', 'page_count': 1,
                'invoice_date': isodate(x['invoice_date']),
                'subtotal': money(x['printed_subtotal_ex_gst']),
                'gst': money(x['printed_gst']),
                'total': money(x['printed_total_incl_gst']), 'lines': lines})
    rec['pk_norm'] = rec['pk_printed'] if re.match(r'^PK\d{6}$', str(rec['pk_printed'])) else NP
    assert r2(sum(D(l['amount']) for l in lines)) == r2(rec['subtotal']), ref
    bad = assert_tokens(x, rec)
    if bad: errs.append((ref, bad))
    corpus.append(rec)

print('cohort A:', len(corpus), '| lines:', sum(len(c['lines']) for c in corpus),
      '| token failures:', len(errs))
for e in errs[:12]: print('   TOKEN', e)
json.dump(corpus, open('/home/claude/cohortA_b38.json', 'w'), indent=1)
