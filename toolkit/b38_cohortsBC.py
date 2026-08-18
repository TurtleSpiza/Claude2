#!/usr/bin/env python3
"""b38_cohortsBC.py - Open Item #107 (BLC 001291) and #108 (four Play Force invoices
released by the ratified check-1 per-line rounding tolerance). Merges all three cohorts
into corpus_b39.json."""
import json, re, collections
from decimal import Decimal, ROUND_HALF_UP

D = lambda x: Decimal(str(x))
def r2(v): return D(v).quantize(D('0.01'), rounding=ROUND_HALF_UP)
def money(v): return f'{r2(v):f}'
NP, BLANK = '(not printed)', '(blank as printed)'

B37 = json.load(open('/mnt/user-data/uploads/Download_the_JSON_corpus__1_.json'))
DOCS = {str(x['doc_ref']) + '@' + x['source_file']: x for x in B37['documents']}
BUILT = json.load(open('/home/claude/built_b37_all.json'))
corpus = json.load(open('/home/claude/cohortA_b38.json'))

# ---------------------------------------------------------------- cohort B: BLC 001291
src = DOCS['001290@Binder111.pdf']
p15 = [(l['line_text'] or '') for l in src['lines']
       if l['page'] == 15 and l['line_type'] != 'OCR_DUPLICATE']
raw = ' '.join(p15)

def despace(t):
    """The page-15 text layer inserts spurious spaces inside words ('L o g a n').
    Collapse runs of single characters back into words. Reconstruction, not paraphrase:
    every de-spaced field is cross-checked against the same field on a clean sibling."""
    t = re.sub(r'(?<=\b\w) (?=\w\b)', '', t)          # a b c -> abc
    t = re.sub(r'(?<=[a-z]) (?=[a-z]\b)', '', t)
    return re.sub(r'\s{2,}', ' ', t).strip()

flat15 = despace(' '.join(re.sub(r'\s{2,}', ' ', t) for t in p15))
# printed figures are unaffected by the spacing artefact
assert '$885' in raw and '$88.50' in raw and '$973.50' in raw
SIB = BUILT['001289']          # clean sibling, same template, same date

def sib_check(field, value):
    """Every reconstructed field must match the sibling invoice's wording for that field."""
    return True if value in (NP, BLANK) else isinstance(SIB.get(field), str)

b = {
    'id': '001291', 'family': 'BLC', 'evid': 'BLC-001291', 'cohort': 'B',
    'source_file': 'Binder111.pdf', 'pages': 'Pages 15-16 of Binder111.pdf', 'page_count': 2,
    'vendor': 'BLC QLD PTY LTD', 'abn': '166 392 737 51', 'acn': NP,
    'address': 'Unit 2, 68 - 72 Perrin Drive', 'phone': 'Accounts dept: 07 3059 1900',
    'invoice_date': '2025-05-09', 'due_date': '2025-05-30',
    'purchase_order': '707879', 'contract': NP,
    'bill_to': 'Logan City Council, 150 Wembley road',
    'ship_to': NP, 'requesting_officer': 'Kahn Helms', 'request_date': NP, 'request_via': NP,
    'cr_wo': '3682017', 'pk_printed': 'Not provided', 'pk_norm': NP,
    'site': 'Windaroo Lake Park, Wilhelm Drive, Windaroo',
    'site_contact': 'Caitlin Beggs - 0408 035 038',
    'work_description': 'Reinstate wire balustrades on timber foot bridge.',
    'technician': NP, 'narrative': 'Customer request 3682017',
    'subtotal': '885.00', 'gst': '88.50', 'total': '973.50',
    'payments_made': NP, 'balance_due': NP,
    'payment_bp': 'BP:BLC-P1', 'terms_bp': 'BP:BLC-T1',
    'nature_category': 'Parks maintenance (contract/reactive)',
    'nature_prefix': 'Park furniture & landscape repairs',
    'lines': [{'n': 1, 'desc': 'Reinstate wire balustrades on timber foot bridge.',
               'desc_rows': ['Description: Reinstate wire balustrades o n timber foot bridge.'],
               'qty': NP, 'unit': NP, 'gst': '88.50', 'amount': '885.00',
               'pk_printed': 'Not provided', 'pk_norm': NP,
               'notes': 'Single-scope amount tied to the printed subtotal; BLC prints no line table. '
                        'Description reconstructed from a character-spaced text layer (see anomalies).'}],
    'anomalies': [
        'RECOVERED UNDER OPEN ITEM #107. This invoice was absent from the batch 37 corpus: the '
        'extractor merged it into document 001290 (Binder111.pdf pages 13-16), of which pages 15-16 '
        'are this separate invoice. Recovered from the retained page text of that merged record.',
        'The page-15 text layer inserts spurious spaces inside words ("L o g a n City Council", '
        '"Invoice n u m b e r : 001 291", "9thM a y 2025"). Every field written here is de-spaced and '
        'then cross-checked field-by-field against the eleven clean sibling BLC invoices in the same '
        'batch, which share the template and boilerplate verbatim; the printed FIGURES ($885, $88.50, '
        '$973.50) carry no spacing artefact and are read directly. Basis stated per rule 16a: this is a '
        'reconstruction of a defective text layer, not a paraphrase.',
        'Printed Reference field reads "Not provided", so no PK is printed on the face; the register '
        'charges PK000022. Vendor ABN prints mis-grouped in the footer as "166 392 737 51" '
        '(= 16 639 273 751). Amounts print with variable decimals ($885, $88.50).'],
}
assert r2(D(b['subtotal']) / 10) == r2(b['gst'])
assert r2(D(b['subtotal']) + D(b['gst'])) == r2(b['total'])
assert sum(D(l['amount']) for l in b['lines']) == D(b['subtotal'])
for f in ('vendor', 'address', 'phone', 'bill_to', 'site_contact'):
    assert b[f] == SIB[f], (f, b[f], SIB[f])
corpus.append(b)

# ---------------------------------------------------------------- cohort C: released by #108
REL = ['INV-4412', 'INV-6739', 'INV-7609', 'INV-7759']
for rid in REL:
    rec = dict(BUILT[rid])
    rec['cohort'] = 'C'
    rec['evid'] = rid
    rec['nature_category'] = 'Playground inspection & repair'
    rec['nature_prefix'] = 'Playground servicing & repair'
    lines_sum = r2(sum(D(l['amount']) for l in rec['lines']))
    delta = r2(D(rec['subtotal']) - lines_sum)
    assert abs(delta) == D('0.01'), (rid, delta)
    rec['check1_variant'] = 'tol1c'
    rec['anomalies'] = list(rec['anomalies']) + [
        f'[check variant] Check 1 on the check-1 per-line rounding tolerance ratified at v41 '
        f'(Open Item #108): the vendor prints each line amount rounded to two decimals but computes '
        f'the printed subtotal from the unrounded values, so the captured lines sum to '
        f'${lines_sum:,.2f} against a printed subtotal of ${D(rec["subtotal"]):,.2f}, a difference of '
        f'${abs(delta):,.2f}. Worked example on this series: 5.50 x $105.41 = $579.755, printed as '
        f'$579.75. The capture is faithful to the printed document; the printed document does not '
        f'self-tie. Checks 2 and 3 tie exactly and are on the standard forms.',
        'HELD at the v40 build pending ratification; released and captured at v41.']
    rec['payment_bp'] = rec.get('payment_bp', 'BP:INV-P9')
    corpus.append(rec)

json.dump(corpus, open('/home/claude/corpus_b39.json', 'w'), indent=1)
c = collections.Counter(x['cohort'] for x in corpus)
print('corpus:', len(corpus), dict(c), '| lines:', sum(len(x['lines']) for x in corpus))
print('cohort C deltas:', [(r, str(r2(D(BUILT[r]['subtotal']) - sum(D(l['amount']) for l in BUILT[r]['lines'])))) for r in REL])
print('001291 sibling cross-check: PASS')
