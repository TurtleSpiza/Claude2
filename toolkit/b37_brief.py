#!/usr/bin/env python3
"""b37_brief.py - assigns boilerplate keys and nature, authors brief_b37.json and
corpus_b37.json in the driver format (rule 20)."""
import json, re, collections
from decimal import Decimal, ROUND_HALF_UP
from python_calamine import CalamineWorkbook

D = lambda x: Decimal(str(x))
def r2(v): return D(v).quantize(D('0.01'), rounding=ROUND_HALF_UP)
NP = '(not printed)'

B = json.load(open('/home/claude/built_b37.json'))
MATCH = {o['doc']: o for o in json.load(open('/home/claude/match_b37.json'))}
RAW = json.load(open('/mnt/user-data/outputs/corpus_b37_repaired.json'))
DOCS = {x['doc_ref']: x for x in RAW['documents'] if not x.get('duplicate_of')}

wb = CalamineWorkbook.from_path('kit/PS_WP_Transaction_Register_3FY_v39_HANDOVER.xlsx')
REG = wb.get_sheet_by_name('Register').to_python()

EVID_PREFIX = {'PFC': 'INV-', 'PFX': 'INV-', 'WEI': 'INV-', 'BLC': 'BLC-',
               'FD': 'INV-', 'GRA': 'GRA-', 'OZI': 'OZI-'}
FAMILY = {
    'PFC': dict(vendor='Play Force Australia Pty Ltd', abn='89 677 476 541', creditor='PLA073',
                label='Play Force Australia Pty Ltd', nature_category='Playground inspection & repair',
                nature_prefix='Playground servicing & repair', evid_prefix='INV-'),
    'PFX': dict(vendor='Play Force Australia Pty Ltd', abn='89 677 476 541', creditor='PLA073',
                label='Play Force Australia Pty Ltd', nature_category='Playground inspection & repair',
                nature_prefix='Playground servicing & repair', evid_prefix='INV-'),
    'WEI': dict(vendor='WEIS CONTRACTORS', abn='71 812 055 648', creditor='WEI013',
                label='Weis Contractors', nature_category='Parks maintenance (contract/reactive)',
                nature_prefix='High-pressure cleaning & site works', evid_prefix='INV-'),
    'BLC': dict(vendor='BLC QLD PTY LTD', abn='16 639 273 751', creditor='BLC001',
                label='BLC Queensland Pty Ltd', nature_category='Parks maintenance (contract/reactive)',
                nature_prefix='Park furniture & landscape repairs', evid_prefix='BLC-'),
    'FD': dict(vendor='The Trustee for Flavell-Dau Family Trust', abn='47 220 358 629', creditor='THE289',
               label='The Trustee for Flavell-Dau Family Trust (t/a F82 Landscaping)',
               nature_category='Parks maintenance (contract/reactive)',
               nature_prefix='Steel fabrication & repair', evid_prefix='INV-'),
    'GRA': dict(vendor='Grange 1951 Pty Ltd atf The Kal Ma Kuta Trust', abn='37 980 809 703', creditor='',
                label='Grange 1951 Pty Ltd', nature_category='Recovery / reimbursement',
                nature_prefix='Reimbursement of third-party fencing invoices', evid_prefix='GRA-'),
    'OZI': dict(vendor='Ozifresh of Australia Pty Ltd', abn='59 500 553 252', creditor='',
                label='Ozifresh of Australia Pty Ltd', nature_category='Cleaning & sanitary',
                nature_prefix='Sanitary disposal unit servicing', evid_prefix='OZI-'),
}

SURFACING = re.compile(r'(softfall|soft fall|rubber|sand|mulch|surfacing|wetpour|wet pour)', re.I)

# ------------------------------------------------------------------ boilerplate
NEW_BP = [
    dict(key='BP:BLC-P1', field='Payment details', vendor='BLC QLD PTY LTD', first='BLC-001270',
         text='Direct Deposit Details Account Name BLC QLD PTY LTD BSB 014315 Account no. 152738418'),
    dict(key='BP:BLC-T1', field='Terms & notices', vendor='BLC QLD PTY LTD', first='BLC-001270',
         text='Payment terms: 21 days'),
    dict(key='BP:GRA-P1', field='Payment details', vendor='Grange 1951 Pty Ltd atf The Kal Ma Kuta Trust',
         first='GRA-380CFR1', text='Payment Bank Bendigo Bank BSB 633 000 Account No. 18340 9721'),
    dict(key='BP:OZI-P1', field='Payment details', vendor='Ozifresh of Australia Pty Ltd',
         first='OZI-0012910',
         text='Payment can be made via; EFT, Cash, Cheque or Credit Card* Bank Account Details: '
              'BSB 084391 Acc # 893078787 Please quote invoice number'),
    dict(key='BP:INV-P23', field='Payment details', vendor='Play Force Australia Pty Ltd',
         first='INV-7624',
         text='Electronic Funds Transfer Account Name: Play Force Australia Pty Ltd BSB: 064 400 '
              'Account: 10367833 Remittance To: accounts@playforce.com.au'),
]


def bp_for(ref, rec):
    fam = rec['family']
    x = DOCS[ref]
    rows = [(l['line_text'] or '').strip() for l in x['lines']
            if l['line_type'] != 'OCR_DUPLICATE'
            and not str(l['line_type']).startswith(('IMAGE', 'ANNOT'))]
    j = ' '.join(rows)
    extra = []
    if fam in ('PFC',):
        pay = 'BP:INV-P23' if any(t.startswith('Account Name: Play Force') for t in rows) else 'BP:INV-P9'
        terms = 'BP:INV-T2' if 'Business Name, ABN and bank details changed' in j else ''
        if rec['page_count'] > 1 and 'Terms & Conditions' in j:
            if terms:
                extra.append('BP:INV-T3 (Play Force Terms & Conditions pages follow the invoice)')
            else:
                terms = 'BP:INV-T3'
        if 'Payment details for direct deposit' in j:
            extra.append('A SECOND payment block prints on the terms pages: '
                         '"Payment details for direct deposit: BSB: 084-123" (one key per cell, '
                         'Amendment 2), recorded here rather than in the payment field.')
    elif fam == 'PFX':
        pay, terms = 'BP:INV-P22', ''
    elif fam == 'WEI':
        pay, terms = 'BP:INV-P14', ''
    elif fam == 'BLC':
        pay, terms = 'BP:BLC-P1', 'BP:BLC-T1'
    elif fam == 'FD':
        pay, terms = 'BP:FD-P1', ''
    elif fam == 'GRA':
        pay, terms = 'BP:GRA-P1', ''
    else:
        pay, terms = 'BP:OZI-P1', ''
    return pay, terms, extra


# ------------------------------------------------------------------ build
corpus, matches, coding_notes, verdicts, followups = [], {}, {}, {}, {}
bp_bumps = collections.Counter()
new_keys = {b['key'] for b in NEW_BP}

for ref, rec in B.items():
    fam = rec['family']; F = FAMILY[fam]
    x = DOCS[ref]
    m = MATCH[ref]
    printed_id = '380CFR1' if ref == '380' else ref
    evid = F['evid_prefix'] + (printed_id if not printed_id.startswith('INV-') else printed_id[4:]) \
        if F['evid_prefix'] != 'INV-' else printed_id
    if fam in ('GRA',): evid = 'GRA-380CFR1'
    if fam == 'OZI': evid = 'OZI-0012910'
    if fam == 'BLC': evid = 'BLC-' + ref

    pay, terms, extra = bp_for(ref, rec)
    bp_bumps[pay] += 1
    if terms: bp_bumps[terms] += 1

    # nature from this invoice's own sighted lines
    val_surf = sum(D(l['amount']) for l in rec['lines'] if SURFACING.search(l['desc']))
    cat = F['nature_category']
    if fam in ('PFC', 'PFX') and val_surf > D(rec['subtotal']) / 2:
        cat = 'Playground surfacing'
    parts = [f"{l['desc']} ${r2(l['amount']):,.2f}" for l in rec['lines'] if D(l['amount']) != 0]
    detail = f"{F['nature_prefix']}: " + '; '.join(parts)
    if len(detail) > 900: detail = detail[:897] + '...'

    rec2 = dict(rec)
    rec2.update({'evid': evid, 'vendor': rec['vendor'], 'payment_bp': pay,
                 'terms_bp': terms or NP,
                 'terms_text': ('Payment terms: 21 days' if fam == 'BLC' else NP),
                 'anomalies': rec['anomalies'] + extra,
                 'nature_category': cat, 'nature_detail': detail})
    corpus.append(rec2)

    # Doc 380 posts under its printed invoice number 380CFR1, so the reference match
    # keyed on the extractor's '380' missed it. Target fixed by printed number AND amount.
    TARGET_OVERRIDE = {'380': 228}
    row = TARGET_OVERRIDE.get(ref) or (m['targets'][0] if m['verdict'] == 'exact' else None)
    if row:
        r = REG[row - 1]
        assert abs(D(r[19]) - D(rec['subtotal'])) <= D('0.01'), (ref, row)
        matches[ref] = {
            'evid': evid, 'row': row, 'linekey': str(r[0]),
            'amount': f'{r2(r[19]):f}', 'pk_charged': str(r[16]),
            'nat_acct': str(r[13]), 'acct_name': str(r[14]),
            'doc_date': str(r[3])[:10], 'status': str(r[32]),
            'creditor': str(r[10]), 'contractor': str(r[11]), 'abn': str(r[12]),
            'companions': m.get('zero_companions', []),
            'doc_type': str(r[8]),
            'nature_cat_now': cat, 'nature_detail_now': detail,
        }
        verdicts[ref] = 'Correct'
        coding_notes[ref] = (
            f"Invoice sighted and captured at line level (rule 17). "
            f"{F['nature_prefix']} at {rec['site'] if rec['site'] != NP else 'the site as printed'}; "
            f"printed subtotal ${r2(rec['subtotal']):,.2f}.")
        followups[ref] = ''
    else:
        matches[ref] = {'evid': evid, 'row': None, 'unmatched': True,
                        'reason': 'No AP creditor line in PS/WP scope; the spend posts at register '
                                  'row 256 as a PCard Expense naming the card provider.'}

brief = {
    'batch_id': 'Batch 37', 'version': 'v40', 'prev_version': 'v39', 'date': '18-Aug-2026',
    'batch_label': 'Batch 37 (binder_batch_20260818: Binder1.pdf + Binder111.pdf, 171 pages)',
    'sources': ['Binder1.pdf', 'Binder111.pdf'],
    'corpus_md5_supplied': '5b553a2e9067e3852c3db3775527ac29',
    'families': FAMILY, 'matches': matches, 'coding_notes': coding_notes,
    'verdicts': verdicts, 'followups': followups,
    'stale_note_prefixes': ['Supplier unidentified', 'Invoice may bundle'],
    'stale_followup_prefixes': ['Sight attachment', 'Supplier confirmed Tier 1'],
    'new_bp': NEW_BP, 'bp_bumps': dict(bp_bumps),
    'held_rows': [], 'held_invoices': ['INV-4412', 'INV-6739', 'INV-7609', 'INV-7759'],
    'variant_count_new': 0, 'gstinc_add': '0.00',
    'unmatched': ['0012910'],
}
json.dump(brief, open('/home/claude/brief_b37.json', 'w'), indent=1)
json.dump(corpus, open('/home/claude/corpus_b37.json', 'w'), indent=1)

print('corpus invoices:', len(corpus), '| lines:', sum(len(i['lines']) for i in corpus))
print('green-block matches:', sum(1 for v in matches.values() if v.get('row')))
print('unmatched:', [k for k, v in matches.items() if not v.get('row')])
print('new BP keys:', [b['key'] for b in NEW_BP])
print('bp bumps:', dict(bp_bumps))
print('nature categories:', dict(collections.Counter(i['nature_category'] for i in corpus)))
print('evid sample:', [i['evid'] for i in corpus[:3]], '...', [i['evid'] for i in corpus[-3:]])
