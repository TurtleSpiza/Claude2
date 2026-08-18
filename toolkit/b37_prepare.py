#!/usr/bin/env python3
"""b37_prepare.py (v2) - Batch 37 corpus adaptation. Emits built_b37.json."""
import json, re, sys
from decimal import Decimal, ROUND_HALF_UP
sys.path.insert(0, '/home/claude')
from b37_parsers import (NP, BLANK, isodate, pk_norm, pf_weis_head,
                         right_col_after, parse_line)

D = lambda x: Decimal(str(x))
def r2(v): return D(v).quantize(D('0.01'), rounding=ROUND_HALF_UP)
def money(v): return f'{r2(v):f}'

RAW = json.load(open('/mnt/user-data/outputs/corpus_b37_repaired.json'))
DOCS = [x for x in RAW['documents'] if not x.get('duplicate_of')]
HELD = {'INV-4412', 'INV-6739', 'INV-7609', 'INV-7759'}


def visible(x):
    return [l for l in x['lines']
            if l['line_type'] != 'OCR_DUPLICATE'
            and not str(l['line_type']).startswith(('IMAGE', 'ANNOT'))]


def rowtexts(x, page=None):
    return [(l['line_text'] or '') for l in visible(x) if page is None or l['page'] == page]


def flat(x): return re.sub(r'\s+', ' ', ' '.join(rowtexts(x)))


def assert_tokens(x, rec):
    f = flat(x); fd = re.sub(r'\D', '', f)
    skip = {'family', 'id', 'lines', 'anomalies', 'pages', 'page_count', 'source_file',
            'invoice_date', 'due_date', 'pk_norm', 'evid', 'payment_bp', 'terms_bp',
            'subtotal', 'gst', 'total'}
    for k, v in rec.items():
        if k in skip or not isinstance(v, str) or v in (NP, BLANK, ''):
            continue
        for piece in re.split(r'\s*\|\s*|\s*;\s*|\s*,\s*', v):
            piece = piece.strip().rstrip('.')
            if len(piece) < 4 or piece.startswith('('):
                continue
            if re.sub(r'\s+', ' ', piece) in f:
                continue
            dg = re.sub(r'\D', '', piece)
            if len(dg) >= 5 and dg in fd:
                continue
            return f'{k}: {piece!r}'
    return None


def ref_value(rows, weis=False):
    """Xero prints 'Reference' then its value some rows below, interleaved with the
    customer address column. Take the first following row that looks like a reference."""
    for i, t in enumerate(rows):
        if t.strip().endswith('Reference'):
            for nxt in rows[i + 1:i + 6]:
                n = nxt.strip()
                m = re.search(r'\b(\d{6,9})\b', n)
                if weis and re.search(r'CR\s*-\s*\d{7}', n):
                    return re.sub(r'\s{2,}', ' ', n)
                if m and not re.search(r'QLD|AUSTRALIA|Wembley|Box', n):
                    return m.group(1)
    return None


def addr_block(rows, anchors):
    """Return the printed address rows verbatim, joined, starting at an anchor row."""
    for i, t in enumerate(rows):
        if any(a in t for a in anchors):
            out = [t.strip()]
            for nxt in rows[i + 1:i + 4]:
                n = nxt.strip()
                if not n or re.match(r'^(Office|Email|Web|Phone|ABN|Invoice|Reference|Due)', n):
                    break
                out.append(n)
            return ', '.join(out)
    return NP


def billto(rows, fam):
    """Verbatim customer block. Play Force classic prints 'undefined' placeholders."""
    for i, t in enumerate(rows):
        if 'Logan City Council' in t:
            out = ['Logan City Council']
            for nxt in rows[i + 1:i + 4]:
                n = re.sub(r'\s{2,}.*$', '', nxt.strip())
                if not n or re.match(r'^(Invoice|Reference|Due|ABN|Qty|Description|AUSTRALIA)', n):
                    if n == 'AUSTRALIA':
                        out.append(n)
                    continue
                out.append(n)
            return ', '.join(dict.fromkeys(out))
    return NP


def headnarr(rows, first_p):
    """Head rows above the line table, verbatim, joined with ' | '."""
    hr = []
    seen = False
    for t in rows[:first_p]:
        s2 = t.strip()
        if re.search(r'\bDescription\b', s2) and re.search(r'(Qty|Quantity)', s2):
            seen = True; continue
        if seen and s2:
            hr.append(re.sub(r'\s{2,}', ' ', s2))
    return ' | '.join(hr) if hr else NP


def family_of(x):
    s = x['supplier']
    if 'Play Force' in s: return 'PFX' if len(visible(x)) < 120 else 'PFC'
    for k, f in (('BLC', 'BLC'), ('Weis', 'WEI'), ('Flavell', 'FD'),
                 ('Grange', 'GRA'), ('Ozifresh', 'OZI')):
        if k in s: return f
    raise AssertionError(s)


def build(x):
    fam = family_of(x)
    p0, p1 = x['page_range']
    rows = rowtexts(x, p0)
    priced = [l for l in x['lines'] if l.get('line_ex_gst') is not None]
    ptexts = {(l['line_text'] or '').rstrip() for l in priced}
    first_p = next((i for i, t in enumerate(rows) if t.rstrip() in ptexts), len(rows))
    hb = pf_weis_head(rows, first_p, fam) if fam in ('PFC', 'PFX', 'WEI') else {}

    rec = {'id': x['doc_ref'], 'family': fam, 'source_file': x['source_file'],
           'pages': f'Pages {p0}-{p1} of {x["source_file"]}', 'page_count': p1 - p0 + 1,
           'acn': NP, 'invoice_date': x['invoice_date'],
           'subtotal': money(x['printed_subtotal_ex_gst']),
           'gst': money(x['printed_gst']), 'total': money(x['printed_total_incl_gst']),
           'payments_made': NP, 'balance_due': NP, 'ship_to': NP, 'request_via': NP,
           'request_date': NP, 'technician': NP, 'site_contact': NP, 'anomalies': []}

    if fam in ('PFC', 'PFX'):
        hraw = [t.strip() for t in rows[:first_p] if t.strip()]
        rec.update({
            'vendor': 'Play Force Australia Pty Ltd', 'abn': '89 677 476 541',
            'address': ('41 Industrial Ave, LOGAN VILLAGE QLD 4207, AUSTRALIA' if fam == 'PFX'
                        else addr_block(rows, ['41 Industrial Ave'])),
            'contract': hb['contract'], 'site': hb['site'],
            'requesting_officer': hb['contact'], 'pk_printed': hb['pk'], 'cr_wo': hb['cr'],
            'bill_to': billto(rows, fam),
            'work_description': hb['site'] if hb['site'] != NP else hb['acc_label'],
            'narrative': headnarr(rows, first_p),
        })
        if fam == 'PFC':
            if 'undefined' in rec['bill_to']:
                rec['anomalies'].append(
                    'The billing block prints the literal placeholder "undefined" in place of the '
                    'customer address lines; captured verbatim as printed.')
            rec['phone'] = 'Office: (07) 3803 1788'
            j = ' '.join(rows)
            rec['purchase_order'] = (re.search(r'Reference\s+(\S+)', j).group(1)
                                     if re.search(r'Reference\s+(\S+)', j) else NP)
            rec['due_date'] = isodate(next((re.search(r'Due Date\s+(\S+)', t).group(1)
                                            for t in rows if re.search(r'Due Date\s+(\S+)', t)), None))
        else:
            rec['phone'] = NP
            rec['purchase_order'] = ref_value(rows) or NP
            rec['due_date'] = isodate(next((t for t in rows if 'Due Date:' in t), None))

    elif fam == 'WEI':
        ref = ref_value(rows, weis=True)
        rec.update({
            'vendor': 'WEIS CONTRACTORS', 'abn': '71 812 055 648',
            'address': '22 Crestview St, LOGANLEA QLD 4131, AUSTRALIA',
            'phone': NP,
            'contract': hb['contract'], 'purchase_order': hb['po'],
            'requesting_officer': NP, 'pk_printed': hb['pk'],
            'cr_wo': (re.search(r'CR\s*-\s*(\d{7})', flat(x)).group(1)
                      if re.search(r'CR\s*-\s*(\d{7})', flat(x)) else NP),
            'bill_to': billto(rows, fam),
            'due_date': isodate(next((t for t in rows if 'Due Date:' in t), None)),
        })
        zero = [l for l in priced if D(l['line_ex_gst']) == 0]
        sites = [re.sub(r'\s{2,}.*$', '', (l['line_text'] or '').strip()) for l in zero]
        works_date = next((s for s in sites if isodate(s)), NP)
        site_names = [s for s in sites if not isodate(s)]
        rec['site'] = '; '.join(site_names) if site_names else NP
        rec['work_description'] = rec['site']
        rec['narrative'] = headnarr(rows, first_p)
        rec['works_date'] = works_date

    elif fam == 'BLC':
        g = lambda p: next((re.search(p, t).group(1).strip()
                            for t in rows if re.search(p, t)), None)
        addrs = [t.strip().split(':', 1)[1].strip() for t in rows if t.strip().startswith('Address:')]
        rec.update({
            'vendor': 'BLC QLD PTY LTD', 'abn': x.get('supplier_abn_printed', '166 392 737 51'),
            'address': 'Unit 2, 68 - 72 Perrin Drive',
            'phone': 'Accounts dept: 07 3059 1900',
            'purchase_order': g(r'Purchase order number:\s*(\S+)') or NP,
            'contract': NP,
            'bill_to': 'Logan City Council, 150 Wembley road',
            'requesting_officer': g(r'Contact:\s*(.+)') or NP,
            'site': (addrs[1] if len(addrs) >= 2 else NP),
            'site_contact': 'Caitlin Beggs - 0408 035 038',
            'work_description': g(r'Description:\s*(.*\S)') or NP,
            'narrative': g(r'Project:\s*(.+)') or NP,
            'due_date': isodate(g(r'due date:\s*(.+)')),
            'pk_printed': g(r'Reference:\s*(.+)') or NP,
            'cr_wo': (re.search(r'Customer request\s*(\d+)', flat(x)).group(1)
                      if re.search(r'Customer request\s*(\d+)', flat(x)) else NP),
        })
        rec['anomalies'].append(
            'Vendor ABN prints mis-grouped in the footer as "166 392 737 51" (= 16 639 273 751, '
            'ABR legal name BLC Queensland Pty Ltd); printed form captured verbatim. '
            'QBCC LIC. No. 15228753 printed. Amounts print with variable decimals.')

    elif fam == 'FD':
        rec.update({
            'vendor': 'The Trustee for Flavell-Dau Family Trust', 'abn': '47 220 358 629',
            'address': 'PO Box 5202', 'phone': NP, 'contract': NP,
            'purchase_order': 'PO# 712540', 'bill_to': 'Logan City Council',
            'requesting_officer': 'Jess Burke', 'site': 'Centre Place',
            'work_description': 'Reinstalled repaired stainless steel handrail & replaced damaged Step tread on',
            'narrative': 'Costing includes welding, battery tools & hardware.',
            'due_date': isodate(next((t for t in rows if 'Due Date:' in t), None)),
            'pk_printed': 'PK000022', 'cr_wo': '3858690',
        })
        rec['anomalies'].append(
            'Trades as F82 Landscaping (payment block). GST prints as a rate column (10%). '
            'Site prints as "Centre Place," / "Rochedale South" over two layout rows. '
            'The single charge is described over seven dated layout rows.')

    elif fam == 'GRA':
        rec.update({
            'vendor': 'Grange 1951 Pty Ltd atf The Kal Ma Kuta Trust', 'abn': '37 980 809 703',
            'address': 'PO Box 194', 'phone': NP, 'contract': NP,
            'purchase_order': BLANK, 'bill_to': 'Logan City Council',
            'requesting_officer': BLANK, 'site': NP,
            'work_description': 'Reimbursement Basketball Court Fence',
            'narrative': 'Reimbursement Basketball Court Fence',
            'due_date': NP, 'pk_printed': NP, 'cr_wo': NP,
        })
        rec['anomalies'].append(
            'Printed invoice no. is 380CFR1 (the register Reference is 380CFR1); the extractor read "380". '
            'Order No., Rep, FOB and the customer Address/City/State/Code/Phone print blank. '
            'Invoice dated 16-Dec-2022 (FY2022/23) and posted 31-Jul-2023 inside the FY2023/24 scope. '
            'Reimbursement of third-party invoices, not a direct supply.')

    elif fam == 'OZI':
        rec.update({
            'vendor': 'Ozifresh of Australia Pty Ltd', 'abn': '59 500 553 252',
            'address': 'PO Box 2001', 'phone': 'Phone: 07 3849 2990',
            'contract': NP, 'purchase_order': NP,
            'bill_to': 'Logan City Council, 177 Chambers Flat Road',
            'requesting_officer': NP, 'site': NP,
            'work_description': 'Sanitary Disposal Units',
            'narrative': 'ALL EQUIPMENT REMAINS THE PROPERTY OF OZIFRESH',
            'due_date': NP, 'pk_printed': NP, 'cr_wo': NP,
        })
        rec['anomalies'].append(
            'No PK printed on the face. Paid by purchase card: the register carries this spend at '
            'row 256 (PCard Expense, reference TE003933, narration "SQ *OZIFRESH"), which names the '
            'card provider Corporate Travel Management as the contractor rather than the supplier. '
            'No AP creditor line exists, so the invoice is captured under rule 16 only, with no green '
            'block. NB: 2.25% surcharge on all CC payments; NOTE BANK DETAILS HAVE BEEN CHANGED.')

    out = []
    for n, l in enumerate(priced, 1):
        raw = (l['line_text'] or '').rstrip()
        i = next((k for k, t in enumerate(rows) if t.rstrip() == raw), 0)
        desc, drows, qty, unit, gst = parse_line(raw, rows, i, fam, l)
        desc = re.sub(r'^Description:\s*', '', desc).strip()
        desc = re.sub(r'\s*\$\s*$', '', desc).strip()
        notes = []
        if fam == 'PFC':
            m = re.match(r'^\s*[\d.]+\s{2,}(\S+)\s{2,}', raw)
            if m: notes.append(f'Printed item code {m.group(1)}')
        if fam == 'WEI' and D(l['line_ex_gst']) == 0:
            notes.append('Zero-amount row as printed (works date / site header row).')
        if fam == 'BLC':
            notes.append('Single-scope amount tied to the printed subtotal; BLC prints no line table.')
        if fam == 'GRA':
            notes.append('Amount column only; no qty or unit price printed. Reimbursement of a third-party invoice.')
        if fam == 'FD' and D(l['line_ex_gst']) == 0:
            notes.append('Zero-amount reference row as printed.')
        if len(drows) > 1:
            notes.append(f'Description printed over {len(drows)} layout rows (F12); joined verbatim.')
        out.append({'n': n, 'desc': desc, 'desc_rows': drows, 'qty': qty or NP,
                    'unit': unit or NP, 'gst': gst or NP,
                    'amount': money(l['line_ex_gst']),
                    'pk_printed': rec.get('pk_printed', NP),
                    'pk_norm': pk_norm(rec.get('pk_printed')),
                    'notes': '; '.join(notes)})
    rec['lines'] = out
    rec['pk_norm'] = pk_norm(rec.get('pk_printed'))
    assert r2(sum(D(l['amount']) for l in out)) == r2(rec['subtotal']), rec['id']
    for k in ('vendor','address','phone','bill_to','site','work_description','contract',
              'purchase_order','requesting_officer','narrative','pk_printed','cr_wo'):
        if rec.get(k) in (None, ''): rec[k] = NP
    return rec


built, errs, tokfail = {}, [], []
for x in DOCS:
    if x['doc_ref'] in HELD: continue
    try:
        r = build(x)
        bad = assert_tokens(x, r)
        if bad: tokfail.append((x['doc_ref'], bad))
        built[x['doc_ref']] = r
    except Exception as e:
        errs.append((x['doc_ref'], repr(e)[:160]))

print('built:', len(built), '| errors:', len(errs), '| token failures:', len(tokfail))
for e in errs: print('   ERR', e)
for t in tokfail[:20]: print('   TOKEN', t)
json.dump(built, open('/home/claude/built_b37.json', 'w'), indent=1)
