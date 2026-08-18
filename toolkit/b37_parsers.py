#!/usr/bin/env python3
"""b37_parsers.py - per-template header and line parsers for Batch 37.
Head-block scanning is bounded to the rows ABOVE the first priced row so bank-account
rows can never be read as the PK 'Account:' row (v40 build trap)."""
import re
NP = '(not printed)'
BLANK = '(blank as printed)'

MON = {m: i for i, m in enumerate(
    ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'], 1)}


def isodate(s):
    if not s: return None
    s = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', str(s).strip(), flags=re.I)
    m = re.search(r'(\d{1,2})[- ]([A-Za-z]{3})[a-z]*[- ](\d{4})', s)
    if m: return f'{m.group(3)}-{MON[m.group(2).lower()]:02d}-{int(m.group(1)):02d}'
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', s)
    if m: return f'{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}'
    return None


def pk_norm(tok):
    if not tok: return NP
    m = re.match(r'^PK\s?0*(\d{1,7})$', str(tok).strip())
    return 'PK' + m.group(1).zfill(6) if m else NP


LABEL = re.compile(r'^(Contact|Acc|Account|Contract|Date Completed|P/O|PO|PK|Description|'
                   r'Qty|Invoice|Reference|Due Date|Bank|BSB|Remittance|Name|Electronic|'
                   r'Subtotal|Total|TOTAL|GST|Payment|Please|Client|Project|Address|'
                   r'Purchase|Tax|Sub total|Direct|Web|QBCC|NB)\b')
AMT = re.compile(r'\d[\d,]*\.\d{2}')


def head_rows(rows, first_priced_i):
    """Rows above the first priced row, after the column-header row."""
    start = 0
    for i, t in enumerate(rows[:first_priced_i]):
        if re.search(r'\bDescription\b', t) and re.search(r'(Qty|Quantity)', t):
            start = i + 1
    return rows[start:first_priced_i], start


def pf_weis_head(rows, first_priced_i, fam):
    """Play Force / Weis head block: contact, PK, contract, completed, site."""
    hr, _ = head_rows(rows, first_priced_i)
    out = {'contact': NP, 'pk': NP, 'contract': NP, 'completed': NP,
           'site': NP, 'po': NP, 'cr': NP, 'acc_label': NP}
    site = []
    for i, t in enumerate(hr):
        s = t.strip()
        if not s: continue
        if s.startswith('Contact:'):
            out['contact'] = s.split(':', 1)[1].strip() or BLANK
        elif re.match(r'^Acc(ount)?:', s):
            body = s.split(':', 1)[1].strip()
            out['acc_label'] = body
            m = re.search(r'PK\s?\d{5,7}', body)
            if m: out['pk'] = m.group(0)
        elif s.startswith('Contract'):
            out['contract'] = re.sub(r'^Contract\s*#?\s*:?\s*', '', s).strip()
        elif s.startswith('Date Completed:'):
            out['completed'] = s.split(':', 1)[1].strip() or BLANK
        elif re.match(r'^P/O\s*#', s):
            out['po'] = re.sub(r'^P/O\s*#\s*', '', s).strip()
        elif re.match(r'^PK\s?\d{5,7}$', s):
            out['pk'] = s
        elif re.match(r'^\d{7}\s*-\s*\S', s):
            site = [s]
            for nxt in hr[i + 1:]:
                n = nxt.strip()
                if not n or LABEL.match(n) or re.match(r'^\d{7}\s*-', n):
                    break
                site.append(n)
    if site:
        out['site'] = ' '.join(site)
        m = re.match(r'^(\d{7})', out['site'])
        if m: out['cr'] = m.group(1)
    return out


def right_col_after(rows, label, maxskip=4):
    """Xero letterheads print the label and its value on separate layout rows."""
    for i, t in enumerate(rows):
        if re.search(rf'\b{label}\b\s*$', t.strip()):
            for j in range(i + 1, min(i + 1 + maxskip, len(rows))):
                s = rows[j].strip()
                if s and not LABEL.match(s):
                    return s
    return None


RE_PFC = re.compile(r'^\s*([\d.]+)\s{2,}(\S+)\s{2,}(.+?)\s{2,}([\d,]+\.\d{2})\s{2,}([\d,]+\.\d{2})\s*$')
RE_XERO = re.compile(r'^(.+?)\s{2,}([\d.]+)\s{2,}([\d,]+\.\d{2})\s{2,}(\d+%)?\s*([\d,]+\.\d{2})?\s*$')
RE_OZI = re.compile(r'^\s*(\d+)\s{2,}(.+?)\s{2,}\$([\d,]+\.\d{2})\s{2,}\$([\d,]+\.\d{2})\s*$')


def f12_tail(rows, i):
    """v34 rule: join the immediately following narrative row when it is a wrapped tail."""
    if i + 1 >= len(rows): return None
    n = rows[i + 1].strip()
    if not n or AMT.search(n) or LABEL.match(n): return None
    if len(n.split()) > 6: return None
    return n


def parse_line(raw, rows, i, fam, corpus_line):
    """Returns (desc, desc_rows, qty, unit, gst)."""
    s = raw.rstrip()
    desc = qty = unit = gst = None
    drows = [s.strip()]
    if fam == 'PFC':
        m = RE_PFC.match(s)
        if m:
            qty, item, desc, unit = m.group(1), m.group(2), m.group(3).strip(), m.group(4).replace(',', '')
            gst = NP
    elif fam in ('PFX', 'WEI', 'FD'):
        m = RE_XERO.match(s)
        if m:
            desc, qty, unit = m.group(1).strip(), m.group(2), m.group(3).replace(',', '')
            gst = m.group(4) or NP
    elif fam == 'OZI':
        m = RE_OZI.match(s)
        if m:
            qty, desc, unit = m.group(1), m.group(2).strip(), m.group(3).replace(',', '')
            gst = NP
    if desc is None:
        # BLC single-scope, Grange amount-column, Flavell wrapped block
        cl = corpus_line.get('desc_rows')
        if cl:
            return ' '.join(cl), list(cl), NP, NP, corpus_line.get('gst') or NP
        desc = re.sub(r'\s{2,}', ' ', s).strip()
        desc = re.sub(r'\s*\$?[\d,]+\.\d{2}\s*$', '', desc).strip()
        return desc, drows, NP, NP, NP
    tail = f12_tail(rows, i)
    if tail:
        desc = f'{desc} {tail}'
        drows.append(tail)
    return desc, drows, qty, unit, gst
