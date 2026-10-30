"""Read and protect the approved fixed-position residential form during edits.

Bookmarks bind editable fields to paragraphs and rent-table cells.
The patch validates content first; the central renderer then restores layout.
"""
import re
import unicodedata
from lxml import etree
from notice_word_rows import inspect_rows

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WIDTHS = {'TENANT_NAMES': 480, 'PROPERTY_STREET': 445, 'PROPERTY_CITY': 445,
          'COUNTY': 445, 'PAYMENT_RECIPIENT': 420, 'PAYMENT_STREET': 420,
          'PAYMENT_CITY': 420, 'PAYMENT_PHONE': 415,
          'SIGNER_NAME_AND_TITLE': 271, 'PROVIDER_STREET': 375,
          'PROVIDER_CITY': 375, 'PROVIDER_PHONE': 375}


def text(el):
    return ''.join(t.text or '' for t in el.iter(W+'t'))


def bindings(root):
    result = {}
    for i, p in enumerate(root.iter(W+'p')):
        for b in p.iter(W+'bookmarkStart'):
            name = b.get(W+'name', '')
            if name.startswith('BE_'):
                if name in result:
                    raise ValueError('Repeated form field')
                result[name] = (i, p)
    return result


def is_standard(root):
    return any(b.get(W+'name') == 'BE_RENT_ROW_1' for b in root.iter(W+'bookmarkStart'))


def cents(value):
    if not re.fullmatch(r'\$(?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2}', value):
        raise ValueError('Invalid fixed-form rent amount')
    a, b = value.replace('$', '').replace(',', '').split('.')
    n = int(a)*100+int(b)
    if not 0 < n <= 100000000000:
        raise ValueError('Invalid fixed-form rent amount')
    return n


def total_cell(root):
    candidates = []
    ps = list(root.iter(W+'p'))
    for tbl in root.iter(W+'tbl'):
        rows = tbl.findall(W+'tr')
        if len(rows) != 1:
            continue
        cells = rows[0].findall(W+'tc')
        if len(cells) == 2 and text(cells[0]).strip() == 'TOTAL AMOUNT DUE:':
            pp = list(cells[1].iter(W+'p'))
            if len(pp) != 1:
                raise ValueError('Invalid total cell')
            candidates.append({'paragraph_index': ps.index(pp[0]),
                               'text': text(pp[0]), 'cents': cents(text(pp[0]))})
    if len(candidates) != 1:
        raise ValueError('Missing or repeated fixed-form total')
    return candidates[0]


def inspect_standard(root):
    b = bindings(root)
    rows, blank_seen = [], False
    for row in inspect_rows(root):
        i, p = row['paragraph_index'], row['p']
        period, amount = row['period'], row['amount']
        if not period and not amount:
            blank_seen = True
            continue
        if blank_seen or not period or not amount:
            raise ValueError('Incomplete fixed-form rent row')
        rows.append({'label': period, 'text': amount, 'cents': cents(amount),
                     'paragraph_index': i})
    total = total_cell(root)
    if not rows or sum(r['cents'] for r in rows) != total['cents']:
        raise ValueError('Fixed-form rent total does not reconcile')
    fields = [{'name': k[3:], 'paragraph_index': i, 'text': text(p)}
              for k, (i, p) in b.items() if k[3:] in WIDTHS]
    return {'rent_tables': [{'rows': rows, 'total': total}], 'form_fields': fields}


def width(s):
    result = 0
    for c in s:
        em = (1 if unicodedata.east_asian_width(c) in {'W', 'F'} else
              .95 if c in 'MW@%' else .28 if c in "ilI.,:;'!| " else
              .73 if c.isupper() else .5 if c.isdigit() else .56)
        result += em*12
    return result


def lines(p):
    out = ['']
    for n in p.iter():
        if n.tag == W+'t':
            out[-1] += n.text or ''
        elif n.tag == W+'br':
            out.append('')
    return out


def printed_signer(root):
    _, p = bindings(root)['BE_SIGNER_NAME_AND_TITLE']
    value = ' '.join(s.strip() for s in lines(p) if s.strip())
    name, comma, title = value.partition(',')
    if not comma or not name.strip() or not title.strip():
        raise ValueError('Printed signer name and title require review')
    return {'name': name.strip(), 'title': title.strip()}


def validate_edit(before, after):
    if not is_standard(before):
        return
    fields = bindings(before)
    allowed = {i for k, (i, p) in fields.items() if k[3:] in WIDTHS or k.startswith('BE_RENT_ROW_')}
    allowed.update(part['period_paragraph_index'] for row in inspect_rows(before) for part in row['parts'])
    allowed.update(part['paragraph_index'] for row in inspect_rows(before) for part in row['parts'])
    allowed.add(total_cell(before)['paragraph_index'])
    a, b = list(before.iter(W+'p')), list(after.iter(W+'p'))
    if len(a) != len(b):
        raise ValueError('Approved form layout changed')
    for i, (old, new) in enumerate(zip(a, b)):
        if i not in allowed and etree.tostring(old) != etree.tostring(new):
            raise ValueError('Approved form boilerplate must not change')
    # Reconcile edited rows and total before returning any file.
    inspect_standard(after)
    printed_signer(after)
    for name, (i, p) in bindings(after).items():
        key = name[3:]
        if key in WIDTHS:
            values = lines(p)
            if len(values) > (2 if key in {'TENANT_NAMES', 'SIGNER_NAME_AND_TITLE'} else 1):
                raise ValueError('Approved field line count changed')
            if any(width(v) > WIDTHS[key]-4 or '{{' in v or '}}' in v for v in values):
                raise ValueError('Edited field exceeds approved space')
        elif name.startswith('BE_RENT_ROW_'):
            # The PDF paginator wraps period text in its own column. Inspect
            # above verifies the independent amount cell and reconciled total.
            if len(text(p)) > 10100 or '{{' in text(p) or '}}' in text(p):
                raise ValueError('Edited rent row exceeds the notice input size limit')
