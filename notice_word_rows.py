"""Read legacy tab rows and paginated two-column Word rent tables.

No network or mutation. Both the generator and revision service ship this same
small validator so saved Word remains the source of every reviewed PDF.
"""
import re

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
HEADINGS = ('TIME PERIOD DUE', 'AMOUNT DUE')


def text(p):
    return ''.join((n.text or '') if n.tag == W+'t' else ' ' if n.tag == W+'br' else '' for n in p.iter())


def bookmarks(p):
    return [x.get(W+'name', '') for x in p.iter(W+'bookmarkStart')]


def rent_table(table):
    rows = table.findall(W+'tr')
    if not rows:
        return False
    cells = rows[0].findall(W+'tc')
    return len(cells) == 2 and tuple(text(c).strip() for c in cells) == HEADINGS


def inspect_rows(root):
    ps = list(root.iter(W+'p'))
    found, mode = [], None
    for p in ps:
        marks = [n for n in bookmarks(p) if n.startswith(('BE_RENT_ROW_', 'BE_RENT_PART_'))]
        if not marks:
            continue
        if len(marks) != 1:
            raise ValueError('Repeated rent row field')
        name = marks[0]
        main = re.fullmatch(r'BE_RENT_ROW_([1-9][0-9]*)', name)
        part = re.fullmatch(r'BE_RENT_PART_([1-9][0-9]*)_([2-9]|[1-9][0-9]+)(_NOSPACE)?', name)
        if not main and not part:
            raise ValueError('Invalid rent row field')
        tc = p.getparent()
        if tc.tag == W+'tc':
            tr = tc.getparent(); table = tr.getparent(); cells = tr.findall(W+'tc')
            if not rent_table(table) or len(cells) != 2 or cells[1] != tc:
                raise ValueError('Rent period and amount must remain separate table columns')
            if any(len(list(c.iter(W+'p'))) != 1 for c in cells):
                raise ValueError('Rent cell structure changed')
            period_p = cells[0].find(W+'p')
            period, amount = text(period_p).strip(), text(p).strip()
            current_mode = 'table'
        else:
            if part:
                raise ValueError('Legacy rent row cannot have a continuation')
            tokens = ['']
            for n in p.iter():
                if n.tag == W+'t': tokens[-1] += n.text or ''
                elif n.tag == W+'tab' and n.getparent().tag == W+'r': tokens.append('')
            if len(tokens) != 2:
                raise ValueError('Saved rent period and amount must remain separate columns')
            period, amount = tokens; period_p = p; current_mode = 'tab'
        if mode and mode != current_mode:
            raise ValueError('Mixed rent column structures')
        mode = current_mode
        number = int((main or part)[1]); part_number = 1 if main else int(part[2])
        if len(period)>10000 or '{{' in period or '}}' in period:
            # The retained blank master is used for structure comparison.
            if not re.fullmatch(r'\{\{PERIOD_[1-9][0-9]*\}\}',period):
                raise ValueError('Rent period exceeds the notice input size limit')
        found.append({'number':number, 'part':part_number, 'period':period, 'amount':amount,
                      'no_space':bool(part and name.endswith('_NOSPACE')),
                      'p':p, 'period_p':period_p, 'paragraph_index':ps.index(p),
                      'period_paragraph_index':ps.index(period_p)})
    if not found:
        raise ValueError('Missing rent rows')
    mains = [r for r in found if r['part'] == 1]
    if not (1 if mode == 'table' else 6) <= len(mains) <= 1000:
        raise ValueError('Rent row count is invalid')
    if [r['number'] for r in mains] != list(range(1, len(mains)+1)):
        raise ValueError('Rent row bindings are missing, repeated or out of sequence')
    grouped = []
    for row in found:
        if row['part'] == 1:
            grouped.append({**row, 'parts':[row]})
        else:
            if (not grouped or row['number'] != grouped[-1]['number'] or
                    row['part'] != len(grouped[-1]['parts'])+1 or row['amount']):
                raise ValueError('Rent continuation order or amount changed')
            grouped[-1]['parts'].append(row)
            grouped[-1]['period'] += ('' if row['no_space'] else ' ') + row['period']
            if len(grouped[-1]['period'])>10000:
                raise ValueError('Rent period exceeds the notice input size limit')
    if mode == 'table':
        # Every data row must have exactly one bound amount cell. An unbound
        # inserted/deleted cell must not disappear from the saved-PDF source.
        for table in root.iter(W+'tbl'):
            if rent_table(table):
                if len(table.findall(W+'tr')) < 2:
                    raise ValueError('Empty rent continuation table')
                for tr in table.findall(W+'tr')[1:]:
                    cells = tr.findall(W+'tc')
                    bound = [x for c in cells for x in c.iter(W+'bookmarkStart')
                             if x.get(W+'name','').startswith(('BE_RENT_ROW_', 'BE_RENT_PART_'))]
                    if len(cells) != 2 or len(bound) != 1:
                        raise ValueError('Unbound or malformed rent table row')
    return grouped


def canonical_paragraphs(root):
    """Ignore only generated page artwork and validated rent display rows."""
    rent = inspect_rows(root)
    ignored = {p for r in rent for part in r['parts'] for p in (part['p'], part['period_p'])}
    for table in root.iter(W+'tbl'):
        if rent_table(table): ignored.update(table.findall('./'+W+'tr/'+W+'tc/'+W+'p'))
    result = []
    for p in root.iter(W+'p'):
        if p in ignored: continue
        value = text(p).strip()
        if not value and p.find('.//'+W+'drawing') is not None and not bookmarks(p): continue
        if not bookmarks(p) and value == 'TIME PERIOD DUEAMOUNT DUE': continue
        result.append(p)
    return result
