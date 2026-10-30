"""Recognize the existing notice's rent table; never infer missing charges."""
import re
from docx_patch import W

def inspect_rent_tables(root):
    from standard_form import is_standard, inspect_standard
    if is_standard(root):
        try:
            return inspect_standard(root)['rent_tables']
        except (ValueError, KeyError):
            return []
    paragraphs=list(root.iter(W+'p')); indices={p:i for i,p in enumerate(paragraphs)}
    def text(el):return ''.join(t.text or '' for t in el.iter(W+'t')).strip()
    def amount(cell):
        ps=list(cell.iter(W+'p')); s=text(cell)
        if len(ps)!=1 or not re.fullmatch(r'\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?',s):raise ValueError('Unclear rent cell')
        whole,_,fraction=s.replace('$','').replace(',','').partition('.')
        return {'paragraph_index':indices[ps[0]],'text':s,'cents':int(whole)*100+int(fraction.ljust(2,'0') or '0')}
    found=[]
    for table in root.iter(W+'tbl'):
        rows=table.findall(W+'tr')
        if not rows:continue
        header=[text(c).upper() for c in rows[0].findall(W+'tc')]
        if header!=['DUE DATE','AMOUNT DUE']:continue
        try:
            charges=[]; total=None
            for row in rows[1:]:
                cells=row.findall(W+'tc')
                if len(cells)!=2:raise ValueError('Unexpected rent table shape')
                label=text(cells[0]); value=text(cells[1])
                if not label and not value:continue
                if label.upper().rstrip(':')=='TOTAL AMOUNT DUE':
                    if total is not None:raise ValueError('Repeated total')
                    total=amount(cells[1])
                else:
                    if total is not None or not label:raise ValueError('Unclear charge row')
                    charges.append({'label':label,**amount(cells[1])})
            if not charges or total is None or sum(r['cents'] for r in charges)!=total['cents']:raise ValueError('Unreconciled rent table')
            found.append({'rows':charges,'total':total})
        except ValueError:
            # Do not offer an unsafe table to the revision planner.
            return []
    return found
