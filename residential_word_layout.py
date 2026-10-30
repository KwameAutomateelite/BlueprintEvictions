"""Apply the approved PDF's measured overflow plan to the editable Word copy.

Keep every static paragraph and named field from the retained Word master.
Only page artwork, rent tables, spacing and page breaks are laid out again.
"""
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.text.paragraph import Paragraph

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def needs_flow(values, count):
    from residential_pdf import _width
    widths = (163.379, 159.536, 163.537, 163.379, 159.536, 163.537)
    return count > 6 or any(_width(values[f'PERIOD_{i}'], 'TiRo') > widths[i-1]-4 or
        _width(values[f'AMOUNT_{i}'], 'TiRo') > 71.636 for i in range(1,min(count,6)+1))


def apply_flow_layout(doc, values, count):
    from residential_pdf_flow import plan_pages, PERIOD_X, AMOUNT_X, AMOUNT_WIDTH, ROW_LEADING, ROW_PADDING
    from notice_word_rows import bookmarks, text
    body = doc.element.body
    original = list(body)
    ps = [p for p in original if p.tag == W+'p']
    artwork = next(p for p in ps if not text(p).strip() and p.find('.//'+W+'drawing') is not None)
    # Select from static phrases/bookmarks, never case values or guessed names.
    def starts(prefix):
        found=[p for p in ps if text(p).startswith(prefix)]
        if len(found)!=1: raise RuntimeError('Word layout master block changed: '+prefix)
        return found[0]
    def bound(name):
        found=[p for p in ps if 'BE_'+name in bookmarks(p)]
        if len(found)!=1: raise RuntimeError('Word layout master field changed: '+name)
        return found[0]
    groups = {
     'premises and introduction': [(starts('THREE (3)'),98.2),(bound('TENANT_NAMES'),130),
      (starts('Tenants in possession'),158),(bound('PROPERTY_STREET'),180),(bound('PROPERTY_CITY'),194),
      (bound('COUNTY'),212),(starts('[hereinafter'),227),(starts('PLEASE TAKE NOTICE'),254)],
     'payment or possession paragraph':[(starts('You are further notified that within'),439)],
     'service-date paragraph':[(starts('This notice was served'),492)],
     'payment recipient and contact':[(starts('Payment of the above'),543.5),(bound('PAYMENT_RECIPIENT'),571),
      (bound('PAYMENT_STREET'),586),(bound('PAYMENT_CITY'),600),(bound('PAYMENT_PHONE'),616)],
     'payment business hours':[(starts('Rent or possession'),630)],
     'payment conditions paragraph':[(starts('Only the total'),670.5)],
     'forfeiture paragraph':[(starts('You are further notified by this Notice'),98)],
     'California heading and credit paragraph':[(starts('NOTICES FROM THE STATE'),140.5),(starts('(A) NOTICE'),167)],
     'abandoned property paragraph':[(starts('(B) NOTICE'),236)],
     'unlawful detainer paragraph':[(starts('Please be advised'),346)],
     'superseded notices paragraph':[(starts('Finally,'),401)],
     'signature and housing provider contact':[(bound('NOTICE_DATE'),484),(bound('SIGNER_NAME_AND_TITLE'),498),
      (starts("Housing Provider's Address:"),540),(bound('PROVIDER_STREET'),556),(bound('PROVIDER_CITY'),571),
      (bound('PROVIDER_PHONE'),585.5)],
    }
    total = next(t for t in original if t.tag==W+'tbl' and text(t).startswith('TOTAL AMOUNT DUE:'))
    sect = next(x for x in original if x.tag==W+'sectPr')
    for node in original:
        if node is not sect:body.remove(node)
    cursor=0
    bookmark_id=10000
    last_table=None
    remaining={i:values[f'PERIOD_{i}'] for i in range(1,count+1)}
    def append(node):body.insert(len(body)-1,node)
    def place(p, top, height=None):
        nonlocal cursor,last_table
        p=deepcopy(p); paragraph=Paragraph(p,doc._body); pf=paragraph.paragraph_format
        pf.page_break_before=False;pf.keep_with_next=False;pf.keep_together=True
        gap=top-cursor
        if gap < -.6: raise RuntimeError(f'Word flow overlap: {text(p)[:40]} ({gap:.2f}pt)')
        pf.space_before=Pt(max(0,gap));pf.space_after=Pt(0)
        append(p)
        leading=pf.line_spacing.pt
        lines=sum(1 for x in p.iter(W+'br') if x.get(W+'type')!='page')+1
        cursor=top+(height if height is not None else leading*lines)
        # A paragraph border occupies extra height outside its text line box.
        if p.find('./'+W+'pPr/'+W+'pBdr') is not None:cursor+=6.75
        last_table=None
        return p
    def new_page(top, index):
        nonlocal cursor,last_table
        if index:
            # Reusing the first paragraph's Word paraId/BodyText style can
            # attach its floating artwork to page one again in Word/LO.
            a=OxmlElement('w:p')
            for child in artwork:
                if child.tag != W+'pPr':a.append(deepcopy(child))
        else:a=deepcopy(artwork)
        para=Paragraph(a,doc._body);pf=para.paragraph_format
        pf.page_break_before=bool(index);pf.space_before=pf.space_after=Pt(0);pf.line_spacing=Pt(top)
        pf.keep_with_next=True;append(a);cursor=top;last_table=None
    def cell_properties(cell, left, right):
        cp=cell._tc.get_or_add_tcPr()
        margins=OxmlElement('w:tcMar')
        for edge,val in [('top',0),('bottom',0),('left',left),('right',right)]:
            n=OxmlElement('w:'+edge);n.set(qn('w:w'),str(round(val*20)));n.set(qn('w:type'),'dxa');margins.append(n)
        cp.append(margins)
        borders=OxmlElement('w:tcBorders')
        for edge in ('top','bottom','left','right'):
            n=OxmlElement('w:'+edge);n.set(qn('w:val'),'nil');borders.append(n)
        cp.insert_element_before(borders,'w:shd','w:noWrap','w:tcMar','w:textDirection','w:tcFitText','w:vAlign')
        # Exact row leading and explicit line wrapping preserve 12pt type.
        para=cell.paragraphs[0];pf=para.paragraph_format
        pf.space_before=pf.space_after=Pt(0);pf.line_spacing=Pt(ROW_LEADING)
        pf.keep_with_next=False;pf.keep_together=True;pf.widow_control=False
        return para
    def rent_table(top):
        nonlocal cursor,last_table
        gap=top-cursor
        if gap < -.6:raise RuntimeError('Word rent heading overlaps the introduction')
        # Use the preceding paragraph's trailing spacing, not tiny empty
        # spacers that Word Online expands into unwanted pages.
        if gap > 0:
            prior=body[-2]
            if prior.tag != W+'p':raise RuntimeError('Missing rent heading anchor')
            Paragraph(prior,doc._body).paragraph_format.space_after=Pt(gap)
        t=doc.add_table(rows=0,cols=2);t.autofit=False
        width1=AMOUNT_X-PERIOD_X; width2=AMOUNT_WIDTH
        t.columns[0].width=Pt(width1);t.columns[1].width=Pt(width2)
        props=t._tbl.tblPr
        indent=OxmlElement('w:tblInd');indent.set(qn('w:w'),str(round((PERIOD_X-54)*20)));indent.set(qn('w:type'),'dxa')
        props.insert_element_before(indent,'w:tblBorders','w:shd','w:tblLayout','w:tblCellMar','w:tblLook')
        row=t.add_row();row.height=Pt(17.5);row.height_rule=WD_ROW_HEIGHT_RULE.EXACTLY
        repeat=OxmlElement('w:tblHeader');row._tr.get_or_add_trPr().append(repeat)
        for c,label in zip(row.cells,('TIME PERIOD DUE','AMOUNT DUE')):
            p=cell_properties(c,2,2);p.paragraph_format.line_spacing=Pt(14.5)
            p.paragraph_format.keep_with_next=True
            rr=p.add_run(label);rr.font.name='Times New Roman';rr.font.size=Pt(11.5)
            rr.bold=rr.underline=True
        # Preserve order: add_table inserts at the end, ahead of sectPr.
        last_table=t;cursor=top+17.5
    def row(item):
        nonlocal cursor,bookmark_id
        t=last_table
        if t is None:raise RuntimeError('Rent row without heading/table')
        r=t.add_row();r.height=Pt(item['height']);r.height_rule=WD_ROW_HEIGHT_RULE.EXACTLY
        no_split=OxmlElement('w:cantSplit');r._tr.get_or_add_trPr().append(no_split)
        width1=AMOUNT_X-PERIOD_X
        original=remaining[item['number']]
        no_space=item['part']>1 and original and not original[0].isspace()
        used=0
        for line in item['lines']:
            while used<len(original) and original[used].isspace():used+=1
            if not original[used:].startswith(line):raise RuntimeError('Rent wrap text did not match source')
            used+=len(line)
        shown=original[:used].strip();remaining[item['number']]=original[used:]
        for i,c in enumerate(r.cells):
            c.width=Pt(width1 if i==0 else AMOUNT_WIDTH)
            p=cell_properties(c,2,(width1-164.8)+2 if i==0 else 2)
            # Keep the actual period text intact; let the fixed-width cell
            # wrap it naturally instead of inserting spaces into long words.
            content=shown if i==0 else item['amount']
            rr=p.add_run(content);rr.font.name='Times New Roman';rr.font.size=Pt(12)
            if i==1:
                bookmark_id+=1
                name='BE_RENT_ROW_'+str(item['number']) if item['part']==1 else f"BE_RENT_PART_{item['number']}_{item['part']}"
                if item['part']>1 and no_space:name+='_NOSPACE'
                mark=OxmlElement('w:bookmarkStart');mark.set(qn('w:id'),str(bookmark_id));mark.set(qn('w:name'),name)
                p._p.insert(1,mark);end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(bookmark_id));p._p.append(end)
        cursor=item['at']+item['height']
    pages=plan_pages(values,count)
    for index,placements in enumerate(pages):
        first=placements[0]
        first_top=(first['at'] + (groups[first['label']][0][1]-first['top'])
                   if first['kind']=='segment' and first['label'] in groups else first['at']-1)
        new_page(first_top,index)
        for item in placements:
            if item['kind']=='row':row(item);continue
            label=item['label']
            if label=='rent column headings':rent_table(item['at']-1);continue
            if label=='total after final rent entry':
                # Leave the measured four-point gap inside the unbordered
                # last rent row, avoiding an empty spacer between two tables.
                if last_table is None:raise RuntimeError('Total without final rent row')
                last=last_table.rows[-1]
                last.height=Pt(last.height.pt+4)
                t=deepcopy(total);append(t);cursor=item['at']+18;last_table=None
                continue
            for p,y in groups[label]:place(p,item['at']+y-item['top'])
    for i,prop in enumerate(doc.element.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'),1):
        prop.set('id',str(i))
    return pages
