"""Apply explicit, uniquely anchored text replacements to the current Word file.
No template regeneration. Other ZIP members are copied byte-for-byte.
"""
import io,zipfile,copy
from lxml import etree
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML='{http://www.w3.org/XML/1998/namespace}'
def patch_docx(raw,changes):
    src=zipfile.ZipFile(io.BytesIO(raw))
    if any(n.startswith('_xmlsignatures/') for n in src.namelist()):
        raise ValueError('Signed document requires human review')
    root=etree.fromstring(src.read('word/document.xml'))
    before=copy.deepcopy(root)
    for tag in ['ins','del','moveFrom','moveTo']:
        if root.find('.//'+W+tag) is not None: raise ValueError('Resolve tracked changes before revising')
    plans=[]
    paragraphs=list(root.iter(W+'p'))
    occupied={}
    for change in changes:
        old,new=change['old'],change['new']
        if not isinstance(old,str) or not old or not isinstance(new,str) or any(c in old+new for c in '\r\n\t'):
            raise ValueError('Replacement must be nonempty anchored paragraph text')
        hits=[]
        selected=change.get('paragraph_index')
        if selected is not None and (type(selected) is not int or selected<0 or selected>=len(paragraphs)):
            raise ValueError('Invalid paragraph binding')
        for i,p in enumerate(paragraphs):
            if selected is not None and i!=selected:continue
            # Do not patch field results or complex content by guessing.
            if p.find('.//'+W+'fldChar') is not None or p.find('.//'+W+'fldSimple') is not None: continue
            ts=list(p.iter(W+'t'));text=''.join(t.text or '' for t in ts)
            start=0
            while (start:=text.find(old,start))>=0:
                hits.append((i,start,start+len(old),ts));start+=len(old)
        if change.get('all_occurrences') is True:
            expected=change.get('expected_occurrences')
            if type(expected) is not int or expected<2 or len(hits)!=expected:
                raise ValueError('Repeated field count changed; review again')
            if any(a!=0 or b!=sum(len(t.text or '') for t in ts) for i,a,b,ts in hits):
                raise ValueError('Repeated replacement must match a whole field paragraph')
        elif len(hits)!=1: raise ValueError('Requested text missing or ambiguous in current saved document')
        for i,a,b,ts in hits:
            if any(a<y and x<b for x,y in occupied.get(i,[])): raise ValueError('Overlapping replacements')
            occupied.setdefault(i,[]).append((a,b));plans.append((i,a,b,ts,new))
    # Right-to-left application preserves original character offsets.
    for i,a,b,ts,new in sorted(plans,key=lambda x:(x[0],x[1]),reverse=True):
        offset=0;inserted=False
        for t in ts:
            text=t.text or '';end=offset+len(text)
            if end>a and offset<b:
                lo=max(0,a-offset);hi=min(len(text),b-offset)
                t.text=text[:lo]+(new if not inserted else '')+text[hi:]
                t.set(XML+'space','preserve');inserted=True
            offset=end
    from standard_form import validate_edit
    validate_edit(before,root)
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w') as dst:
        for entry in src.infolist():
            data=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True) if entry.filename=='word/document.xml' else src.read(entry.filename)
            dst.writestr(entry,data)
    return out.getvalue()
