"""Exercise the saved-document editor against the actual fixed form."""
from copy import deepcopy
from io import BytesIO
from pathlib import Path
import sys, zipfile
import pytest
from lxml import etree as E
sys.path.insert(0, str(Path(__file__).parents[1]/'revision-service'))
from docx_patch import patch_docx, W
from standard_form import inspect_standard
from residential_notice import render_residential_notice
sys.path.insert(0, str(Path(__file__).parent))
from test_residential_standard import case

def root(raw):
    return E.fromstring(zipfile.ZipFile(BytesIO(raw)).read('word/document.xml'))

@pytest.mark.parametrize('count', [1,3,6])
def test_rent_edit_reconciles_and_preserves_every_layout_element(count):
    raw=render_residential_notice(case(count))
    before=root(raw);table=inspect_standard(before)['rent_tables'][0]
    row,total=table['rows'][0],table['total']
    changes=[{'old':row['text'],'new':'$1,250.25','paragraph_index':row['paragraph_index']},
             {'old':total['text'],'new':f'${(total["cents"]+25025)/100:,.2f}','paragraph_index':total['paragraph_index']}]
    result=patch_docx(raw,changes);after=root(result)
    assert inspect_standard(after)['rent_tables'][0]['total']['cents']==total['cents']+25025
    a,b=zipfile.ZipFile(BytesIO(raw)),zipfile.ZipFile(BytesIO(result))
    for name in a.namelist():
        if name!='word/document.xml':assert a.read(name)==b.read(name)
    for r in [before,after]:
        for t in r.iter(W+'t'):
            t.text='';t.attrib.pop('{http://www.w3.org/XML/1998/namespace}space',None)
    assert E.tostring(before)==E.tostring(after)
    with pytest.raises(ValueError):patch_docx(raw,changes[:1])

@pytest.mark.parametrize('changes',[
    [{'old':'THREE (3) DAY NOTICE TO PAY OR QUIT','new':'DIFFERENT TITLE'}],
    [{'old':'Sam Example','new':'W'*200}],
    [{'old':'$3,000.00','new':'$2,000.00','paragraph_index':26}],
    [{'old':'Sam Example','new':'{{SIGNER}}'}],
])
def test_unsafe_fixed_form_edits_are_held(changes):
    with pytest.raises(ValueError):patch_docx(render_residential_notice(case(3)),changes)

def test_both_phone_slots_can_change_without_touching_footer_or_rent():
    raw=render_residential_notice(case(3));fields=inspect_standard(root(raw))['form_fields']
    changes=[{'old':f['text'],'new':('Phone: ' if f['name']=='PROVIDER_PHONE' else '')+'(408) 555-0123','paragraph_index':f['paragraph_index']}
             for f in fields if f['name'] in ['PAYMENT_PHONE','PROVIDER_PHONE']]
    out=patch_docx(raw,changes)
    assert inspect_standard(root(out))['rent_tables']==inspect_standard(root(raw))['rent_tables']
    assert ''.join(t.text or '' for t in root(out).iter(W+'t')).count('(408) 555-0123')==2

def test_original_table_notices_keep_their_revision_path():
    from docx import Document
    d=Document();t=d.add_table(rows=3,cols=2)
    for row,vals in zip(t.rows,[['DUE DATE','AMOUNT DUE'],['September 1, 2026','$1,000.00'],['TOTAL AMOUNT DUE:','$1,000.00']]):
        for c,v in zip(row.cells,vals):c.text=v
    b=BytesIO();d.save(b)
    from rent_table import inspect_rent_tables
    assert inspect_rent_tables(root(b.getvalue()))[0]['total']['cents']==100000
    out=patch_docx(b.getvalue(),[{'old':'September 1, 2026','new':'October 1, 2026'}])
    assert inspect_rent_tables(root(out))[0]['rows'][0]['label']=='October 1, 2026'
