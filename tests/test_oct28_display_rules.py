"""Acceptance examples from the full October 28 call and corrected reference PDF."""
from copy import deepcopy
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import sys

import pytest
from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfReader

sys.path[:0] = [str(Path(__file__).parents[1]), str(Path(__file__).parent)]
from notice_display import rent_period, signer_title
from residential_notice import fields_for_case, render_residential_notice, NoticeInputError, MASTER
from residential_pdf import render_residential_pdf
from residential_word_snapshot import render_saved_word_pdf
from test_residential_standard import case


@pytest.mark.parametrize('raw', ['Child of Property Owner', 'SON OF OWNER', 'Wife', 'Niece', 'Grandparent', 'Daughter', 'Husband', 'Sibling'])
def test_informal_relationship_has_professional_printed_title(raw):
    assert signer_title(raw) == 'Authorized Agent for Property Owner'


@pytest.mark.parametrize('raw,expected', [('Property Owner', 'Property Owner'), ('PROPERTY OWNER', 'Property Owner'), ('Property Manager', 'Property Manager'), ('Trustee', 'Trustee'), ('', '')])
def test_owner_and_formal_titles_are_not_replaced(raw, expected):
    assert signer_title(raw) == expected


@pytest.mark.parametrize('raw,expected', [('October 1, 2025', 'October 1-31, 2025'), ('November 3, 2025', 'November 1-30, 2025'), ('2/1/2024', 'February 1-29, 2024'), ('February 2025', 'February 1-28, 2025'), ('November 15-30, 2025', 'November 15-30, 2025')])
def test_month_ranges_include_calendar_end_and_preserve_explicit_partial_periods(raw, expected):
    assert rent_period(raw) == expected


def call_case():
    data = case(12)
    data.update(tenant_names='ALEX EXAMPLE\r\nJORDAN EXAMPLE, and all other occupants',
                property_address_street='123 Example St, Unit 2', property_address_city='San Jose, CA 95116', apt='',
                landlord_name='EXAMPLE FAMILY TRUST', payment_recipient_name='', payment_address_street='', payment_address_city='',
                payment_address='PO BOX 12345 SAN JOSE, CA 95152', landlord_address='PO BOX 12345 SAN JOSE, CA 95152',
                signer_name='PAT EXAMPLE', signer_title='Child of Property Owner',
                amounts_due=[{'due_date': f'{i}/1/2026', 'amount': '1500'} for i in range(1, 13)], total_amount_due='18000')
    return data


def test_call_input_preserves_intake_and_groups_addresses_correctly():
    data = call_case(); original = deepcopy(data)
    fields = fields_for_case(data)
    assert data == original
    assert fields['TENANT_NAMES'] == 'Alex Example, Jordan Example, and All Other Occupants'
    assert fields['PAYMENT_RECIPIENT'] == 'Example Family Trust'
    assert fields['PAYMENT_STREET'] == fields['PROVIDER_STREET'] == 'P.O. Box 12345'
    assert fields['PAYMENT_CITY'] == fields['PROVIDER_CITY'] == 'San Jose, CA 95152'
    assert fields['PROPERTY_CITY'] == 'San Jose, CA 95116'
    assert ' '.join([fields['SIGNER_NAME_AND_TITLE'], fields['SIGNER_CONTINUED']]).strip() == 'Pat Example, Authorized Agent for Property Owner'
    assert fields['TOTAL_AMOUNT'] == '$18,000.00'
    assert fields['PERIOD_12'] == 'December 1-31, 2026'


def test_every_bound_word_variable_uses_times_new_roman_12():
    doc = Document(BytesIO(render_residential_notice(call_case())))
    # Track replacement-run indexes from the unchanged Word master. Static
    # prefixes/captions keep their approved typography.
    bindings = {}
    for paragraph in Document(MASTER).paragraphs:
        marks = [b.get(qn('w:name')) for b in paragraph._p.iter(qn('w:bookmarkStart'))]
        if marks:
            bindings[marks[0]] = [i for i, run in enumerate(paragraph.runs) if '{{' in run.text]
    matching = 0
    from docx.text.paragraph import Paragraph
    for node in doc.element.iter(qn('w:p')):
        paragraph=Paragraph(node,doc._body)
        marks = [b.get(qn('w:name')) for b in paragraph._p.iter(qn('w:bookmarkStart'))]
        if not marks:
            continue
        key = marks[0]
        if key.startswith('BE_RENT_ROW_'):
            # Overflow rent now uses real cells rather than tab alignment.
            for cell in paragraph._p.getparent().getparent().findall(qn('w:tc')):
                for value in cell.iter(qn('w:p')):
                    for run in Paragraph(value,doc._body).runs:
                        assert run.font.name == 'Times New Roman'
                        assert run.font.size.pt == 12
                        matching += 1
            continue
        for index in bindings.get(key, []):
            run = paragraph.runs[index]
            assert run.font.name == 'Times New Roman'
            assert run.font.size.pt == 12
            matching += 1
    assert matching >= 25
    word = render_residential_notice(call_case())
    saved = render_saved_word_pdf(word, sha256(word).hexdigest())
    assert saved['case_fields']['signer_title'] == 'Authorized Agent for Property Owner'


def test_all_pdf_inputs_use_embedded_times_new_roman_12_even_twelve_rows():
    reader = PdfReader(BytesIO(render_residential_pdf(call_case())))
    assert len(reader.pages) == 2
    for page in reader.pages:
        for ref in page.get('/Annots', []):
            widget = ref.get_object(); field = widget.get('/Parent', widget).get_object()
            if field.get('/FT') != '/Tx':
                continue
            appearance = widget['/AP']['/N']
            assert b'12 Tf' in appearance.get_data()
            for key in ('/TiRo', '/TiBo'):
                font = appearance['/Resources']['/Font'][key]
                assert str(font['/BaseFont']).startswith('/TimesNewRomanPS')
                assert len(font['/FontDescriptor']['/FontFile2'].get_data()) > 1000
    fields = reader.get_fields()
    assert fields['Month12']['/V'] == 'December 1-31, 2026'
    assert fields['PLAINTIFF CITY/STATE/ZIP']['/V'] == 'San Jose, CA 95152'


def test_incomplete_city_state_zip_is_held_instead_of_guessed():
    data = call_case(); data['property_address_city'] = 'San Jose'
    with pytest.raises(NoticeInputError, match='city/state/ZIP'):
        fields_for_case(data)


def test_date_input_aligns_with_original_signature_caption_baseline():
    import re
    reader = PdfReader(BytesIO(render_residential_pdf(case(1))))
    for ref in reader.pages[1]['/Annots']:
        widget = ref.get_object(); field = widget.get('/Parent', widget).get_object()
        if field.get('/T') == 'NAME OF PERSON SERVED':
            position = re.search(rb'1 0 0 1 [-\d.]+ ([-\d.]+) Tm', widget['/AP']['/N'].get_data())
            assert abs(float(widget['/Rect'][1]) + float(position[1]) - 296.57) < .00001
            break
    else:
        pytest.fail('Signature date field is missing')
