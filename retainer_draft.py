"""Render an unsigned review agreement; never sends, signs, or records payment."""
import base64
import hashlib
import io
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, ArrayObject

TEMPLATE = Path(__file__).parent / 'templates' / 'santa_clara_retainer.pdf'
TEMPLATE_SHA = '20457475f9728c3ebf29196414fc20d4006a3e6134315906cce5e6b4ea450927'
NOTICE_SHA = '202cfbd160a62f7afa74e25983d4389fb4e396ddfd207d025bade463095876e4'


def render_test_retainer(plan):
    expected = {
        'test_only': True, 'record_id': 'recxLwYEXtb2ZrWen',
        'eviction_id': 'recuIADHYf86dTEUR',
        'case_folder_id': '012HN2NSVZVBQ5IDEF7VBZOVFEHMF6SYJN',
        'source_notice_sha256': NOTICE_SHA,
        'state': 'ready_for_owner_price_review', 'retainer_price': 3395,
        'county': 'Santa Clara County', 'property_type': 'residential',
        'tenant_names': 'Test Tenant Maddy',
    }
    if any(type(plan.get(k)) is not type(v) or plan.get(k) != v for k, v in expected.items()):
        raise ValueError('Retainer plan outside verified synthetic review scope')
    if plan.get('missing_inputs') or plan.get('conflicting_inputs'):
        raise ValueError('Retainer plan unresolved')
    case_name = plan.get('case_name', '')
    address = plan.get('property_address', '')
    if case_name != 'Tenant Maddy v. Maddy' or address != '123 Test St, Santa Clara, CA 95050':
        raise ValueError('Synthetic case display fields changed')
    template = TEMPLATE.read_bytes()
    if hashlib.sha256(template).hexdigest() != TEMPLATE_SHA:
        raise ValueError('Retainer template changed; review required')
    values = {
        'SHORT CASE NAME': case_name, 'DEFENDANT': plan['tenant_names'],
        'DEFENDANT STREET ADDRESS': address, 'COUNTY OF': 'Santa Clara',
        'TOTAL': '$3,395.00', 'PAPERCHECK': '', 'CREDITCARD': '',
    }
    reader = PdfReader(io.BytesIO(template))
    fields = reader.get_fields() or {}
    if not set(values).issubset(fields):
        raise ValueError('Retainer fields missing')
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    # The template repeats SHORT CASE NAME as child widgets. pypdf 4 reads
    # their own value when building appearances, not the parent field value.
    # Set both consistently before generating appearances on every page.
    for page in writer.pages:
        for ref in page.get('/Annots', []):
            widget = ref.get_object()
            field = widget.get('/Parent', widget).get_object()
            name = field.get('/T')
            if name in values:
                widget[NameObject('/V')] = TextStringObject(values[name])
                widget[NameObject('/FT')] = field.get('/FT', NameObject('/Tx'))
                widget[NameObject('/DA')] = field.get('/DA', TextStringObject('/Helv 10 Tf 0 g'))
                field[NameObject('/V')] = TextStringObject(values[name])
        # A reset button obscures the header and can wipe a reviewed form.
        page[NameObject('/Annots')] = ArrayObject([
            ref for ref in page.get('/Annots', [])
            if ref.get_object().get('/T') != 'RESET FORM'])
        if '/Annots' in page:
            writer.update_page_form_field_values(page, values, auto_regenerate=False)
            for ref in page['/Annots']:
                widget = ref.get_object()
                field = widget.get('/Parent', widget).get_object()
                name = field.get('/T')
                if name in values:
                    widget[NameObject('/V')] = TextStringObject(values[name])
                    writer._update_text_field(widget)
    form = writer._root_object['/AcroForm']
    form[NameObject('/Fields')] = ArrayObject([ref for ref in form['/Fields'] if ref.get_object().get('/T') != 'RESET FORM'])
    output = io.BytesIO()
    writer.write(output)
    pdf = output.getvalue()
    saved = PdfReader(io.BytesIO(pdf))
    canonical = saved.get_fields() or {}
    for key, value in values.items():
        if str(canonical.get(key, {}).get('/V', '')) != value:
            raise ValueError('Retainer canonical field did not save: ' + key)
    seen = set()
    for page in saved.pages:
        for ref in page.get('/Annots', []):
            widget = ref.get_object()
            if widget.get('/Subtype') != '/Widget':
                continue
            field = widget
            while '/T' not in field and '/Parent' in field:
                field = field['/Parent'].get_object()
            name = field.get('/T')
            if name not in values:
                continue
            value_holder = widget
            while '/V' not in value_holder and '/Parent' in value_holder:
                value_holder = value_holder['/Parent'].get_object()
            if str(value_holder.get('/V', '')) != values[name]:
                raise ValueError('Retainer widget value mismatch: ' + name)
            appearance = widget.get('/AP', {}).get('/N')
            if appearance is None or not appearance.get_object().get_data():
                raise ValueError('Retainer appearance missing: ' + name)
            seen.add(name)
    if seen != set(values):
        raise ValueError('Retainer widgets missing')
    digest = hashlib.sha256(pdf).hexdigest()
    return {
        'test_only': True, 'record_id': plan['record_id'],
        'eviction_id': plan['eviction_id'], 'folder_id': plan['case_folder_id'],
        'source_notice_sha256': NOTICE_SHA, 'template_sha256': TEMPLATE_SHA,
        'sha256': digest, 'mime_type': 'application/pdf',
        'filename': 'TEST - Maddy retainer draft - 3395.pdf',
        'pdf_base64': base64.b64encode(pdf).decode(),
        'fields': values, 'page_count': len(saved.pages),
        'state': 'unsigned_draft_for_review', 'retainer_price': 3395,
        'payment_method': 'not_selected', 'approved': False, 'sent': False,
        'signed': False, 'payment_verified': False,
    }
