"""Render an unsigned review agreement; never sends, signs, or records payment."""
import base64
import hashlib
import io
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pypdf._cmap import build_char_map
from pypdf.generic import (NameObject, TextStringObject, ArrayObject,
                           DictionaryObject, NumberObject, RectangleObject,
                           ByteStringObject, ContentStream, DecodedStreamObject, FloatObject)

TEMPLATE = Path(__file__).parent / 'templates' / 'santa_clara_retainer.pdf'
TEMPLATE_SHA = '20457475f9728c3ebf29196414fc20d4006a3e6134315906cce5e6b4ea450927'
NOTICE_SHA = '202cfbd160a62f7afa74e25983d4389fb4e396ddfd207d025bade463095876e4'
VERIFICATION_VERSION = 'retainer-office-fields-2026-11-10'
COUNTY_VERSION = 'retainer-common-county-form-2026-11-10'
PREPARATION_VERSION = 'retainer-notice-or-eviction-2026-11-11'
SUPPORTED_COUNTIES = ('Santa Clara', 'Santa Cruz', 'San Benito', 'San Mateo')


def _property_county(value):
    if not isinstance(value, str):
        raise ValueError('Supported property county required')
    normalized = re.sub(r'\s+county$', '', value.strip(), flags=re.I).casefold()
    for county in SUPPORTED_COUNTIES:
        if county.casefold() == normalized:
            return county
    raise ValueError('Property county is outside the reviewed business county configuration')


def _cid_widths(font):
    descendant = font['/DescendantFonts'][0].get_object()
    source, widths, i = descendant['/W'], {}, 0
    while i < len(source):
        first, second = int(source[i]), source[i + 1]
        i += 2
        if isinstance(second, list):
            widths.update({first + j: float(width) for j, width in enumerate(second)})
        else:
            width = float(source[i])
            i += 1
            widths.update({j: width for j in range(first, int(second) + 1)})
    return widths


def _fill_printed_county(writer, county):
    """AM June26 12:26: one agreement covers her counties; county is case data.

    Replace only the two source county runs. Reuse the exact embedded Inter
    font from page 2, whose subset includes the z needed by Santa Cruz. Keep
    all source text matrices; adjust the following space by the measured
    word-width difference. No redaction, overlay, font substitution or reflow.
    """
    if county == 'Santa Clara':
        return
    county_font = writer.pages[1]['/Resources']['/Font']['/F19']
    info = build_char_map('/F19', 200, writer.pages[1])
    reverse = {text: code for code, text in info[3].items()
               if isinstance(code, str) and isinstance(text, str) and len(text) == 1}
    if info[2] != 'utf-16-be' or any(c not in reverse for c in county):
        raise ValueError('Original county font cannot represent this county')
    encoded = ''.join(reverse[c] for c in county).encode('utf-16-be')
    widths = _cid_widths(county_font.get_object())
    xobjects = writer.pages[0]['/Resources']['/XObject']
    replaced = 0
    for key, ref in list(xobjects.items()):
        form = ref.get_object()
        if form.get('/Subtype') != '/Form' or '/Font' not in form.get('/Resources', {}):
            continue
        maps = {name: build_char_map(name, 200, form) for name in form['/Resources']['/Font']}
        stream = ContentStream(form, writer)
        flat, positions, font, size = '', [], None, None
        for index, (args, operator) in enumerate(stream.operations):
            if operator == b'Tf':
                font, size = str(args[0]), float(args[1])
            if operator != b'Tj':
                continue
            value = args[0]
            raw = value.original_bytes if hasattr(value, 'original_bytes') else bytes(value)
            encoding, cmap = maps[font][2:4]
            text = raw.decode(encoding) if isinstance(encoding, str) else ''.join(encoding[c] for c in raw)
            text = ''.join(cmap.get(c, c) for c in text)
            for char in text:
                flat += char
                positions.append((index, font, size, len(text)))
        matches = list(re.finditer(r'Santa\s+Clara', flat))
        if not matches:
            continue
        if len(matches) != 1:
            raise ValueError('Source county layout changed')
        match = matches[0]
        run = positions[match.start():match.end()]
        first, old_font, size, _ = run[0]
        last = run[-1][0]
        if any(f != old_font or s != size or length != 1 for _, f, s, length in run):
            raise ValueError('Source county glyph run changed')
        following, operator = stream.operations[last + 1]
        region = stream.operations[first + 1:last + 2]
        if operator != b'Td' or float(following[1]) != 0 or any(op not in (b'Td', b'Tj') for _, op in region):
            raise ValueError('Source county spacing changed')
        old_width = sum(float(args[0]) for args, op in region if op == b'Td')
        new_width = sum(widths[ord(reverse[c])] for c in county) * size / 1000
        if new_width > old_width + .1:
            raise ValueError('County exceeds its source text allocation')
        following[0] = FloatObject(float(following[0]) + new_width - old_width)
        old_indices = {index for index, _, _, _ in run}
        operations = []
        for index, operation in enumerate(stream.operations):
            if index == first:
                operations += [([NameObject('/RetainerCounty'), FloatObject(size)], b'Tf'),
                               ([ByteStringObject(encoded)], b'Tj'),
                               ([NameObject(old_font), FloatObject(size)], b'Tf')]
            elif index not in old_indices:
                operations.append(operation)
        stream.operations = operations
        form['/Resources']['/Font'][NameObject('/RetainerCounty')] = county_font
        replacement = DecodedStreamObject()
        replacement.update({k: v for k, v in form.items() if k not in ('/Filter', '/DecodeParms', '/Length')})
        replacement.set_data(stream.get_data())
        xobjects[key] = writer._add_object(replacement)
        replaced += 1
    if replaced != 2:
        raise ValueError('Expected both printed source county references')


def _verification_values(plan):
    """Only explicit, sourced office facts; signatures and dates stay empty."""
    office = plan.get('verification') or {}
    allowed = {'capacity', 'printed_name', 'title', 'plaintiff_name',
               'entity_name', 'execution_county', 'payment_method', 'source'}
    if not isinstance(office, dict) or set(office) - allowed:
        raise ValueError('Invalid verification fields')
    for key, value in office.items():
        if not isinstance(value, str) or any(ord(c) < 32 for c in value):
            raise ValueError('Invalid verification value: ' + key)
    office = {key: value.strip() for key, value in office.items()}
    if any(value for key, value in office.items() if key != 'source') and not office.get('source'):
        raise ValueError('Verification facts require their source')
    capacity = office.get('capacity', '')
    capacities = {'plaintiff', 'agent', 'llc_member', 'general_partner',
                  'limited_partner', 'trustee', 'successor_trustee'}
    if capacity and capacity not in capacities:
        raise ValueError('Unsupported verification capacity')
    if office.get('payment_method', '') not in ('', 'check', 'card'):
        raise ValueError('Unsupported payment method')
    if office.get('plaintiff_name') and capacity != 'agent':
        raise ValueError('Vested owner line applies to the agent capacity')
    entity_capacities = {'llc_member', 'general_partner', 'limited_partner',
                         'trustee', 'successor_trustee'}
    if office.get('entity_name') and capacity not in entity_capacities:
        raise ValueError('Entity name requires its supported capacity')
    missing = [key for key in ('capacity', 'printed_name', 'title', 'execution_county', 'payment_method')
               if not office.get(key)]
    if capacity == 'agent' and not office.get('plaintiff_name'):
        missing.append('plaintiff_name')
    if capacity in entity_capacities and not office.get('entity_name'):
        missing.append('entity_name')
    # Fixed coordinates measured from AM's unchanged verification page.
    # Each rectangle sits above its existing line; no legal text is replaced.
    fields = {
        'VERIFICATION PRINTED NAME': (office.get('printed_name', ''), (174, 92, 312, 105), 12),
        'VERIFICATION TITLE': (office.get('title', ''), (102, 64.5, 312, 77.5), 12),
        'VERIFICATION COUNTY': (office.get('execution_county', ''), (445, 195.5, 527, 208.5), 12),
        'VERIFICATION VESTED OWNER': (office.get('plaintiff_name', ''), (387, 552.5, 535, 565.5), 12),
    }
    for key, y in {'plaintiff': 634.86, 'agent': 607.26, 'llc_member': 524.47,
                   'partnership': 469.28, 'trustee': 414.07,
                   'successor_trustee': 372.68}.items():
        selected = capacity == key or (key == 'partnership' and capacity in
                                      ('general_partner', 'limited_partner'))
        fields['VERIFICATION CAPACITY ' + key.upper()] = ('X' if selected else '',
                                                          (107, y + .5, 119, y + 11.5), 10)
    entity = office.get('entity_name', '')
    for key, rect in {
        'llc_member': (253, 497.4, 535, 510.4),
        'general_partner': (339, 456, 535, 469),
        'limited_partner': (339, 442.2, 535, 455.2),
        'trustee': (238, 400.8, 535, 413.8),
        'successor_trustee': (238, 359.4, 535, 372.4),
    }.items():
        fields['VERIFICATION ENTITY ' + key.upper()] = (entity if capacity == key else '', rect, 12)
    for key, y in [('general_partner', 455.47), ('limited_partner', 441.68)]:
        fields['VERIFICATION PARTNERSHIP ' + key.upper()] = (
            'X' if capacity == key else '', (180, y + .5, 192, y + 11.5), 10)
    return office, fields, missing


def _add_verification_fields(writer, definitions):
    form = writer._root_object['/AcroForm']
    font = form['/DR']['/Font']['/TimesNewRoman'].get_object()
    first, widths = int(font['/FirstChar']), font['/Widths']
    page = writer.pages[2]
    if '/Annots' not in page:
        page[NameObject('/Annots')] = ArrayObject()
    for name, (value, rect, size) in definitions.items():
        try:
            encoded = value.encode('cp1252')
        except UnicodeEncodeError:
            raise ValueError('Verification text is outside the template font: ' + name)
        units = sum(float(widths[c - first]) for c in encoded)
        width = rect[2] - rect[0] - 4
        if units:
            size = min(size, width * 1000 / units)
        if size < 9:
            raise ValueError('Verification text exceeds its allocated line: ' + name)
        # pypdf puts the baseline at top - font size - 1. Keep that baseline
        # at the source line when a long supported name needs a smaller font.
        rect = (rect[0], rect[1], rect[2], rect[1] + size + 1)
        widget = DictionaryObject({
            NameObject('/Type'): NameObject('/Annot'),
            NameObject('/Subtype'): NameObject('/Widget'),
            NameObject('/FT'): NameObject('/Tx'),
            NameObject('/T'): TextStringObject(name),
            NameObject('/Rect'): RectangleObject(rect),
            NameObject('/DA'): TextStringObject(f'/TimesNewRoman {size:.4f} Tf 0 g'),
            NameObject('/V'): TextStringObject(value),
            NameObject('/F'): NumberObject(4),
            NameObject('/P'): page.indirect_reference,
        })
        ref = writer._add_object(widget)
        page['/Annots'].append(ref)
        form['/Fields'].append(ref)
        writer._update_text_field(widget)


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
    result = render_retainer_draft(plan)
    result['filename'] = 'TEST - Maddy retainer draft - 3395.pdf'
    return result


def render_retainer_draft(plan):
    """Populate an unsigned Santa Clara agreement from an explicit case plan."""
    if not isinstance(plan, dict):
        raise ValueError('Expected a case plan')
    required = ['record_id', 'case_folder_id', 'case_name',
                'property_address', 'tenant_names']
    for key in required:
        if not isinstance(plan.get(key), str) or not plan[key].strip():
            raise ValueError('Missing case field: ' + key)
    case_source = plan.get('case_source', 'linked_eviction')
    eviction_id = plan.get('eviction_id')
    if case_source == 'notice_only':
        if eviction_id is not None:
            raise ValueError('Notice-only preparation must not invent an eviction record')
    elif case_source != 'linked_eviction' or not isinstance(eviction_id, str) or not eviction_id.strip():
        raise ValueError('Linked eviction preparation requires its actual record')
    if plan.get('missing_inputs') or plan.get('conflicting_inputs'):
        raise ValueError('Retainer plan unresolved')
    county = _property_county(plan.get('county'))
    if plan.get('property_type') not in ('residential', 'commercial'):
        raise ValueError('Property type required')
    try:
        amount = Decimal(str(plan.get('retainer_price', '')))
        if not amount.is_finite() or amount <= 0 or amount != amount.quantize(Decimal('0.01')):
            raise ValueError('Invalid retainer amount')
    except (InvalidOperation, TypeError):
        raise ValueError('Invalid retainer amount')
    case_name = plan['case_name'].strip()
    address = plan['property_address'].strip()
    template = TEMPLATE.read_bytes()
    if hashlib.sha256(template).hexdigest() != TEMPLATE_SHA:
        raise ValueError('Retainer template changed; review required')
    office, verification_fields, missing_office = _verification_values(plan)
    values = {
        'SHORT CASE NAME': case_name, 'DEFENDANT': plan['tenant_names'],
        'DEFENDANT STREET ADDRESS': address, 'COUNTY OF': county,
        'TOTAL': f'${amount:,.2f}',
        'PAPERCHECK': 'X' if office.get('payment_method') == 'check' else '',
        'CREDITCARD': 'X' if office.get('payment_method') == 'card' else '',
    }
    reader = PdfReader(io.BytesIO(template))
    fields = reader.get_fields() or {}
    if not set(values).issubset(fields):
        raise ValueError('Retainer fields missing')
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    _fill_printed_county(writer, county)
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
    _add_verification_fields(writer, verification_fields)
    values.update({name: value for name, (value, _rect, _size) in verification_fields.items()})
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
        'test_only': plan.get('test_only') is True, 'record_id': plan['record_id'],
        'eviction_id': eviction_id, 'case_source': case_source,
        'preparation_version': PREPARATION_VERSION, 'folder_id': plan['case_folder_id'],
        'source_notice_sha256': plan.get('source_notice_sha256'),
        'source_case_sha256': plan.get('source_case_sha256'), 'template_sha256': TEMPLATE_SHA,
        'sha256': digest, 'mime_type': 'application/pdf',
        'filename': ('TEST - ' if plan.get('test_only') is True else '') + 'retainer-draft.pdf',
        'pdf_base64': base64.b64encode(pdf).decode(),
        'fields': values, 'page_count': len(saved.pages),
        'verification_version': VERIFICATION_VERSION,
        'county_version': COUNTY_VERSION, 'property_county': county,
        'office_fields_complete': not missing_office,
        'missing_office_fields': missing_office,
        'verification_source': office.get('source', ''),
        'state': 'unsigned_draft_for_review', 'retainer_price': float(amount),
        'payment_method': office.get('payment_method') or 'not_selected', 'approved': False, 'sent': False,
        'signed': False, 'payment_verified': False,
    }
