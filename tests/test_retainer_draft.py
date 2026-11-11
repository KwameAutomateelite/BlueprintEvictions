import base64
import unittest
from retainer_draft import render_test_retainer, NOTICE_SHA

class RetainerDraftTests(unittest.TestCase):
    def plan(self):
        return dict(test_only=True,record_id='recxLwYEXtb2ZrWen',eviction_id='recuIADHYf86dTEUR',case_folder_id='012HN2NSVZVBQ5IDEF7VBZOVFEHMF6SYJN',source_notice_sha256=NOTICE_SHA,state='ready_for_owner_price_review',retainer_price=3395,county='Santa Clara County',property_type='residential',tenant_names='Test Tenant Maddy',case_name='Tenant Maddy v. Maddy',property_address='123 Test St, Santa Clara, CA 95050',missing_inputs=[],conflicting_inputs=[])
    def test_interactive_draft_has_verified_fields(self):
        result=render_test_retainer(self.plan())
        self.assertTrue(base64.b64decode(result['pdf_base64']).startswith(b'%PDF'))
        self.assertEqual(result['page_count'],3)
        self.assertFalse(result['signed'])
        self.assertEqual(result['fields']['PAPERCHECK'],'')
    def test_live_or_changed_case_rejected(self):
        for key,value in [('test_only',False),('record_id','live'),('source_notice_sha256','changed'),('retainer_price',3500),('county','San Mateo'),('missing_inputs',['rent'])]:
            with self.subTest(key=key),self.assertRaises(ValueError):
                render_test_retainer({**self.plan(),key:value})
    def test_repeated_case_name_widgets_have_text_appearances(self):
        from pypdf import PdfReader
        import io
        result=render_test_retainer(self.plan())
        reader=PdfReader(io.BytesIO(base64.b64decode(result['pdf_base64'])))
        count=0
        for page in reader.pages:
            for ref in page.get('/Annots',[]):
                widget=ref.get_object()
                field=widget.get('/Parent',widget).get_object()
                if field.get('/T')=='SHORT CASE NAME':
                    self.assertIn(b'Tenant Maddy v. Maddy',widget['/AP']['/N'].get_object().get_data())
                    count+=1
                self.assertNotEqual(widget.get('/T'),'RESET FORM')
        self.assertEqual(count,2)

class RetainerRouteTests(unittest.TestCase):
    def test_authentication_and_test_scope(self):
        import os
        from unittest.mock import patch
        from fastapi import FastAPI
        from fastapi.testclient import TestClient
        from retainer_routes import router
        app=FastAPI();app.include_router(router)
        with patch.dict(os.environ,{'SIGNATURE_STORAGE_BRIDGE_SECRET':'test-only-secret'}):
            client=TestClient(app)
            self.assertEqual(client.post('/test-retainer-draft',json={}).status_code,401)
            headers={'x-narnia-retainer-secret':'test-only-secret'}
            self.assertEqual(client.post('/test-retainer-draft',json={'test_only':False},headers=headers).status_code,422)
            self.assertEqual(client.post('/test-retainer-draft',json=RetainerDraftTests().plan(),headers=headers).status_code,200)

class GeneralRetainerTests(unittest.TestCase):
    def test_explicit_notice_only_preparation_needs_no_eviction_record(self):
        from retainer_draft import render_retainer_draft
        result = render_retainer_draft({**self.plan(), 'case_source': 'notice_only', 'eviction_id': None})
        self.assertIsNone(result['eviction_id'])
        self.assertEqual(result['case_source'], 'notice_only')
        self.assertEqual(result['preparation_version'], 'retainer-notice-or-eviction-2026-11-11')
        self.assertEqual(result['fields']['SHORT CASE NAME'], self.plan()['case_name'])
        self.assertFalse(result['sent'])
        self.assertFalse(result['approved'])

    def test_notice_only_cannot_impersonate_a_linked_eviction(self):
        from retainer_draft import render_retainer_draft
        for extra in ({'eviction_id': None}, {'case_source': 'unknown'}, {'case_source': 'notice_only'}):
            with self.subTest(extra=extra), self.assertRaises(ValueError):
                render_retainer_draft({**self.plan(), **extra})

    def test_case_snapshot_does_not_require_a_signed_notice(self):
        from retainer_draft import render_retainer_draft
        plan = self.plan()
        plan.pop('source_notice_sha256')
        plan['source_case_sha256'] = 'a' * 64
        result = render_retainer_draft(plan)
        self.assertIsNone(result['source_notice_sha256'])
        self.assertEqual(result['source_case_sha256'], 'a' * 64)

    def plan(self):
        return {**RetainerDraftTests().plan(), 'test_only': False,
                'record_id': 'recExampleNotice01', 'eviction_id': 'recExampleEvict001',
                'case_folder_id': 'example-folder', 'case_name': 'Example Owner v. Example Tenant',
                'tenant_names': 'Example Tenant', 'property_address': '456 Example St, San Jose, CA 95112',
                'retainer_price': 3500}

    def test_general_case_and_owner_supplied_fee(self):
        from retainer_draft import render_retainer_draft
        result = render_retainer_draft(self.plan())
        self.assertEqual(result['fields']['TOTAL'], '$3,500.00')
        self.assertEqual(result['fields']['SHORT CASE NAME'], self.plan()['case_name'])
        self.assertFalse(result['test_only'])
        self.assertFalse(result['sent'])
        self.assertFalse(result['approved'])

    def test_reject_missing_case_unsupported_county_and_invalid_money(self):
        from retainer_draft import render_retainer_draft
        for key, value in [('case_folder_id', ''), ('county', 'Unreviewed County'),
                           ('retainer_price', 'NaN'), ('retainer_price', -1),
                           ('retainer_price', '3500.001'), ('conflicting_inputs', ['owner'])]:
            with self.subTest(key=key, value=value), self.assertRaises(ValueError):
                render_retainer_draft({**self.plan(), key: value})

    def test_general_route_requires_authentication(self):
        import os
        from unittest.mock import patch
        from fastapi import FastAPI
        from fastapi.testclient import TestClient
        from retainer_routes import router
        app = FastAPI(); app.include_router(router)
        with patch.dict(os.environ, {'SIGNATURE_STORAGE_BRIDGE_SECRET': 'test-only-secret'}):
            client = TestClient(app)
            self.assertEqual(client.post('/retainer-draft', json=self.plan()).status_code, 401)
            self.assertEqual(client.post('/retainer-draft', json=self.plan(), headers={
                'x-narnia-retainer-secret': 'test-only-secret'}).status_code, 200)


class VerificationFieldsTests(unittest.TestCase):
    def test_scaled_vested_owner_has_valid_visible_pdf_appearance(self):
        import io
        import re
        from pypdf import PdfReader
        result = self.render(capacity='agent', plaintiff_name='Example Property Holdings LLC',
                             printed_name='Fictional Agent', title='Property Manager')
        self.assertEqual(result['appearance_version'], 'retainer-office-appearance-2026-11-11')
        saved = PdfReader(io.BytesIO(base64.b64decode(result['pdf_base64'])))
        widget = next(ref.get_object() for ref in saved.pages[2]['/Annots']
                      if ref.get_object().get('/T') == 'VERIFICATION VESTED OWNER')
        appearance = widget['/AP']['/N'].get_object().get_data()
        self.assertIn(b'Example Property Holdings LLC', appearance)
        self.assertIsNone(re.search(rb'[-+]?\d+(?:\.\d+)?[eE][-+]?\d+', appearance))
        self.assertEqual(saved.get_fields()['VERIFICATION VESTED OWNER']['/V'],
                         'Example Property Holdings LLC')

    def render(self, **office):
        from retainer_draft import render_retainer_draft
        return render_retainer_draft({**GeneralRetainerTests().plan(), 'verification': {
            'capacity': 'plaintiff', 'printed_name': 'Maddy Test Owner', 'title': 'Owner',
            'execution_county': 'Santa Clara', 'payment_method': 'check',
            'source': 'Fictional owner fixture confirmed by Kwame', **office}})

    def test_complete_office_fields_are_saved_with_appearances(self):
        import io
        from pypdf import PdfReader
        result = self.render()
        self.assertTrue(result['office_fields_complete'])
        self.assertEqual(result['missing_office_fields'], [])
        self.assertEqual(result['fields']['PAPERCHECK'], 'X')
        self.assertEqual(result['fields']['CREDITCARD'], '')
        self.assertFalse(result['approved'])
        self.assertFalse(result['sent'])
        saved = PdfReader(io.BytesIO(base64.b64decode(result['pdf_base64'])))
        fields = saved.get_fields()
        for name, value in result['fields'].items():
            self.assertEqual(fields[name]['/V'], value)
        widgets = [ref.get_object() for ref in saved.pages[2]['/Annots']]
        for widget in widgets:
            self.assertTrue(widget['/AP']['/N'].get_object().get_data())
        self.assertEqual(fields['VERIFICATION PRINTED NAME']['/V'], 'Maddy Test Owner')
        self.assertEqual(fields['VERIFICATION COUNTY']['/V'], 'Santa Clara')
        self.assertFalse(any('SIGNATURE' in name or 'DATE' in name for name in fields))

    def test_each_supported_capacity_marks_only_its_own_lines(self):
        options = ['plaintiff', 'agent', 'llc_member', 'general_partner',
                   'limited_partner', 'trustee', 'successor_trustee']
        for capacity in options:
            with self.subTest(capacity=capacity):
                extra = {'capacity': capacity}
                if capacity == 'agent':
                    extra['plaintiff_name'] = 'Fictional Owner'
                elif capacity != 'plaintiff':
                    extra['entity_name'] = 'Fictional Entity'
                result = self.render(**extra)
                self.assertTrue(result['office_fields_complete'])
                marks = [k for k, v in result['fields'].items()
                         if k.startswith('VERIFICATION CAPACITY') and v]
                self.assertEqual(len(marks), 1)
                self.assertEqual(marks[0], 'VERIFICATION CAPACITY ' + (
                    'PARTNERSHIP' if capacity in ('general_partner', 'limited_partner') else capacity.upper()))
                entities = [v for k, v in result['fields'].items()
                            if k.startswith('VERIFICATION ENTITY') and v]
                self.assertEqual(len(entities), int(capacity not in ('agent', 'plaintiff')))

    def test_missing_office_data_never_claims_ready_for_signature(self):
        result = self.render(capacity='', printed_name='', title='', execution_county='', payment_method='')
        self.assertFalse(result['office_fields_complete'])
        self.assertEqual(set(result['missing_office_fields']),
                         {'capacity', 'printed_name', 'title', 'execution_county', 'payment_method'})
        result = self.render(capacity='agent')
        self.assertIn('plaintiff_name', result['missing_office_fields'])

    def test_bad_or_unsourced_values_are_rejected(self):
        for extra in [{'source': ''}, {'capacity': 'son'}, {'entity_name': 'Wrong Entity'},
                      {'printed_name': 'Name\nInjected'}, {'payment_method': 'wire'},
                      {'printed_name': 'Extremely Long Name ' * 30}, {'signed_date': 'today'}]:
            with self.subTest(extra=extra), self.assertRaises(ValueError):
                self.render(**extra)

    def test_original_page_contents_and_signature_lines_are_unchanged(self):
        import io
        from pypdf import PdfReader
        from retainer_draft import TEMPLATE
        source = PdfReader(TEMPLATE)
        saved = PdfReader(io.BytesIO(base64.b64decode(self.render()['pdf_base64'])))
        self.assertEqual(len(saved.pages), 3)
        for old, new in zip(source.pages, saved.pages):
            self.assertEqual(old.get_contents().get_data(), new.get_contents().get_data())


class CommonCountyFormTests(unittest.TestCase):
    def test_reviewed_counties_change_both_printed_references_and_canonical_field(self):
        import io
        from pypdf import PdfReader
        from retainer_draft import render_retainer_draft, SUPPORTED_COUNTIES
        for county in SUPPORTED_COUNTIES:
            with self.subTest(county=county):
                result = render_retainer_draft({**GeneralRetainerTests().plan(), 'county': county + ' County'})
                saved = PdfReader(io.BytesIO(base64.b64decode(result['pdf_base64'])))
                self.assertEqual(result['property_county'], county)
                self.assertEqual(saved.get_fields()['COUNTY OF']['/V'], county)
                text = saved.pages[0].extract_text()
                self.assertEqual(text.count(county + ' County'), 2)
                if county != 'Santa Clara':
                    self.assertNotIn('Santa Clara', text)
                self.assertEqual(len(saved.pages), 3)
                self.assertFalse(result['sent'])

    def test_new_county_text_keeps_source_font_and_other_pages(self):
        import io
        from pypdf import PdfReader
        from retainer_draft import render_retainer_draft, TEMPLATE
        source = PdfReader(TEMPLATE)
        result = render_retainer_draft({**GeneralRetainerTests().plan(), 'county': 'Santa Cruz'})
        saved = PdfReader(io.BytesIO(base64.b64decode(result['pdf_base64'])))
        for old, new in zip(source.pages[1:], saved.pages[1:]):
            self.assertEqual(old.get_contents().get_data(), new.get_contents().get_data())
        resources = saved.pages[0]['/Resources']['/XObject']
        changed = []
        for obj in resources.values():
            form = obj.get_object()
            res = form.get('/Resources')
            if not res:
                continue
            fonts = res.get_object().get('/Font')
            if fonts and '/RetainerCounty' in fonts.get_object():
                changed.append(form)
        self.assertEqual(len(changed), 2)
        for obj in changed:
            self.assertIn('Inter-Regular', obj['/Resources']['/Font']['/RetainerCounty']['/BaseFont'])

    def test_county_is_case_insensitive_but_never_guessed(self):
        from retainer_draft import render_retainer_draft
        result = render_retainer_draft({**GeneralRetainerTests().plan(), 'county': ' san mateo COUNTY '})
        self.assertEqual(result['property_county'], 'San Mateo')
        for value in ('', None, 'San Jose', 'Alameda', 'Narnia County'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                render_retainer_draft({**GeneralRetainerTests().plan(), 'county': value})
