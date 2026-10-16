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
        for key, value in [('case_folder_id', ''), ('county', 'San Mateo'),
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
