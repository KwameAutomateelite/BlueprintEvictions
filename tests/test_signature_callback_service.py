from pathlib import Path
import sys,unittest,hmac,hashlib
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from unittest.mock import AsyncMock,patch
from signature_callback_service import process_signature_callback
class Response:
 def __init__(self,data=None,content=b'%PDF-test'):self.data=data;self.content=content
 def raise_for_status(self):pass
 def json(self):return self.data
class Tests(unittest.IsolatedAsyncioTestCase):
 def setup_case(self,status='Sent for Signature'):
  kind='signature_request_downloadable';key='test';req={'signature_request_id':'sig','is_complete':True,'test_mode':True,'metadata':{'record_id':'recTest','approved_document_sha256':'a'*64}}
  p={'signature_request':req,'event':{'event_type':kind,'event_time':1,'event_hash':hmac.new(key.encode(),('1'+kind).encode(),hashlib.sha256).hexdigest()}}
  provider=AsyncMock();provider.get.side_effect=[Response({'signature_request':req}),Response()];airtable=AsyncMock();airtable.get.side_effect=[Response({'id':'recTest','fields':{'Case Folder ID':'folder','Status':status}}),Response({'id':'recTest','fields':{'Status':'With Process Server (N)'}})];airtable.patch.return_value=Response()
  args=dict(provider=provider,airtable=airtable,storage_http=AsyncMock(),api_key=key,bridge_url='https://blueprintevictions.app.n8n.cloud/webhook/test',bridge_secret='test',base_id='base',table_id='table',allowed_records={'recTest'})
  stored={'stored':True,'manifest':{'folder_id':'folder'},'expected_status':'Sent for Signature'};return p,args,stored
 async def test_saved_before_status_write(self):
  p,a,s=self.setup_case()
  async def store(*args,**kwargs):a['airtable'].patch.assert_not_called();return s
  with patch('signature_callback_service.store_signed_package',side_effect=store):r=await process_signature_callback(p,**a)
  self.assertTrue(r['status_verified']);self.assertEqual(a['airtable'].patch.await_count,1)
 async def test_storage_failure_no_status_write(self):
  p,a,s=self.setup_case()
  with patch('signature_callback_service.store_signed_package',side_effect=ValueError('Storage mismatch')):
   with self.assertRaises(ValueError):await process_signature_callback(p,**a)
  a['airtable'].patch.assert_not_called()
 async def test_no_regression(self):
  p,a,s=self.setup_case('Served (A)')
  with patch('signature_callback_service.store_signed_package',return_value=s):
   with self.assertRaises(ValueError):await process_signature_callback(p,**a)
  a['airtable'].patch.assert_not_called()
 async def test_duplicate_signed_no_patch(self):
  p,a,s=self.setup_case('With Process Server (N)')
  with patch('signature_callback_service.store_signed_package',return_value=s):await process_signature_callback(p,**a)
  a['airtable'].patch.assert_not_called()
 async def test_advanced_status_cannot_be_approved_baseline(self):
  for status in ['Served (A)','Expired (A)','Closed - Archive (N)','Cancelled','Hold (N)']:
   p,a,s=self.setup_case(status);s['expected_status']=status
   with patch('signature_callback_service.store_signed_package',return_value=s):
    with self.assertRaises(ValueError):await process_signature_callback(p,**a)
   a['airtable'].patch.assert_not_called()
 async def test_wrong_scope_no_storage(self):
  p,a,s=self.setup_case();a['allowed_records']=set()
  with patch('signature_callback_service.store_signed_package') as store:
   with self.assertRaises(ValueError):await process_signature_callback(p,**a)
   store.assert_not_called()
 async def test_new_signed_notice_clears_previous_dispatch_date(self):
  p,a,s=self.setup_case()
  a['airtable'].get.side_effect=[Response({'id':'recTest','fields':{'Case Folder ID':'folder','Status':'Sent for Signature','Dispatched On':'2026-09-08'}}),Response({'id':'recTest','fields':{'Status':'With Process Server (N)'}})]
  with patch('signature_callback_service.store_signed_package',return_value=s):await process_signature_callback(p,**a)
  self.assertEqual(a['airtable'].patch.call_args.kwargs['json']['fields'],{'Status':'With Process Server (N)','Dispatched On':None})
 async def test_uncleared_dispatch_date_not_reported_as_queued(self):
  p,a,s=self.setup_case()
  a['airtable'].get.side_effect=[Response({'id':'recTest','fields':{'Case Folder ID':'folder','Status':'Sent for Signature'}}),Response({'id':'recTest','fields':{'Status':'With Process Server (N)','Dispatched On':'2026-09-08'}})]
  with patch('signature_callback_service.store_signed_package',return_value=s):
   with self.assertRaisesRegex(ValueError,'Dispatch queue'):await process_signature_callback(p,**a)
 async def test_duplicate_callback_preserves_completed_dispatch_date(self):
  p,a,s=self.setup_case('With Process Server (N)')
  a['airtable'].get.side_effect=[Response({'id':'recTest','fields':{'Case Folder ID':'folder','Status':'With Process Server (N)','Dispatched On':'2026-09-14'}})]
  with patch('signature_callback_service.store_signed_package',return_value=s):r=await process_signature_callback(p,**a)
  self.assertTrue(r['status_verified']);a['airtable'].patch.assert_not_called()
if __name__=='__main__':unittest.main()
