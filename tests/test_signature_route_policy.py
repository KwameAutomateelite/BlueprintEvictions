"""Exercise the actual send endpoint body with external I/O replaced by fakes."""
import ast,hashlib,json,logging,os,tempfile,unittest
from pathlib import Path
from types import SimpleNamespace as N
from unittest.mock import patch,AsyncMock

class Hold(Exception):
 def __init__(self,status_code,detail):self.status_code=status_code;self.detail=detail
class RouteTests(unittest.IsolatedAsyncioTestCase):
 async def asyncSetUp(self):
  self.tmp=tempfile.TemporaryDirectory();self.file=Path(self.tmp.name)/'test.pdf';self.file.write_bytes(b'%PDF-test')
  self.sent=[];self.case={'id':'recxLwYEXtb2ZrWen','fields':{'Case Folder ID':'folder','Status':'To Owner for Review (N)',"Landlord's Preferred Email":'client@example.com'}}
  me=self
  class Client:
   def __init__(self,*a,**k):pass
   async def __aenter__(self):return self
   async def __aexit__(self,*a):pass
   async def get(self,*a,**k):return N(raise_for_status=lambda:None,json=lambda:me.case)
  class APIClient:
   def __init__(self,*a,**k):pass
   def __enter__(self):return self
   def __exit__(self,*a):pass
  def provider_send(data):
   self.sent.append(data);data.files[0].close();return N(signature_request=N(signature_request_id='test-request'))
  self.g={'SendSignatureRequest':N,'SendSignatureResponse':N,'logger':logging.getLogger('test'),'os':os,'Path':Path,'hashlib':hashlib,'HTTPException':Hold,'httpx':N(AsyncClient=Client,HTTPError=RuntimeError),'AIRTABLE_API_KEY':'fixture','AIRTABLE_BASE_ID':'base','AIRTABLE_TABLE_ID':'table','ApiClient':APIClient,'configuration':None,'ApiException':type('FakeAPIError',(Exception,),{}),'apis':N(SignatureRequestApi=lambda _:N(signature_request_send=provider_send)),'models':N(SubSignatureRequestSigner=N,SubSigningOptions=N,SignatureRequestSendRequest=N),'download_file':AsyncMock(return_value=str(self.file))}
  tree=ast.parse(Path('main.py').read_text());f=next(x for x in tree.body if isinstance(x,ast.AsyncFunctionDef) and x.name=='send_signature');f.decorator_list=[];exec(compile(ast.Module(body=[f],type_ignores=[]),'main.py','exec'),self.g)
  self.req=N(require_verified_signature=True,file_url='https://example.invalid/test.pdf',fields={},attachments_required=[],record_id=self.case['id'],signer_email='kwame@automateelite.com',signer_name='Test',document_name='Test',notice_type='Test',case_name='Test')
  self.env={k:'fixture' for k in ['SIGNATURE_APPROVAL_BRIDGE_URL','SIGNATURE_STORAGE_BRIDGE_URL','SIGNATURE_STORAGE_BRIDGE_SECRET','SIGNATURE_SEND_CLAIM_URL']};self.env['DROPBOX_SIGN_TEST_MODE']='0'
  self.claim=AsyncMock(return_value={'folder_id':'folder','expected_status':'To Owner for Review (N)'})
 async def asyncTearDown(self):self.tmp.cleanup()
 async def run_route(self):
  with patch.dict(os.environ,self.env,clear=True),patch('signature_bridge.claim_signature_send',self.claim),patch('signature_bridge.record_signature_approval',AsyncMock()):return await self.g['send_signature'](self.req)
 async def test_test_profile_forces_provider_test_mode(self):
  await self.run_route();self.assertEqual(len(self.sent),1);self.assertIs(self.sent[0].test_mode,True);self.claim.assert_awaited_once()
 async def test_production_matches_case_and_allows_production_mode(self):
  self.env['SIGNATURE_ROLLOUT_POLICY']=json.dumps({'profile':'anne-michelle','allowed_case_ids':None,'allowed_signers':None,'signer_fields':["Landlord's Preferred Email",'Email'],'test_mode':False});self.req.signer_email='client@example.com';await self.run_route();self.assertIs(self.sent[0].test_mode,False)
 async def test_production_wrong_signer_never_calls_provider(self):
  self.env['SIGNATURE_ROLLOUT_POLICY']=json.dumps({'profile':'anne-michelle','allowed_case_ids':None,'allowed_signers':None,'signer_fields':["Landlord's Preferred Email",'Email'],'test_mode':False})
  with self.assertRaises(Hold) as c:await self.run_route()
  self.assertEqual(c.exception.status_code,403);self.assertFalse(self.sent);self.claim.assert_not_awaited()
 async def test_required_verified_path_cannot_fall_back_to_legacy(self):
  self.req.record_id='recABCDEFGHIJKLMN'
  with self.assertRaises(Hold) as c:await self.run_route()
  self.assertEqual(c.exception.status_code,409);self.assertFalse(self.sent)
 async def test_missing_bridge_never_calls_provider(self):
  del self.env['SIGNATURE_APPROVAL_BRIDGE_URL']
  with self.assertRaises(Hold) as c:await self.run_route()
  self.assertEqual(c.exception.status_code,503);self.assertFalse(self.sent)
 async def test_unselected_legacy_call_remains_unchanged(self):
  self.req.record_id='recABCDEFGHIJKLMN';self.req.require_verified_signature=False
  await self.run_route();self.assertEqual(len(self.sent),1);self.assertIs(self.sent[0].test_mode,False);self.claim.assert_not_awaited()

 async def test_exact_approved_pdf_is_sent(self):
  self.req.approved_pdf_sha256=hashlib.sha256(self.file.read_bytes()).hexdigest()
  await self.run_route();self.assertEqual(len(self.sent),1)
 async def test_changed_pdf_never_claims_or_sends(self):
  self.req.approved_pdf_sha256='0'*64
  with self.assertRaises(Hold) as c:await self.run_route()
  self.assertEqual(c.exception.status_code,409);self.assertFalse(self.sent);self.claim.assert_not_awaited()
 async def test_approved_pdf_cannot_regenerate_from_old_fields(self):
  self.req.approved_pdf_sha256='0'*64;self.req.file_url=None
  with self.assertRaises(Hold) as c:await self.run_route()
  self.assertEqual(c.exception.status_code,400);self.assertFalse(self.sent);self.claim.assert_not_awaited()

 async def test_repeat_cycle_disabled_by_default(self):
  from signature_bridge import SignatureClaimExists
  self.claim.side_effect=SignatureClaimExists('duplicate')
  with patch('signature_rollover_service.claim_completed_cycle',AsyncMock()) as repeat:
   with self.assertRaises(Hold):await self.run_route()
   repeat.assert_not_awaited();self.assertFalse(self.sent)

 async def test_repeat_cycle_uses_exact_approved_pdf_and_one_provider_send(self):
  from signature_bridge import SignatureClaimExists
  self.env['SIGNATURE_ROLLOVER_TEST_ENABLED']='1'
  self.req.approved_pdf_sha256=hashlib.sha256(self.file.read_bytes()).hexdigest()
  self.claim.side_effect=SignatureClaimExists('duplicate')
  with patch('signature_rollover_service.claim_completed_cycle',AsyncMock(return_value={'folder_id':'folder','expected_status':'To Owner for Review (N)','repeat_cycle_reservation':'fixture-transition'})) as repeat:
   await self.run_route();repeat.assert_awaited_once()
   self.assertEqual(repeat.call_args.kwargs['request']['approved_document_sha256'],self.req.approved_pdf_sha256)
   self.assertIs(repeat.call_args.kwargs['request']['test_mode'],True)
   self.assertEqual(len(self.sent),1);self.assertIs(self.sent[0].test_mode,True)

 async def test_uncertain_claim_never_enters_repeat_path(self):
  self.env['SIGNATURE_ROLLOVER_TEST_ENABLED']='1'
  self.req.approved_pdf_sha256=hashlib.sha256(self.file.read_bytes()).hexdigest()
  self.claim.side_effect=ValueError('unknown outcome')
  with patch('signature_rollover_service.claim_completed_cycle',AsyncMock()) as repeat:
   with self.assertRaises(Hold):await self.run_route()
   repeat.assert_not_awaited();self.assertFalse(self.sent)

 async def test_failed_rollover_never_sends_signature(self):
  from signature_bridge import SignatureClaimExists
  self.env['SIGNATURE_ROLLOVER_TEST_ENABLED']='1'
  self.req.approved_pdf_sha256=hashlib.sha256(self.file.read_bytes()).hexdigest()
  self.claim.side_effect=SignatureClaimExists('duplicate')
  with patch('signature_rollover_service.claim_completed_cycle',AsyncMock(side_effect=ValueError('incomplete prior cycle'))):
   with self.assertRaises(Hold):await self.run_route()
   self.assertFalse(self.sent)

 async def test_repeat_cycle_requires_exact_approved_hash(self):
  from signature_bridge import SignatureClaimExists
  self.env['SIGNATURE_ROLLOVER_TEST_ENABLED']='1';self.claim.side_effect=SignatureClaimExists('duplicate')
  with patch('signature_rollover_service.claim_completed_cycle',AsyncMock()) as repeat:
   with self.assertRaises(Hold):await self.run_route()
   repeat.assert_not_awaited();self.assertFalse(self.sent)

 async def test_case_changed_during_rollover_never_sends(self):
  from signature_bridge import SignatureClaimExists
  self.env['SIGNATURE_ROLLOVER_TEST_ENABLED']='1'
  self.req.approved_pdf_sha256=hashlib.sha256(self.file.read_bytes()).hexdigest()
  self.claim.side_effect=SignatureClaimExists('duplicate')
  async def changed(*a,**kw):
   self.case['fields']['Status']='Closed'
   return {'folder_id':'folder','expected_status':'To Owner for Review (N)','repeat_cycle_reservation':'fixture-transition'}
  with patch('signature_rollover_service.claim_completed_cycle',changed):
   with self.assertRaises(Hold):await self.run_route()
   self.assertFalse(self.sent)
