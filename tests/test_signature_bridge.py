from pathlib import Path
import sys,unittest,hashlib
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from unittest.mock import AsyncMock
from signature_bridge import store_signed_package, record_signature_approval, claim_signature_send, SignatureClaimExists
class Response:
 def __init__(self,result,status=200):self.result=result;self.status=status
 def raise_for_status(self):
  if self.status!=200:raise RuntimeError('Storage failed')
 def json(self):return self.result
class Tests(unittest.IsolatedAsyncioTestCase):
 def setup_case(self):
  c=AsyncMock();pdf=b'%PDF-controlled-test';m={'record_id':'recTest','signature_request_id':'sig','sha256':hashlib.sha256(pdf).hexdigest(),'test_mode':True,'signature_status':'complete','file_id':'file','folder_id':'folder'};r={'stored':True,'manifest':m};c.post.return_value=Response(r)
  args={'record_id':'recTest','request_id':'sig','approved_sha256':'a'*64,'pdf':pdf,'test_mode':True};return c,r,args
 async def test_valid(self):
  c,r,a=self.setup_case();self.assertEqual(await store_signed_package(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','test-secret',**a),r);self.assertEqual(c.post.call_args.kwargs['headers']['X-Maddy-Signature-Key'],'test-secret')
 async def test_endpoint_guard(self):
  for url in ['http://blueprintevictions.app.n8n.cloud/webhook/test','https://other.example/webhook/test','https://blueprintevictions.app.n8n.cloud/api/v1/workflows']:
   c,r,a=self.setup_case()
   with self.assertRaises(ValueError):await store_signed_package(c,url,'test-secret',**a)
   c.post.assert_not_called()
 async def test_reject_mismatched_receipt(self):
  for field,value in [('record_id','other'),('signature_request_id','other'),('sha256','wrong'),('test_mode',False),('file_id',''),('folder_id','')]:
   c,r,a=self.setup_case();r['manifest'][field]=value
   with self.assertRaises(ValueError):await store_signed_package(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','test-secret',**a)
 async def test_http_failure_not_retried(self):
  c,r,a=self.setup_case();c.post.return_value=Response({},500)
  with self.assertRaises(RuntimeError):await store_signed_package(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','test-secret',**a)
  self.assertEqual(c.post.await_count,1)
 async def test_only_bound_explicit_duplicate_can_enter_rollover(self):
  c=AsyncMock();r={'claimed':False,'record_id':'recTest','approved_document_sha256':'a'*64,'signer_email':'kwame@automateelite.com','folder_id':'folder','expected_status':'To Owner for Review (N)','claim_folder_id':None}
  c.post.return_value=Response(r)
  with self.assertRaises(SignatureClaimExists):await claim_signature_send(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','secret','recTest','a'*64,'kwame@automateelite.com')
  for field,value in [('record_id','wrong'),('claimed',None),('claim_folder_id','unexpected')]:
   c.post.return_value=Response({**r,field:value})
   with self.assertRaises(ValueError) as error:await claim_signature_send(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','secret','recTest','a'*64,'kwame@automateelite.com')
   self.assertNotIsInstance(error.exception,SignatureClaimExists)
 async def test_approval_recording(self):
  c=AsyncMock();a={'record_id':'recTest','signature_request_id':'sig','folder_id':'folder','expected_status':'Waiting','approved_document_sha256':'a'*64};c.post.return_value=Response({'recorded':True,'approval':{**a,'approved':True}})
  result=await record_signature_approval(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','secret',a);self.assertTrue(result['recorded'])
 async def test_approval_mismatch_not_retried(self):
  c=AsyncMock();c.post.return_value=Response({'recorded':True,'approval':{}})
  with self.assertRaises(ValueError):await record_signature_approval(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','secret',{'record_id':'recTest'})
  self.assertEqual(c.post.await_count,1)
 async def test_signature_send_claim(self):
  c=AsyncMock();c.post.return_value=Response({'claimed':True,'record_id':'recTest','approved_document_sha256':'a'*64,'signer_email':'kwame@automateelite.com','folder_id':'folder','expected_status':'Waiting','claim_folder_id':'claim'})
  result=await claim_signature_send(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','secret','recTest','a'*64,'kwame@automateelite.com');self.assertTrue(result['claimed'])
 async def test_duplicate_signature_claim_holds(self):
  c=AsyncMock();c.post.return_value=Response({'claimed':False})
  with self.assertRaises(ValueError):await claim_signature_send(c,'https://blueprintevictions.app.n8n.cloud/webhook/test','secret','recTest','a'*64,'kwame@automateelite.com')
  self.assertEqual(c.post.await_count,1)
if __name__=='__main__':unittest.main()
