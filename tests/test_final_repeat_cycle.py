import hashlib
import unittest
from signature_rollover import plan_completed_signature_rollover, RolloverHeld
class FinalRepeatTests(unittest.TestCase):
 def fixture(self):
  r='recxLwYEXtb2ZrWen';f='folder';s='b24ac860408c7f574b45ee6fc6f2f6310a366633';h='f64020b5de2c4dd6db53e3186a88f93d58d4c3c03f03eb79feacf4ebd6f2942c';pdf=b'%PDF-test';d=hashlib.sha256(pdf).hexdigest();c=hashlib.sha256(('owner-signature-case-v1\n'+r).encode()).hexdigest()
  names=['signature-approval.json','signed-package.json','Maddy signature approval publication','Maddy signature send - '+c]
  return dict(request=dict(record_id=r,approved_document_sha256=h,test_mode=True,authorized_test_cycle='narnia-final-receipt-20261016'),case=dict(id=r,fields={'Case Folder ID':f,'Status':'To Owner for Review (N)'}),approval=dict(record_id=r,folder_id=f,signature_request_id=s,approved_document_sha256=h,approved=True),package=dict(record_id=r,folder_id=f,signature_request_id=s,signature_status='complete',test_mode=True,file_id='file',sha256=d),receipt=dict(record_id=r,case_folder_id=f,signature_request_id=s,sha256=d,test_mode=True,delivery_verified=True,sent_message_id='sent',dispatch_id='dispatch'),provider_request=dict(signature_request_id=s,is_complete=True,is_declined=False,has_error=False,test_mode=True,metadata=dict(record_id=r,approved_document_sha256=h)),signed_file=dict(id='file',parentReference=dict(id=f),file={'mimeType':'application/pdf'},size=len(pdf)),signed_pdf=pdf,controls=[dict(id=str(i),name=n,eTag='etag',parentReference=dict(id=f)) for i,n in enumerate(names)])
 def test_authorized_repeat_preserves_controls_and_same_lock(self):
  f=self.fixture();p=plan_completed_signature_rollover(**f);self.assertEqual(len(p['preserve_by_rename']),3);self.assertFalse(p['delete_files']);self.assertFalse(p['send_signature']);f['request']['approved_document_sha256']='a'*64;self.assertEqual(p['exclusive_transition_name'],plan_completed_signature_rollover(**f)['exclusive_transition_name'])
 def test_unapproved_or_incomplete_repeats_rejected(self):
  for edit in [lambda f:f['request'].pop('authorized_test_cycle'),lambda f:f['request'].update(authorized_test_cycle='other'),lambda f:f['request'].update(test_mode=False),lambda f:f['approval'].update(signature_request_id='c'*40),lambda f:f['receipt'].update(delivery_verified=False),lambda f:f['provider_request'].update(is_complete=False),lambda f:f.update(signed_pdf=b'%PDF-changed')]:
   f=self.fixture();edit(f)
   with self.assertRaises(RolloverHeld):plan_completed_signature_rollover(**f)
