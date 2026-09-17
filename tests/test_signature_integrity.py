from pathlib import Path
import sys,pathlib,hashlib,hmac,unittest
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from signature_integrity import verify_event,verified_request_record
class Integrity(unittest.TestCase):
 def test_authentic_event(self):
  e={'event_time':'100','event_type':'signature_request_all_signed'}
  e['event_hash']=hmac.new(b'test-secret',b'100signature_request_all_signed',hashlib.sha256).hexdigest()
  self.assertTrue(verify_event({'event':e},'test-secret'))
  self.assertFalse(verify_event({'event':e},'wrong-secret'))
  e['event_type']='signature_request_downloadable'
  self.assertFalse(verify_event({'event':e},'test-secret'))
 def test_missing_event(self):
  for p in [{},{'event':None},{'event':{'event_time':True}}]:self.assertFalse(verify_event(p,'key'))
 def test_provider_binding(self):
  callback={'signature_request_id':'sig-test','metadata':{'record_id':'rec-test'}}
  provider={**callback,'is_complete':True}
  self.assertEqual(verified_request_record(callback,provider),'rec-test')
  for update in [{'is_complete':False},{'signature_request_id':'other'},{'metadata':{'record_id':'rec-other'}},{'is_declined':True},{'has_error':True}]:
   with self.assertRaises(ValueError):verified_request_record(callback,{**provider,**update})
if __name__=='__main__':unittest.main()
