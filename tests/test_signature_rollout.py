import json,unittest
from pathlib import Path
from signature_rollout import parse_rollout
class RolloutTests(unittest.TestCase):
 def setUp(self):self.d=json.loads(Path('signature-rollout.json').read_text())
 def test_current_scope(self):
  p=parse_rollout(json.dumps(self.d))
  for rid in self.d['allowed_case_ids']:self.assertTrue(p.includes(rid))
  self.assertFalse(p.includes('recABCDEFGHIJKLMN'));self.assertTrue(p.accepts_signer('KWAME@AUTOMATEELITE.COM',{}));self.assertFalse(p.accepts_signer('am@blueprintevictions.com',{}));self.assertTrue(p.test_mode)
 def test_production_matches_case(self):
  self.d.update(profile='anne-michelle',allowed_case_ids=None,allowed_signers=None,test_mode=False);p=parse_rollout(json.dumps(self.d));self.assertTrue(p.includes('recABCDEFGHIJKLMN'))
  f={"Landlord's Preferred Email":'Client@example.com','Email':'someone@example.com'}
  self.assertTrue(p.accepts_signer('client@example.com',f));self.assertFalse(p.accepts_signer('someone@example.com',f));self.assertFalse(p.accepts_signer('client@example.com',{}));self.assertTrue(p.accepts_signer('client@example.com',{'Email':'client@example.com'}));self.assertFalse(p.includes('../records'))
 def test_test_profile_cannot_broaden(self):
  for change in ({'allowed_case_ids':None},{'allowed_signers':None},{'test_mode':False}):
   with self.subTest(change=change),self.assertRaises(ValueError):parse_rollout(json.dumps({**self.d,**change}))
 def test_invalid_configuration(self):
  for raw in ('{}','null','[]','{"profile":"anne-michelle"}',json.dumps({**self.d,'test_mode':'false'}),json.dumps({**self.d,'allowed_case_ids':['bad']})):
   with self.subTest(raw=raw),self.assertRaises(ValueError):parse_rollout(raw)

class CoexistingRolloutTests(unittest.TestCase):
 def test_test_cases_remain_test_mode_after_production_cutover(self):
  from unittest.mock import patch
  from signature_rollout import load_rollout
  test=json.loads(Path('signature-rollout.json').read_text())
  prod={**test,'profile':'anne-michelle','allowed_case_ids':None,'allowed_signers':None,'test_mode':False}
  with patch.dict('os.environ',{'SIGNATURE_ROLLOUT_POLICY':json.dumps(prod),'SIGNATURE_TEST_ROLLOUT_POLICY':json.dumps(test)}):
   self.assertTrue(load_rollout(test['allowed_case_ids'][0]).test_mode)
   self.assertFalse(load_rollout('recABCDEFGHIJKLMN').test_mode)
   self.assertFalse(load_rollout('recABCDEFGHIJKLMN').accepts_signer('kwame@automateelite.com',{'Email':'client@example.com'}))
