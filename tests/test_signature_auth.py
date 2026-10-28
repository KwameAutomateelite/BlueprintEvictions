import unittest
from unittest.mock import patch
from fastapi.testclient import TestClient
with patch.dict('os.environ', {'DROPBOX_SIGN_API_KEY': 'test-only', 'AIRTABLE_API_KEY': 'test-only'}):
    import main

class SignatureAuthenticationTests(unittest.TestCase):
 def test_unsigned_http_request_never_reaches_signature_action(self):
  with patch.dict('os.environ',{'SIGNATURE_STORAGE_BRIDGE_SECRET':'auth-test-only'}):
   with TestClient(main.app) as client:
    self.assertEqual(client.post('/send-signature',json={}).status_code,401)
    self.assertEqual(client.post('/send-signature',json={},headers={'x-maddy-signature-key':'wrong'}).status_code,401)
    self.assertEqual(client.post('/send-signature',json={},headers={'x-maddy-signature-key':'auth-test-only'}).status_code,422)
 def test_missing_server_secret_does_not_allow_empty_header(self):
  with patch.dict('os.environ',{'SIGNATURE_STORAGE_BRIDGE_SECRET':''}):
   with TestClient(main.app) as client:self.assertEqual(client.post('/send-signature',json={}).status_code,401)
