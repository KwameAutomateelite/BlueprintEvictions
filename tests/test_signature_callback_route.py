"""HTTP acknowledgement contract for real account events and completion retries."""
import hashlib
import hmac
import json
import os
import unittest
from unittest.mock import AsyncMock, patch

import httpx
from fastapi.testclient import TestClient

with patch.dict(os.environ, {'DROPBOX_SIGN_API_KEY': 'test-only', 'AIRTABLE_API_KEY': 'test-only'}):
    import main


class CallbackRouteTests(unittest.TestCase):
    def setUp(self):
        self.key = 'callback-test-key'
        self.env = {'SIGNATURE_STORAGE_BRIDGE_URL': 'https://example.invalid/storage',
                    'SIGNATURE_STORAGE_BRIDGE_SECRET': 'test-only',
                    'SIGNATURE_TEST_ROLLOUT_POLICY': '',
                    'SIGNATURE_ROLLOUT_POLICY': json.dumps({
                        'profile': 'anne-michelle', 'allowed_case_ids': None,
                        'allowed_signers': None, 'signer_fields': ['Email'], 'test_mode': False})}
        self.client = TestClient(main.app)

    def event(self, kind, record=None):
        event = {'event_type': kind, 'event_time': '1794942000'}
        event['event_hash'] = hmac.new(self.key.encode(),
            (event['event_time'] + kind).encode(), hashlib.sha256).hexdigest()
        payload = {'event': event}
        if record:
            payload['signature_request'] = {'signature_request_id': 'synthetic-only',
                                           'metadata': {'record_id': record}}
        return payload

    def send(self, payload):
        with patch.object(main, 'DROPBOX_SIGN_API_KEY', self.key), patch.dict(os.environ, self.env):
            return self.client.post('/signature-callback', data={'json': json.dumps(payload)})

    def assert_ack(self, response):
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.text, 'Hello API Event Received')
        self.assertTrue(response.headers['content-type'].startswith('text/plain'))

    def test_account_events_without_case_metadata_ack_without_business_io(self):
        with patch.object(main, 'update_airtable_status', AsyncMock()) as update, \
             patch('signature_callback_service.process_signature_callback', AsyncMock()) as process:
            for kind in ['callback_test', 'signature_request_sent', 'signature_request_viewed',
                         'signature_request_signed', 'signature_request_canceled',
                         'signature_request_all_signed', 'signature_request_downloadable',
                         'account_confirmed']:
                with self.subTest(kind=kind):
                    self.assert_ack(self.send(self.event(kind)))
            update.assert_not_awaited()
            process.assert_not_awaited()

    def test_ignored_event_does_not_depend_on_rollout_configuration(self):
        self.env['SIGNATURE_ROLLOUT_POLICY'] = 'invalid'
        self.assert_ack(self.send(self.event('signature_request_sent', 'recABCDEFGHIJKLMN')))

    def test_forged_events_never_acknowledge_or_change_cases(self):
        with patch.object(main, 'update_airtable_status', AsyncMock()) as update, \
             patch('signature_callback_service.process_signature_callback', AsyncMock()) as process:
            for kind in ['callback_test', 'signature_request_sent', 'signature_request_all_signed']:
                payload = self.event(kind, 'recABCDEFGHIJKLMN')
                payload['event']['event_hash'] = '0' * 64
                self.assertEqual(self.send(payload).status_code, 409)
            update.assert_not_awaited()
            process.assert_not_awaited()

    def test_invalid_payloads_are_rejected(self):
        for data in [{}, {'json': '{'}, {'json': '[]'}, {'json': 'null'}]:
            with self.subTest(data=data):
                self.assertEqual(self.client.post('/signature-callback', data=data).status_code, 400)

    def test_verified_completion_acknowledged_after_storage_service_finishes(self):
        with patch('signature_callback_service.process_signature_callback', AsyncMock(return_value={'stored': True})) as process:
            self.assert_ack(self.send(self.event('signature_request_downloadable', 'recABCDEFGHIJKLMN')))
            process.assert_awaited_once()
            self.assertEqual(process.call_args.kwargs['allowed_records'], {'recABCDEFGHIJKLMN'})

    def test_storage_failure_keeps_provider_retry_available(self):
        for error, status in [(ValueError('identity changed'), 409),
                              (httpx.ConnectError('unavailable'), 503),
                              (TimeoutError(), 503)]:
            with self.subTest(error=type(error).__name__), \
                 patch('signature_callback_service.process_signature_callback', AsyncMock(side_effect=error)):
                response = self.send(self.event('signature_request_downloadable', 'recABCDEFGHIJKLMN'))
                self.assertEqual(response.status_code, status)
                self.assertNotIn('Hello API Event Received', response.text)

    def test_legacy_completion_also_returns_required_acknowledgement(self):
        self.env['SIGNATURE_STORAGE_BRIDGE_URL'] = ''
        with patch.object(main, 'update_airtable_status', AsyncMock()) as update:
            self.assert_ack(self.send(self.event('signature_request_all_signed', 'recABCDEFGHIJKLMN')))
            update.assert_awaited_once_with('recABCDEFGHIJKLMN', 'Signed (A)')


if __name__ == '__main__':
    unittest.main()
