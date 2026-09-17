"""Narrow authenticated n8n transport for the isolated Maddy rollover test."""
import json
from urllib.parse import urlparse

from signature_rollover import RolloverHeld


class GraphResult:
    def __init__(self, status_code, body):
        self.status_code, self.body = status_code, body

    def json(self):
        return self.body


def verify_bridge_url(url):
    parsed = urlparse(url)
    if (parsed.scheme != 'https' or parsed.hostname != 'blueprintevictions.app.n8n.cloud'
            or not parsed.path.startswith('/webhook/') or parsed.username
            or parsed.password or parsed.query or parsed.fragment or parsed.port):
        raise RolloverHeld('invalid_rollover_bridge')


class RolloverGraphTransport:
    def __init__(self, client, *, url, secret, record_id, test_mode):
        verify_bridge_url(url)
        if not secret or record_id != 'recxLwYEXtb2ZrWen' or test_mode is not True:
            raise RolloverHeld('rollover_transport_outside_test_scope')
        self.client, self.url, self.secret = client, url, secret
        self.record_id = record_id

    async def request(self, method, url, **kwargs):
        if not url.startswith('https://graph.microsoft.com/v1.0/me/drive/items/'):
            raise RolloverHeld('invalid_graph_destination')
        if set(kwargs) - {'headers', 'json', 'content', 'follow_redirects'}:
            raise RolloverHeld('unsupported_graph_arguments')
        headers = kwargs.get('headers') or {}
        if set(headers) - {'If-Match', 'Content-Type'}:
            raise RolloverHeld('credential_forwarding_forbidden')
        payload = kwargs.get('json')
        if 'content' in kwargs:
            if payload is not None:
                raise RolloverHeld('multiple_graph_payloads')
            payload = json.loads(kwargs['content'])
        response = await self.client.post(
            self.url, headers={'X-Maddy-Signature-Key': self.secret},
            json={'record_id': self.record_id, 'test_mode': True,
                  'method': method, 'url': url, 'payload': payload,
                  'if_match': headers.get('If-Match')})
        response.raise_for_status()
        result = response.json()
        if (type(result.get('status_code')) is not int
                or not 100 <= result['status_code'] <= 599 or 'body' not in result):
            raise RolloverHeld('rollover_bridge_response_invalid')
        return GraphResult(result['status_code'], result['body'])
