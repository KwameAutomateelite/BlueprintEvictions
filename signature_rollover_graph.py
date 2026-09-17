"""Graph storage for the guarded rollover executor.

Uses an injected authenticated HTTP client; does not acquire/export credentials.
Reservation data is immutable. Its filename represents the phase, allowing the
commit to use Graph's documented PATCH If-Match rather than an assumed atomic
read/replace of JSON content. No live signing route imports this module yet.
"""
import json
from urllib.parse import quote

from signature_rollover import RolloverHeld


class GraphRolloverStore:
    BASE = 'https://graph.microsoft.com/v1.0/me/drive/items/'

    def __init__(self, client, plan):
        self.client = client
        self.plan = plan
        self.reservations = {}

    async def _json(self, method, suffix, *, statuses=(200,), **kwargs):
        response = await self.client.request(method, self.BASE + suffix,
                                             follow_redirects=True, **kwargs)
        if response.status_code not in statuses:
            raise RolloverHeld('graph_operation_failed_' + str(response.status_code))
        return response.status_code, response.json()

    @staticmethod
    def _item(value):
        return quote(value, safe='')

    def _identity(self, item, *, name, parent, item_id=None):
        if (not item.get('id') or not item.get('eTag')
                or item.get('name') != name
                or item.get('parentReference', {}).get('id') != parent
                or (item_id and item['id'] != item_id)):
            raise RolloverHeld('graph_item_identity_unverified')

    async def create_exclusive(self, key, state):
        expected = {'record_id': self.plan['record_id'],
                    'folder_id': self.plan['folder_id'],
                    'prior_request_id': self.plan['prior_request_id'],
                    'new_approved_sha256': self.plan['new_approved_sha256']}
        if (key != self.plan['exclusive_transition_name']
                or any(state.get(k) != v for k, v in expected.items())
                or state.get('phase') != 'preserving' or not state.get('owner_token')):
            raise RolloverHeld('invalid_transition_state')
        status, folder = await self._json(
            'POST', self._item(self.plan['folder_id']) + '/children',
            statuses=(201, 409), json={'name': key, 'folder': {},
                                     '@microsoft.graph.conflictBehavior': 'fail'})
        if status == 409:
            if folder.get('error', {}).get('code') != 'nameAlreadyExists':
                raise RolloverHeld('uncertain_transition_conflict')
            return None
        self._identity(folder, name=key, parent=self.plan['folder_id'])
        if 'folder' not in folder:
            raise RolloverHeld('transition_is_not_folder')
        immutable = {k: v for k, v in state.items() if k != 'phase'}
        _, item = await self._json(
            'PUT', self._item(folder['id']) + ':/preserving.json:/content?@microsoft.graph.conflictBehavior=fail',
            statuses=(201,), headers={'Content-Type': 'application/json'},
            content=json.dumps(immutable).encode())
        self._identity(item, name='preserving.json', parent=folder['id'])
        _, saved = await self._json('GET', self._item(item['id']) + '/content')
        if saved != immutable:
            raise RolloverHeld('transition_state_readback_mismatch')
        self.reservations[key] = {'item_id': item['id'], 'folder_id': folder['id'],
                                  'state': immutable, 'etag': item['eTag']}
        return {'state': {**saved, 'phase': 'preserving'}, 'etag': item['eTag']}

    async def rename_if_match(self, action):
        if action not in self.plan['preserve_by_rename']:
            raise RolloverHeld('rename_outside_verified_plan')
        current = await self.read_item(action['item_id'])
        self._identity(current, name=action['old_name'], parent=action['parent_id'], item_id=action['item_id'])
        if current['eTag'] != action['if_match']:
            raise RolloverHeld('control_version_changed')
        _, changed = await self._json(
            'PATCH', self._item(action['item_id']),
            headers={'If-Match': action['if_match']},
            json={'name': action['new_name'], '@microsoft.graph.conflictBehavior': 'fail'})
        self._identity(changed, name=action['new_name'], parent=action['parent_id'], item_id=action['item_id'])
        actual = await self.read_item(action['item_id'])
        self._identity(actual, name=action['new_name'], parent=action['parent_id'], item_id=action['item_id'])
        return actual

    async def read_item(self, item_id):
        allowed = {x['item_id'] for x in self.plan['preserve_by_rename']}
        allowed.add(self.plan['retained_case_lock']['item_id'])
        if item_id not in allowed:
            raise RolloverHeld('item_outside_verified_plan')
        _, result = await self._json('GET', self._item(item_id))
        return result

    async def compare_and_swap(self, key, etag, state):
        bound = self.reservations.get(key)
        if (not bound or bound['etag'] != etag
                or state != {**bound['state'], 'phase': 'reserved'}):
            raise RolloverHeld('reservation_commit_not_owned')
        _, changed = await self._json(
            'PATCH', self._item(bound['item_id']), headers={'If-Match': etag},
            json={'name': 'reserved.json', '@microsoft.graph.conflictBehavior': 'fail'})
        self._identity(changed, name='reserved.json', parent=bound['folder_id'], item_id=bound['item_id'])
        _, actual = await self._json('GET', self._item(bound['item_id']))
        self._identity(actual, name='reserved.json', parent=bound['folder_id'], item_id=bound['item_id'])
        _, saved = await self._json('GET', self._item(bound['item_id']) + '/content')
        if saved != bound['state'] or actual['eTag'] != changed['eTag']:
            raise RolloverHeld('reservation_commit_readback_mismatch')
        return {'state': {**saved, 'phase': 'reserved'}, 'etag': actual['eTag']}
