"""Verified provider event -> saved package -> checked case status.

Clients are supplied by the route; this service never guesses credentials or scope.
"""
from urllib.parse import quote
from signature_integrity import verify_event, verified_request_record
from signature_bridge import store_signed_package

async def process_signature_callback(payload, *, provider, airtable, storage_http,
                                     api_key, bridge_url, bridge_secret,
                                     base_id, table_id, allowed_records):
    if not verify_event(payload, api_key):
        raise ValueError('Invalid signature event authentication')
    if payload['event']['event_type'] != 'signature_request_downloadable':
        return {'acknowledged':True, 'stored':False}
    request = payload.get('signature_request') or {}
    rid = request.get('signature_request_id')
    if not isinstance(rid, str) or not rid:
        raise ValueError('Missing provider request ID')
    root='https://api.hellosign.com/v3/signature_request/'
    response=await provider.get(root+quote(rid,safe=''),auth=(api_key,''))
    response.raise_for_status()
    actual=response.json()['signature_request']
    record_id=verified_request_record(request,actual)
    if record_id not in allowed_records:
        raise ValueError('Case outside callback rollout scope')
    mode=actual.get('test_mode')
    digest=(actual.get('metadata') or {}).get('approved_document_sha256')
    if type(mode) is not bool or not isinstance(digest,str) or len(digest)!=64:
        raise ValueError('Provider approval identity incomplete')
    response=await provider.get(root+'files/'+quote(rid,safe=''),params={'file_type':'pdf'},auth=(api_key,''))
    response.raise_for_status()
    stored=await store_signed_package(storage_http,bridge_url,bridge_secret,
        record_id=record_id,request_id=rid,approved_sha256=digest,pdf=response.content,test_mode=mode)
    url=f'https://api.airtable.com/v0/{base_id}/{table_id}/{quote(record_id,safe="")}'
    response=await airtable.get(url);response.raise_for_status();case=response.json()
    fields=case.get('fields') or {}
    if case.get('id')!=record_id or fields.get('Case Folder ID')!=stored['manifest']['folder_id']:
        raise ValueError('Case identity changed after package storage')
    status=fields.get('Status')
    if status=='With Process Server (N)':
        return {'acknowledged':True,'stored':True,'status_verified':True,'duplicate_status_write':False}
    if stored.get('expected_status') not in {'Sent for Signature', 'To Owner for Review (N)'} or status!=stored['expected_status']:
        raise ValueError('Case advanced or changed; signed status was not overwritten')
    response=await airtable.patch(url,json={'fields':{'Status':'With Process Server (N)','Dispatched On':None}})
    response.raise_for_status()
    response=await airtable.get(url);response.raise_for_status();check=response.json()
    if check.get('id')!=record_id or check.get('fields',{}).get('Status')!='With Process Server (N)' or check.get('fields',{}).get('Dispatched On'):
        raise ValueError('Dispatch queue status read-back mismatch')
    return {'acknowledged':True,'stored':True,'status_verified':True}
