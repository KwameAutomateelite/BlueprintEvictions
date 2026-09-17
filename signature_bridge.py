"""Narrow authenticated client for verified signed-package storage."""
import hashlib
from urllib.parse import urlparse


class SignatureClaimExists(ValueError):
    """An authenticated, correctly bound response explicitly reported a duplicate."""

async def store_signed_package(client, url, secret, *, record_id, request_id,
                               approved_sha256, pdf, test_mode):
    target = urlparse(url)
    if target.scheme != 'https' or target.hostname != 'blueprintevictions.app.n8n.cloud' or not target.path.startswith('/webhook/') or target.username or target.password or target.query or target.fragment:
        raise ValueError('Invalid signature storage endpoint')
    if not secret or type(test_mode) is not bool or not pdf.startswith(b'%PDF-'):
        raise ValueError('Incomplete signed-package request')
    import base64
    response = await client.post(url, headers={'X-Maddy-Signature-Key':secret}, json={
        'record_id':record_id, 'signature_request_id':request_id,
        'approved_document_sha256':approved_sha256,
        'pdf_base64':base64.b64encode(pdf).decode(), 'test_mode':test_mode})
    response.raise_for_status()
    result = response.json()
    manifest = result.get('manifest') or {}
    expected = {'record_id':record_id, 'signature_request_id':request_id,
                'sha256':hashlib.sha256(pdf).hexdigest(), 'test_mode':test_mode,
                'signature_status':'complete'}
    if result.get('stored') is not True or any(manifest.get(k) != v for k,v in expected.items()) or not manifest.get('file_id') or not manifest.get('folder_id'):
        raise ValueError('Storage service did not verify the requested signed package')
    return result

async def record_signature_approval(client, url, secret, approval):
    target=urlparse(url)
    if target.scheme!='https' or target.hostname!='blueprintevictions.app.n8n.cloud' or not target.path.startswith('/webhook/') or target.username or target.password or target.query or target.fragment or not secret:
        raise ValueError('Invalid signature approval endpoint')
    response=await client.post(url,headers={'X-Maddy-Signature-Key':secret},json=approval)
    response.raise_for_status()
    result=response.json()
    expected={**approval,'approved':True}
    if result.get('recorded') is not True or result.get('approval')!=expected:
        raise ValueError('Signing request was sent but approval storage could not be verified; do not resend automatically')
    return result

async def claim_signature_send(client, url, secret, record_id, approved_sha256, signer_email):
    target=urlparse(url)
    if target.scheme!='https' or target.hostname!='blueprintevictions.app.n8n.cloud' or not target.path.startswith('/webhook/') or target.username or target.password or target.query or target.fragment or not secret:
        raise ValueError('Invalid signature claim endpoint')
    response=await client.post(url,headers={'X-Maddy-Signature-Key':secret},json={
        'record_id':record_id,'approved_document_sha256':approved_sha256,'signer_email':signer_email})
    response.raise_for_status()
    result=response.json()
    if result.get('record_id')!=record_id or result.get('approved_document_sha256')!=approved_sha256 or result.get('signer_email')!=signer_email.strip().lower() or not result.get('folder_id') or not result.get('expected_status'):
        raise ValueError('Signing claim could not be verified; no signature email sent')
    if result.get('claimed') is False and result.get('claim_folder_id') is None:
        raise SignatureClaimExists('Existing case signature claim verified; inspect completed cycle before any new send')
    if result.get('claimed') is not True or not result.get('claim_folder_id'):
        raise ValueError('Signing claim outcome uncertain; no signature email sent')
    return result
