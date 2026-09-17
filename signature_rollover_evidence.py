"""Fresh authenticated evidence loader for the Maddy-only rollover trial."""
import base64
import hashlib
import re
from datetime import datetime, timezone

from signature_rollover import RolloverHeld
from signature_rollover_transport import verify_bridge_url


async def read_completed_cycle_evidence(client, *, url, secret, provider_key, request):
    verify_bridge_url(url)
    if (not secret or not provider_key or request.get('test_mode') is not True
            or request.get('record_id') != 'recxLwYEXtb2ZrWen'):
        raise RolloverHeld('evidence_reader_outside_test_scope')
    response = await client.post(url, headers={'X-Maddy-Signature-Key': secret},
                                 json={'record_id': request['record_id'], 'test_mode': True})
    response.raise_for_status()
    raw = response.json()
    if raw.get('test_only') is not True:
        raise RolloverHeld('evidence_scope_not_verified')
    try:
        observed = datetime.fromisoformat(raw['observed_at'].replace('Z', '+00:00'))
        age = (datetime.now(timezone.utc) - observed).total_seconds()
        pdf = base64.b64decode(raw['signed_pdf_base64'], validate=True)
    except (KeyError, ValueError, TypeError):
        raise RolloverHeld('evidence_bytes_or_timestamp_invalid')
    if not 0 <= age <= 120:
        raise RolloverHeld('evidence_is_not_fresh')
    keys = ['case', 'approval', 'package', 'receipt', 'signed_file', 'controls']
    if any(key not in raw for key in keys):
        raise RolloverHeld('evidence_inventory_incomplete')
    package = raw['package']
    prior = package.get('signature_request_id', '')
    if (not re.fullmatch(r'[a-f0-9]{40}', prior)
            or package.get('record_id') != request['record_id']
            or not pdf.startswith(b'%PDF-')
            or hashlib.sha256(pdf).hexdigest() != package.get('sha256')):
        raise RolloverHeld('evidence_package_identity_failed')
    response = await client.get('https://api.hellosign.com/v3/signature_request/' + prior,
                                auth=(provider_key, ''))
    response.raise_for_status()
    provider = response.json().get('signature_request')
    if not isinstance(provider, dict):
        raise RolloverHeld('provider_evidence_missing')
    # The planner checks complete/declined/error, metadata, mode, case/folder,
    # approval, receipt and bytes together before any preservation is allowed.
    return {**{key: raw[key] for key in keys}, 'signed_pdf': pdf,
            'provider_request': provider}
