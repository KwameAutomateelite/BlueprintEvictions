"""Validate Dropbox Sign notifications before any case mutation."""
import hashlib
import hmac


def verify_event(payload: dict, api_key: str) -> bool:
    event = payload.get('event')
    if not isinstance(event, dict) or not api_key:
        return False
    timestamp, kind, supplied = (event.get(k) for k in ('event_time', 'event_type', 'event_hash'))
    if not isinstance(timestamp, (str, int)) or isinstance(timestamp, bool):
        return False
    if not isinstance(kind, str) or not kind or not isinstance(supplied, str):
        return False
    expected = hmac.new(api_key.encode(), (str(timestamp) + kind).encode(), hashlib.sha256).hexdigest()
    return hmac.compare_digest(expected, supplied)


def verified_request_record(callback_request: dict, authoritative_request: dict) -> str:
    """Bind event to provider-fetched request; callback metadata alone is untrusted."""
    wanted = callback_request.get('signature_request_id')
    actual = authoritative_request.get('signature_request_id')
    if not wanted or wanted != actual:
        raise ValueError('Signature request mismatch')
    if authoritative_request.get('is_complete') is not True:
        raise ValueError('Signature request is not complete')
    if authoritative_request.get('is_declined') or authoritative_request.get('has_error'):
        raise ValueError('Signature request failed or was declined')
    record = (authoritative_request.get('metadata') or {}).get('record_id')
    if not isinstance(record, str) or not record.startswith('rec'):
        raise ValueError('Provider request has no bound case record')
    if (callback_request.get('metadata') or {}).get('record_id') != record:
        raise ValueError('Callback record mismatch')
    return record
