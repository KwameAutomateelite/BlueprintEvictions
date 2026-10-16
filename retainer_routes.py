"""Authenticated, synthetic-only unsigned retainer rendering."""
import hmac
import os
from fastapi import APIRouter, HTTPException, Request
from retainer_draft import render_test_retainer

router = APIRouter()

@router.post('/test-retainer-draft')
async def test_retainer_draft(request: Request):
    secret = os.environ.get('SIGNATURE_STORAGE_BRIDGE_SECRET', '')
    supplied = request.headers.get('x-narnia-retainer-secret', '')
    if not secret or not hmac.compare_digest(secret, supplied):
        raise HTTPException(status_code=401, detail='Retainer authentication required')
    try:
        plan = await request.json()
        return render_test_retainer(plan)
    except (ValueError, TypeError, AttributeError) as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
