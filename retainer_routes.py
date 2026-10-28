"""Authenticated unsigned retainer rendering for case plans and legacy tests."""
import hmac
import os
from fastapi import APIRouter, HTTPException, Request
from retainer_draft import render_test_retainer, render_retainer_draft

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


@router.post('/retainer-draft')
async def retainer_draft(request: Request):
    """Authenticated draft-only action; never sends or approves an agreement."""
    secret = os.environ.get('SIGNATURE_STORAGE_BRIDGE_SECRET', '')
    supplied = request.headers.get('x-narnia-retainer-secret', '')
    if not secret or not hmac.compare_digest(secret, supplied):
        raise HTTPException(status_code=401, detail='Retainer authentication required')
    try:
        return render_retainer_draft(await request.json())
    except (ValueError, TypeError, AttributeError) as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
