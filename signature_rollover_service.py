"""Reserve a later signing cycle only after an explicit duplicate claim.

Called by the signing route behind an opt-in, Maddy-only, test-mode gate.
All actual preservation is delegated to the verified executor and Graph store.
"""
from signature_rollover import plan_completed_signature_rollover, RolloverHeld
from signature_rollover_evidence import read_completed_cycle_evidence
from signature_rollover_executor import reserve_completed_cycle
from signature_rollover_graph import GraphRolloverStore
from signature_rollover_transport import RolloverGraphTransport


async def claim_completed_cycle(client, *, evidence_url, storage_url, secret,
                                provider_key, request, signer_email,
                                expected_folder, expected_status):
    if (request.get('test_mode') is not True
            or request.get('record_id') != 'recxLwYEXtb2ZrWen'
            or signer_email.strip().lower() != 'kwame@automateelite.com'
            or expected_status != 'To Owner for Review (N)'):
        raise RolloverHeld('repeat_signing_outside_approved_test_scope')

    async def read():
        return await read_completed_cycle_evidence(
            client, url=evidence_url, secret=secret,
            provider_key=provider_key, request=request)

    initial = await read()
    plan = plan_completed_signature_rollover(request=request, **initial)
    if plan['folder_id'] != expected_folder:
        raise RolloverHeld('case_folder_changed_before_rollover')
    transport = RolloverGraphTransport(client, url=storage_url, secret=secret,
                                      record_id=request['record_id'], test_mode=True)
    store = GraphRolloverStore(transport, plan)
    first = True

    async def evidence(_request):
        nonlocal first
        if first:
            first = False
            return initial
        return await read()

    reservation = await reserve_completed_cycle(
        store=store, load_evidence=evidence, request=request)
    if (reservation.get('reserved') is not True
            or reservation.get('record_id') != request['record_id']
            or reservation.get('approved_document_sha256') != request['approved_document_sha256']):
        raise RolloverHeld('repeat_signature_reservation_unverified')
    return {'claimed': True, 'record_id': request['record_id'],
            'folder_id': expected_folder, 'expected_status': expected_status,
            'approved_document_sha256': request['approved_document_sha256'],
            'claim_folder_id': reservation['case_lock_id'],
            'repeat_cycle_reservation': reservation['reservation_name']}
