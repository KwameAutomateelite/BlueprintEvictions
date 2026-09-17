"""Adapter-driven rollover reservation. Not connected to the live signing route.

The storage adapter must implement atomic create-if-absent and compare-and-swap
using provider version preconditions. It must never implement these as read/PUT.
The caller supplies freshly authenticated evidence to the existing planner.
An uncertain write is held for reconciliation; this function never sends mail.
"""
import secrets

from signature_rollover import plan_completed_signature_rollover, RolloverHeld


async def reserve_completed_cycle(*, store, load_evidence, request):
    evidence = await load_evidence(request)
    plan = plan_completed_signature_rollover(request=request, **evidence)
    token = secrets.token_hex(32)
    state = {
        'record_id': plan['record_id'], 'folder_id': plan['folder_id'],
        'prior_request_id': plan['prior_request_id'],
        'new_approved_sha256': plan['new_approved_sha256'],
        'owner_token': token, 'phase': 'preserving',
    }
    # The key is the prior completed cycle, not the proposed next document.
    # Different candidate versions must compete for this one reservation.
    reservation = await store.create_exclusive(plan['exclusive_transition_name'], state)
    if reservation is None:
        raise RolloverHeld('transition_already_attempted_reconcile_before_retry')
    if reservation.get('state') != state or not reservation.get('etag'):
        raise RolloverHeld('exclusive_transition_unverified')
    # Recheck all provider evidence after winning the reservation. Do not act on
    # evidence captured while another request might have been changing it.
    fresh = plan_completed_signature_rollover(
        request=request, **await load_evidence(request))
    if fresh != plan:
        raise RolloverHeld('completed_cycle_changed_after_reservation')
    for action in plan['preserve_by_rename']:
        saved = await store.rename_if_match(action)
        if (saved.get('id') != action['item_id']
                or saved.get('name') != action['new_name']
                or saved.get('parentReference', {}).get('id') != action['parent_id']):
            raise RolloverHeld('preserved_control_unverified')
    lock = plan['retained_case_lock']
    retained = await store.read_item(lock['item_id'])
    if (retained.get('id') != lock['item_id']
            or retained.get('name') != lock['old_name']
            or retained.get('eTag') != lock['if_match']
            or retained.get('parentReference', {}).get('id') != lock['parent_id']):
        raise RolloverHeld('case_lock_changed')
    committed = {**state, 'phase': 'reserved'}
    result = await store.compare_and_swap(
        plan['exclusive_transition_name'], reservation['etag'], committed)
    if result.get('state') != committed or not result.get('etag'):
        raise RolloverHeld('new_cycle_reservation_unverified_do_not_send')
    return {
        'reserved': True, 'record_id': plan['record_id'],
        'approved_document_sha256': plan['new_approved_sha256'],
        'reservation_name': plan['exclusive_transition_name'],
        'reservation_etag': result['etag'],
        'case_lock_id': lock['item_id'],
    }
