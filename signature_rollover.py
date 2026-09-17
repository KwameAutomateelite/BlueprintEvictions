"""Plan preservation of a completed signing cycle; performs no external writes.

The caller must freshly read these inputs after obtaining exclusive ownership.
Never apply this plan without the exclusive transition claim and If-Match checks.
An incomplete transition must be reconciled, not retried as a new send.
"""
import hashlib
import re


class RolloverHeld(ValueError):
    pass


def plan_completed_signature_rollover(*, request, case, approval, package,
                                     receipt, provider_request, signed_file,
                                     signed_pdf, controls):
    def require(condition, reason):
        if not condition:
            raise RolloverHeld(reason)

    record = request.get('record_id')
    folder = case.get('fields', {}).get('Case Folder ID')
    new_hash = request.get('approved_document_sha256', '')
    require(bool(re.fullmatch(r'rec[A-Za-z0-9]{14}', record or '')) and
            case.get('id') == record and bool(folder), 'case_identity_missing')
    require(case.get('fields', {}).get('Status') == 'To Owner for Review (N)',
            'new_notice_not_awaiting_owner_review')
    require(bool(re.fullmatch(r'[a-f0-9]{64}', new_hash)), 'new_approved_hash_missing')
    old_id = approval.get('signature_request_id')
    old_hash = approval.get('approved_document_sha256', '')
    require(bool(re.fullmatch(r'[a-f0-9]{40}', old_id or '')) and
            bool(re.fullmatch(r'[a-f0-9]{64}', old_hash)) and
            approval.get('approved') is True, 'prior_approval_incomplete')
    require(new_hash != old_hash, 'same_document_already_attempted')
    mode = request.get('test_mode')
    require(type(mode) is bool, 'signing_mode_missing')
    require(approval.get('record_id') == record and approval.get('folder_id') == folder,
            'prior_approval_case_mismatch')
    require(package.get('record_id') == record and package.get('folder_id') == folder
            and package.get('signature_request_id') == old_id
            and package.get('signature_status') == 'complete'
            and package.get('test_mode') is mode, 'prior_package_not_complete')
    provider = provider_request
    metadata = provider.get('metadata') or {}
    require(provider.get('signature_request_id') == old_id
            and provider.get('is_complete') is True
            and provider.get('is_declined') is False
            and provider.get('has_error') is False
            and provider.get('test_mode') is mode
            and metadata.get('record_id') == record
            and metadata.get('approved_document_sha256') == old_hash,
            'provider_completion_not_verified')
    require(isinstance(signed_pdf, bytes) and signed_pdf.startswith(b'%PDF-'),
            'signed_pdf_missing')
    digest = hashlib.sha256(signed_pdf).hexdigest()
    require(signed_file.get('id') == package.get('file_id')
            and signed_file.get('parentReference', {}).get('id') == folder
            and bool(signed_file.get('file'))
            and signed_file.get('size') == len(signed_pdf)
            and digest == package.get('sha256'), 'signed_file_binding_failed')
    require(receipt.get('record_id') == record
            and receipt.get('case_folder_id') == folder
            and receipt.get('signature_request_id') == old_id
            and receipt.get('sha256') == digest
            and receipt.get('test_mode') is mode
            and receipt.get('delivery_verified') is True
            and bool(receipt.get('sent_message_id'))
            and bool(receipt.get('dispatch_id')), 'prior_delivery_not_verified')
    old_claim = hashlib.sha256(('owner-signature-case-v1\n' + record).encode()).hexdigest()
    names = ['signature-approval.json', 'signed-package.json',
             'Maddy signature approval publication', 'Maddy signature send - ' + old_claim]
    require(isinstance(controls, list) and len(controls) == len(names),
            'control_inventory_incomplete')
    by_name = {item.get('name'): item for item in controls}
    require(set(by_name) == set(names), 'control_inventory_mismatch')
    require(len({item.get('id') for item in controls}) == len(names),
            'duplicate_control_item')
    actions = []
    for name in names:
        item = by_name[name]
        require(bool(item.get('id')) and bool(item.get('eTag')) and
                item.get('parentReference', {}).get('id') == folder,
                'control_location_or_version_missing')
        actions.append({'item_id': item['id'], 'if_match': item['eTag'],
                        'parent_id': folder, 'old_name': name,
                        'new_name': 'completed-' + old_id + '-' + name})
    transition = hashlib.sha256((record + '\n' + old_id).encode()).hexdigest()
    # Never rename/remove the case-wide lock: legacy claim callers must remain
    # blocked throughout the transition, including between remote writes.
    retained_lock = actions.pop()
    return {'record_id': record, 'folder_id': folder, 'prior_request_id': old_id,
            'prior_signed_sha256': digest, 'new_approved_sha256': new_hash,
            'exclusive_transition_name': 'Maddy completed signing transition - ' + transition,
            'conflict_behavior': 'fail', 'preserve_by_rename': actions,
            'retained_case_lock': retained_lock,
            'delete_files': False, 'send_signature': False,
            'requires_fresh_revalidation_after_exclusive_claim': True}
