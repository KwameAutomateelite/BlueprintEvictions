# Controlled signing integration

This candidate applies only to the four explicitly allowlisted Maddy/Amankwah records and Kwame’s two controlled signer addresses. Provider signing is forced into test mode on this path. Other records retain existing behavior.

All four environment variables must be configured together: SIGNATURE_APPROVAL_BRIDGE_URL, SIGNATURE_STORAGE_BRIDGE_URL, SIGNATURE_STORAGE_BRIDGE_SECRET, SIGNATURE_SEND_CLAIM_URL. Use the authenticated n8n bridges, with their referenced workers published first. Never commit the secret.

The case must be Sent for Signature or To Owner for Review (N). A case-level exclusive claim precedes provider send. Persist the approved PDF hash and provider request identity. A downloadable callback is authenticated and checked against Dropbox Sign before saving and verifying the completed PDF and manifest. Only then does the controlled callback move the case to With Process Server (N), the existing dispatch queue status. This means queued; actual delivery is tracked separately by the dispatch worker.

Incomplete or conflicting requests require reconciliation before retry. A provider send followed by an approval-storage failure must never be blindly resent. Concurrent external case edits are not transactionally locked by Airtable.

Validation: python3 -m unittest discover -s tests -p 'test_signature_*.py'

Repeat-notice handling is separately disabled by default. Set SIGNATURE_ROLLOVER_TEST_ENABLED=1 only with SIGNATURE_ROLLOVER_EVIDENCE_URL and SIGNATURE_ROLLOVER_STORAGE_URL pointing to the authenticated test bridges. This path is restricted to the Maddy test record and the Automate Elite test signer. It requires an explicit, correctly bound duplicate claim, the exact approved PDF hash, fresh provider completion and delivery evidence, stable control versions, and an exclusive preservation reservation. The retained case lock is never renamed. An uncertain or interrupted transition holds for reconciliation rather than resending.

The repeat-notice end-to-end acceptance test is not complete. Unit checks and a healthy deployment must not be presented as proof of owner signing, storage, delivery, or production readiness. The legacy non-controlled callback is unchanged and outside the controlled validation.

Railway rebuilds this connected repository when its configuration changes. Keep the deployed implementation in this repository; do not depend on an uploaded-only code patch. Verify the deployed commit and the /openapi.json SendSignatureRequest properties require_verified_signature and approved_pdf_sha256 before any test approval. A health response alone does not prove the correct signing version is running.
