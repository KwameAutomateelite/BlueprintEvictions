# Blueprint case access and dispatch approval

These pure checks back the n8n case-name search, saved retainer reader, and post-signature dispatch gate. They contain no provider credentials or client files.

- `case-name-query.cjs`: case/party/address terms for fresh reads of both Airtable tables; no record ID question is needed.
- `case-asset-boundary.cjs`: a requested file must belong to the verified case root, or its verified retainer review child. Another case or similarly named folder is rejected. Version, parent and case bindings are checked again after reading.
- `dispatch-approval.cjs`: owner signature and notice approval do not authorize process-server delivery. Approval binds the current signed request, PDF digest, case/folder, intended recipient, mode and reviewer. The existing delivery claim still prevents duplicate sends.
- `dispatch-command-guards.cjs`: an explicit approval command must match the actual completed signed package, reviewer, recipient and test/live mode before delivery can run.
- `dispatch-review-receipts.cjs`: read-only review reconciles the saved approval and exact delivery receipt. An owned test email is reported as test delivery, and a missing or mismatched receipt stays unresolved. Signing, sending and case tracking remain distinct facts.
- `service-proof-policy.cjs`: proof must match the case name and full premises, including unit/ADU. Multiple matching orders remain ambiguous rather than selecting the newest. Dates must be valid calendar dates, and approval binds the actual proof, proposed fields and unchanged case snapshot.

Service deadline automation is limited to supported three-day pay-or-cure notices. Other notice types require office review of the deadline. This follows the [California Courts notice-type guidance](https://selfhelp.courts.ca.gov/eviction-landlord/notice-types) and [counting instructions](https://www.selfhelp.courts.ca.gov/eviction-landlord/fill-out). The current proof worker's human approval gate is retained while its Grok review journey is being completed.

Active workflows are identified in the business CTO runbook. Inspect their actual current versions before a change; preserve current credentials, settings, static data and unrelated nodes. Back up, test, publish, read back and retain rollback evidence. These source checks do not establish that an automatic card, signature completion, or delivery has passed acceptance.

Readable PDF source text uses the authenticated `/inspect-source-pdf` endpoint, hashes the uploaded bytes, includes actual filled form values and never fetches caller URLs or infers missing facts. Blank scans report an OCR need; a read failure is not absence. Existing source-research and signing approval guards remain independent.

Run `node integrations/n8n/case-access/test-case-access.cjs`, `node integrations/n8n/case-access/test-dispatch-state.cjs`, `node integrations/n8n/case-access/test-service-proof-policy.cjs`, and the PDF/source callback tests before changing these paths.
