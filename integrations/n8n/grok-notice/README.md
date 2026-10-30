Grok residential three-day notice tools
=====================================

These callable n8n tools connect Maddy to the existing Word generator and exact
PDF renderer. They do not call a model, send messages, change case status,
approve, sign, or replay intake. Older notice families and historical files stay
untouched. Runtime reviewer policy remains authoritative.

`draft_case_notice(record_id, request_id)` reads the exact current notice case,
verifies its bound folder under the approved case root, and parses intake facts
with the existing production deterministic parser. Missing signer information,
unclear rent periods/years or inconsistent totals stop before saving a notice.
It creates a new Word file using a conflict-failing upload session, verifies the
saved bytes, grants the configured reviewer edit access without sending an
invitation, then runs the already published exact-PDF review helper.

`prepare_case_notice_review(record_id, item_id, expected_version)` reads the
exact bound current Word file and produces the matching PDF. It is an explicit
file-writing action, separate from read-only inspection. Use it after manual
Word edits and to recover PDF preparation after Word was already saved.

The existing `revise_case_notice` still uses its conditional edit and verified
Airtable index; its successful edit then calls the same PDF preparation tool.
If PDF preparation fails after an edit, the receipt says Word was saved but the
PDF is not ready. Retry only PDF preparation, not the original edit.

Maddy must use the returned receipt as the source of truth. `draft_ready_for_review`
means saved and read back, with PDF and Word links. It never means approved,
signed, emailed or delivered. A notification without such a receipt does not
prove notice creation. Return the PDF for review and Word for edits, including
for rent lists that continue to additional pages. Do not recreate the notice in
chat or change the template layout.

`case-inputs.cjs` is the current production parser/payload snapshot from Step 2
Draft & Review, version febec03b-26c5-4442-98d7-dfe5a67d81f5, adapted only to pure
function parameters and exports. No historical rehearsal case restrictions or
model-generated values are used. The generator continues to validate all fields
and total amounts independently.

October 30 quality-control display

The shared exact-PDF helper gives newly saved review PDFs readable case/street/date/REVIEW names with the original source-hash suffix. It reads the actual bound address-folder metadata before returning its URL. Grok draft/refresh receipts include three explicit destinations: `pdf_url`, `word_edit_url`, and `folder_url`, along with the actual saved names and case/address. A review filename never indicates approval or signing.

Storage IDs, upload destinations, source-Word names/idempotency, existing case links, old files and approval byte/hash checks remain unchanged. The extra folder read validates the link; it does not create, move or rename a folder. Historical files in older/shared locations need individual reconciliation and are outside this display change.
