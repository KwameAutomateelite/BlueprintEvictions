# Current-order rent research

These are the reviewed n8n Code node sources for `Blueprint Proactive Notice Case Research`.

- `rent-evidence.cjs` validates amounts as cents, including equivalent `$4300` and `$4,300.00` formatting.
- `research-validator.js` retains source quotations, existing confirmed facts, source conflicts, case holds and case-folder/unit guards. It records the period-year basis separately from literal quoted facts.
- A single month-only current order can use an actual current unpaid-rent statement only when its message and order are in that same month and year, no more than 31 days apart, with the exact amount already recorded in the order. Quoted old mail, stale years, additional months and changes to the recorded amount are rejected. A lease date or monthly rate alone does not establish arrears.
- `select-research-path.js` avoids reading documents again when a complete current order only lacks this one rent year. It still reads current business mail and conversation threads, including holds and corrections. Incomplete addresses, parties, signer or phone, and complex rent lists retain the full asset-research path.
- The output only prepares an unsigned notice draft. It does not update financial case facts or approve, sign or send a document.

Validation: `node integrations/n8n/case-research/test-rent-evidence.cjs`. The reported execution was also replayed locally with its bound private sources, followed through the unchanged notice payload validator. No case data or private exports are committed.

Deployment: preserve all existing workflow IDs, credentials, settings and static data. Save the current export, compare its version immediately before an update, replace only the reviewed Code sources, and insert Select Required Research plus its boolean branch after Current Case Inventory. The true branch enters Build Case Mail Query; the false branch retains Select Relevant Source Assets. Read back the active version after publication. A source commit alone is not production proof.
