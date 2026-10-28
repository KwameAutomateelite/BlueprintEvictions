# Exact residential PDF review

The original residential three-day PDF is the review and approval copy. The saved
Word file remains the editing source. PDF generation reads the actual saved Word
bytes through `/render-saved-notice-pdf`; it does not reuse the initial intake
values after an edit.

`build.py` accepts a freshly read parent n8n workflow and builds an offline
candidate. It performs no API writes. Private workflow exports, credential
references and case records do not belong in this repository.

```sh
python integrations/n8n/exact-pdf-review/build.py \
  /private/workflow-before.json /private/output EXISTING_PREVIEW_CHILD_ID
```

The preview child has only an Execute Workflow trigger. It verifies the bound
case folder and Word file, renders the exact PDF, saves it beside Word, checks
the saved PDF bytes, and grants the configured reviewer read access without an
invitation email. Disabled reviewer policy and other notice families preserve
their existing paths. A residential document that has lost its form fields
fails instead of silently switching formats.

The parent candidate adds this preview after initial drafting and both existing
edit paths. Existing source-node code, credentials, workflow settings, reviewer
policy, claims and signing gates are preserved. The review email links to PDF
first and Word second. After manual Word edits, the authorized reviewer can reply
`REFRESH PDF` to the current review email. That command requires the existing
verified reply, current case and token checks; it does not approve anything.

Before approval, the saved Word is read again and rendered again. Its resulting
PDF must match the reviewed PDF hash, and the saved review file must still have
those same bytes. The existing Word version recheck and frozen delivery continue
afterward. A changed file requires fresh review. Previous review PDFs are retained
when the Word contents change.

Installation must compare the latest parent version, nodes, connections and
settings with the baseline, then publish the specific saved candidate version
and verify its active readback. Do not trigger or replay the business workflow
to test installation. Do not enable rollout, signing or other workflows.

Validation:

```sh
python -m pytest tests/test_pdf_review_builder.py tests/test_saved_word_pdf.py -q
node --test tests/test_pdf_review_contracts.cjs
```

The cloud rehearsal used a fictional Word notice in an isolated QA folder,
updated its rent and signer, regenerated its PDF, preserved the earlier review
file, and verified that approval selected exactly the reviewed PDF bytes. No
client message, signing request, Grok call or production case job was executed.
