# Residential three day pay or quit notice standard

For residential three-day pay-or-quit requests, `/generate-notice` fills one
approved, editable, two-page Word master. Rent row count does not rearrange the
form. Commercial notices and other notice families retain their existing paths.

The design authority is Anne-Michelle's October 26, 2026 residential PDF,
SHA256 `0e23f57004a2f9842acaab92a334d46c031a5d09e31f4f3900b91f531418825d`.
Its filled example values are not stored in the new master. The original public
Word templates remain intact. Their existing Blueprint artwork is retained.
The PDF's non-printing reset-form utility is not part of the notice output.

## Field contract

Existing request names remain accepted. The fields below distinguish
people and contact blocks that the old generator conflated.

| Field | Visible position and rule |
| --- | --- |
| `tenant_names` | Addressees after To; two reserved lines |
| `property_address_street`, `property_address_city`, `apt` / `unit` | Premises block; no duplicated unit |
| `county` / `county_name` | County of the premises; printed once after COUNTY OF |
| `amounts_due` | Paired period/amount rows; one through six, unused rows blank |
| `total_amount_due` | Boxed total; must equal the supplied rows |
| `payment_recipient_name` | Payment/possession recipient; falls back to `landlord_name` |
| `payment_address_street`, `payment_address_city`, `payment_phone` | Payment block; legacy landlord address/phone accepted |
| `signer_name`, `signer_title` | Printed name and role immediately below signature line |
| `housing_provider_street`, `housing_provider_city`, `housing_provider_phone` | Housing-provider block; defaults to supplied payment contact |
| `notice_date` | Date beside signature; defaults to current America/Los_Angeles date |
| `service_date` | Does not populate the service line; that line remains blank |
| `landlord_email` | Accepted for compatibility but never inserted into the notice body |

`authorized_signer_name` and `authorized_signer_title` are accepted aliases.
The owner/payee is not assumed to be the authorized signer. Both signer name and title are required. Missing signatory information returns
HTTP 422 to the existing human-review branch; it never produces a blank-signatory
notice or substitutes an owner/payee. Example names, phone numbers, addresses and dates in
the reference never become defaults. Preserve intentional company spelling.

An inconsistent rent total, malformed amount/date, ambiguous raw order-form
labels, missing premises/county, more than six rent rows, or text too long for
its reserved space produces HTTP 422. Do not truncate values, shrink type,
reword static text, or substitute a different notice family to bypass review.

## Integration and deployment

The endpoint still returns DOCX bytes, Content-Disposition and the existing
X-Attachments-Needed header. X-Notice-Template-Version identifies this master.
`/health` reports `residential_template_version` for deployment verification.

Deploy the application with its updated Dockerfile, `residential_notice.py` and
`templates/residential-pay-or-quit-2026-10.docx` together. The existing n8n URL,
folder bindings, review process and signature controls remain intact. Do not
redeploy from a different branch or transfer ownership as part of this fix.
Existing case files are not automatically overwritten.

The n8n intake's First Name + Last Name and Signer Title describe the designated
signatory, as confirmed in the April 10 and May 11 calls. The separate Property
Owner's Name remains the owner/payee. Initial generation and revision carry the
signatory separately, and the approval email displays both name and title.
Extracting an owner from a document must never overwrite the signatory surname.
Existing waiting executions retain their captured versions and are not replayed.

The saved-document revision service recognizes the six bookmarked rent rows,
reconciles the total after edits, binds both phone positions explicitly and
rejects edits to fixed wording or fields that would overflow their reserved
space. The legacy table editor remains available for older notice families.

For later case edits, update case fields in the same master/version and retain
its static paragraphs, page breaks, typography and artwork. Avoid generic
paragraph/layout rewrite passes. Recheck overflow and arithmetic after edits.

## Verification

The regression suite checks semantic field placement, blank service date,
separate signer, omission of owner email, arithmetic, row bounds, overflow,
stale-example exclusion and endpoint compatibility. Package comparisons verify
that one-, three- and six-row generation changes field text while preserving
layout structures, artwork and relationships. Fictional review copies are
rendered on both pages to compare with the approved form.

Generate local review copies with `tools/preview_residential_standard.py`.
No production case, bot message, signing request or paid service is used by
these tests.
