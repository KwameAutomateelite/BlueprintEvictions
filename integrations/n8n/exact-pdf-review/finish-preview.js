const before = $('Verify Saved Word Location').first().json;
const after = $('Recheck Saved Word Metadata').first().json;
const folder = $input.first().json;
const expected = $('Validate Review Context').first().json;
if (folder.id !== expected.snapshot.context.case_folder_id || !folder.folder ||
    !/^https:\/\//.test(folder.webUrl || '')) {
  throw Error('Address-folder link could not be verified. Do not claim the review is ready.');
}
if (after.id !== before.id || after.eTag !== before.eTag) {
  throw Error('Word changed while the PDF was being prepared. A fresh PDF review is required.');
}
const s = JSON.parse(JSON.stringify($('Validate Review Context').first().json.snapshot));
const r = $('Render Saved Word as Exact PDF').first().json;
const pdf = $('Verify Saved PDF Location').first().json;
const f = r.case_fields;
s.renderedPayload = {...s.renderedPayload, ...f};
s.fields = {...s.fields, TENANT_NAMES: f.tenant_names,
  PROPERTY_ADDRESS_STREET: f.property_address_street, PROPERTY_ADDRESS_CITY: f.property_address_city,
  COUNTY: f.county, LANDLORD_NAME: f.payment_recipient_name, LANDLORD_PHONE: f.payment_phone,
  TOTAL_AMOUNT_DUE: f.total_amount_due, AMOUNTS_DUE: f.amounts_due.map(x => x.amount),
  RENT_DUE_DATES: f.amounts_due.map(x => x.due_date)};
s.context = {...s.context, review_format: 'pdf', word_edit_link: before.webUrl,
  case_folder_url: folder.webUrl, review_pdf_name: pdf.name, word_edit_name: before.name,
  onedrive_share_link: pdf.webUrl, review_pdf_file_id: pdf.id, review_pdf_sha256: r.pdf_sha256,
  review_docx_sha256: r.source_docx_sha256, review_pdf_template_version: r.template_version,
  signer_name: f.signer_name, signer_title: f.signer_title,
  tenant_name: f.tenant_names, property_address: f.property_address_street + ', ' + f.property_address_city};
return [{json: s}];
