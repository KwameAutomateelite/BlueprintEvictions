const r = $('Render Approved Saved Word as Exact PDF').item.json;
const snapshot = $('Build Approval Email').item.json.reviewSnapshot;
const c = snapshot.context;
const source = $('Capture Approved Word Snapshot').item.json;
const shared = await this.helpers.getBinaryDataBuffer(0, 'data');
const hash = b => require('crypto').createHash('sha256').update(b).digest('hex');
const pdf = Buffer.from(r.pdf_base64 || '', 'base64');
if (c.review_format !== 'pdf' || !c.review_pdf_file_id || !c.review_pdf_sha256 ||
    !r.is_standard || r.source_docx_sha256 !== source.docx_sha256 ||
    r.template_version !== c.review_pdf_template_version ||
    hash(shared) !== c.review_pdf_sha256 || r.pdf_sha256 !== c.review_pdf_sha256 ||
    hash(pdf) !== c.review_pdf_sha256) {
  throw Error('The saved Word or review PDF changed. A fresh PDF review is required before approval.');
}
return [{json: {}, binary: {data: await this.helpers.prepareBinaryData(pdf, 'approved-notice.pdf', 'application/pdf')}}];
