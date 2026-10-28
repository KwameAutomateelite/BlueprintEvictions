const r = $input.first().json;
const source = $('Prepare Saved Word Snapshot').first().json;
const b = Buffer.from(r.pdf_base64 || '', 'base64');
if (!r.is_standard || r.source_docx_sha256 !== source.expected_sha256 ||
    r.template_version !== 'residential-original-pdf-2026-10-27' ||
    b.subarray(0, 5).toString() !== '%PDF-' || !r.case_fields ||
    require('crypto').createHash('sha256').update(b).digest('hex') !== r.pdf_sha256) {
  throw Error('Exact PDF does not match this saved Word version. Nothing approved.');
}
const name = '3-Day Notice - PDF Review - ' + source.expected_sha256.slice(0, 16) + '.pdf';
return [{json: {name}, binary: {data: await this.helpers.prepareBinaryData(b, name, 'application/pdf')}}];
