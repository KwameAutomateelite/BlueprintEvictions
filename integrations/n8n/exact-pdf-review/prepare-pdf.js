const r = $input.first().json;
const source = $('Prepare Saved Word Snapshot').first().json;
const b = Buffer.from(r.pdf_base64 || '', 'base64');
if (!r.is_standard || r.source_docx_sha256 !== source.expected_sha256 ||
    !['residential-original-pdf-2026-10-27', 'residential-original-pdf-2026-10-28-format3'].includes(r.template_version) ||
    b.subarray(0, 5).toString() !== '%PDF-' || !r.case_fields ||
    require('crypto').createHash('sha256').update(b).digest('hex') !== r.pdf_sha256) {
  throw Error('Exact PDF does not match this saved Word version. Nothing approved.');
}
const ctx = $('Validate Review Context').first().json.snapshot.context;
const safe = value => String(value || '').normalize('NFKC')
  .replace(/[<>:"/\\|?*\x00-\x1f]/g, ' ').replace(/\s+/g, ' ').trim().replace(/[. ]+$/, '');
const f = r.case_fields;
const date = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(f.notice_date || '');
const stamp = date ? `${date[3].length === 2 ? '20' : ''}${date[3]}-${date[1].padStart(2,'0')}-${date[2].padStart(2,'0')}` : safe(f.notice_date);
const prefix = ctx.runtime_policy?.test_mode === true ? 'TEST ONLY - ' : '';
const stem = prefix + ['3-Day Notice', safe(ctx.case_name).slice(0,45), safe(f.property_address_street).slice(0,65)].filter(Boolean).join(' - ');
const suffix = (stamp ? ' - ' + stamp : '') + ' - REVIEW - ' + source.expected_sha256.slice(0,16) + '.pdf';
const name = stem.slice(0,180-suffix.length).replace(/[. ]+$/, '') + suffix;
return [{json: {name}, binary: {data: await this.helpers.prepareBinaryData(b, name, 'application/pdf')}}];
