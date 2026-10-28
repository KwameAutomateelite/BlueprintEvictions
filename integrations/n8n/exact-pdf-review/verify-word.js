const s = $('Validate Review Context').first().json;
const m = $input.first().json;
if (m.id !== s.document_id || m.parentReference?.id !== s.folder_id ||
    !m.eTag || !m.webUrl || !/\.docx$/i.test(m.name || '')) {
  throw Error('Saved Word file is outside this case or is not an editable notice. Nothing approved.');
}
return [{json: m}];
