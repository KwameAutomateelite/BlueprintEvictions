const m = $input.first().json;
const s = $('Validate Review Context').first().json;
if (!m.id || !m.webUrl || m.parentReference?.id !== s.folder_id ||
    m.name !== $('Prepare Exact PDF').first().json.name || !m.eTag) {
  throw Error('Review PDF was not saved beside this case’s Word file. Nothing approved.');
}
return [{json: m}];
