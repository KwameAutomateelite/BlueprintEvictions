const snapshot = JSON.parse(JSON.stringify($input.first().json));
const c = snapshot.context || {};
const p = c.runtime_policy || {};
if (p.enabled !== true) return [{json: {snapshot, needs_preview: false}}];
const reviewer = String(p.reviewer_email || '').toLowerCase();
if (!snapshot.fields || !c.record_id || !c.onedrive_file_id || !c.case_folder_id ||
    !reviewer || !p.reviewer_senders?.some(x => String(x).toLowerCase() === reviewer)) {
  throw Error('PDF review is missing the verified case, Word file, folder or reviewer. Nothing approved.');
}
return [{json: {snapshot, needs_preview: true, record_id: c.record_id,
  document_id: c.onedrive_file_id, folder_id: c.revision_folder_id || c.case_folder_id, reviewer}}];
