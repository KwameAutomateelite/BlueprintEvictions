const rec=$('Read Bound Case').first().json;
const word=$('Require Bound Training Document').first().json;
const policy=$('Load Current Reviewer Policy').first().json._noticePolicy;
if (!policy?.enabled) throw Error('Review preparation held by current policy.');
let draft;try{draft=$('Prepare Verified Draft').first().json;}catch{}
return [{json:{fields:{Record_ID:rec.id},renderedPayload:draft?.payload || {},context:{record_id:rec.id,
  case_name:rec.fields?.['Case Name'] || '',
  case_folder_id:$('Require Case Folder ID').first().json.folder_id,
  onedrive_file_id:word.id,onedrive_share_link:word.webUrl,word_version:word.eTag,
  runtime_policy:policy,request_id:draft?.request_id || null,attachments_required:draft?.payload?.attachments_required || []}}}];
