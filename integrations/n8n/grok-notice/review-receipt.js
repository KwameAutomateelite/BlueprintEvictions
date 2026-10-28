const s=$input.first().json,c=s.context || {};
if (c.review_format!=='pdf' || !c.review_pdf_file_id || !c.word_edit_link ||
    !/^[a-f0-9]{64}$/.test(c.review_pdf_sha256 || '') ||
    !/^[a-f0-9]{64}$/.test(c.review_docx_sha256 || '')) {
  throw Error('Saved Word is not the standardized residential form. No ready PDF was confirmed.');
}
return [{json:{state:'draft_ready_for_review',case_id:c.record_id,
  folder_id:c.case_folder_id,item_id:c.onedrive_file_id,version:c.word_version,request_id:c.request_id || null,
  pdf_item_id:c.review_pdf_file_id,pdf_url:c.onedrive_share_link,
  word_edit_url:c.word_edit_link,url:c.onedrive_share_link,
  pdf_sha256:c.review_pdf_sha256,docx_sha256:c.review_docx_sha256,
  template_version:c.review_pdf_template_version,
  signer_name:c.signer_name,signer_title:c.signer_title,
  rent_entry_count:s.renderedPayload?.amounts_due?.length,
  total_amount_due:s.renderedPayload?.total_amount_due,
  attachments_required:c.attachments_required || [],
  attachments_need_separate_review:!!c.attachments_required?.length,
  saved:true,approved:false,signed:false,sent:false,notification_sent:false,
  requires_fresh_approval:true,
  presentation:'Tell Anne-Michelle: Your draft is ready for review. Show PDF first and Word for edits. It is unsigned and has not been sent. After editing Word, call prepare_case_notice_review again. Do not claim email/Slack delivery or approval. Keep any existing approval email and its controls; this receipt itself cannot approve or send.'}}];
