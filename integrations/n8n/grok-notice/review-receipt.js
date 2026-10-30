const s=$input.first().json,c=s.context || {};
if (c.review_format!=='pdf' || !c.review_pdf_file_id || !c.word_edit_link || !c.case_folder_url ||
    !/^[a-f0-9]{64}$/.test(c.review_pdf_sha256 || '') ||
    !/^[a-f0-9]{64}$/.test(c.review_docx_sha256 || '')) {
  throw Error('Saved Word is not the standardized residential form. No ready PDF was confirmed.');
}
return [{json:{state:'draft_ready_for_review',case_id:c.record_id,
  folder_id:c.case_folder_id,item_id:c.onedrive_file_id,version:c.word_version,request_id:c.request_id || null,
  folder_url:c.case_folder_url,case_name:c.case_name || '',property_address:c.property_address,
  pdf_item_id:c.review_pdf_file_id,pdf_url:c.onedrive_share_link,
  pdf_name:c.review_pdf_name,word_name:c.word_edit_name,
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
  presentation:'Tell Anne-Michelle the case and property address, then show three clickable links: REVIEW PDF (pdf_url), EDIT WORD (word_edit_url), OPEN ADDRESS FOLDER (folder_url). Use the actual returned URLs and readable filenames. This is a review copy, not an approved, signed or sent notice. After editing Word, call prepare_case_notice_review for a fresh PDF and fresh approval. Keep existing approval controls; this receipt cannot approve or send.'}}];
