const s = JSON.parse(JSON.stringify($input.first().json));
const c = s.context || {}, p = c.runtime_policy || {};
const command = String(s.reply_text || '').trim().replace(/[.!]+$/, '').toUpperCase();
if (s.revision_channel !== 'email' || c.review_format !== 'pdf' || p.enabled !== true ||
    !s.verifyOk || !s.tokenMatches || s.staleReview || s.timedOut ||
    s.recordId !== c.record_id || !s.replyMessageId || command !== 'REFRESH PDF' ||
    !p.reviewer_senders?.some(x => String(x).toLowerCase() === String(s.verifiedFrom || '').toLowerCase())) {
  throw Error('PDF refresh needs a verified reply to this case’s current review. Nothing approved.');
}
return [{json: {...s, reviewRevision: true, documentRevision: true,
  reviewReplyMessageId: s.replyMessageId, reminder_count: 0}}];
