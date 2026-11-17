function applyDispatchReceipts(review, approval, receipt) {
  const bound = x => x && x.record_id === review.record_id &&
    (x.folder_id || x.case_folder_id) === review.folder_id &&
    x.signature_request_id === review.signature_request_id &&
    (x.signed_sha256 || x.sha256) === review.signed_sha256 &&
    x.recipient === review.recipient && x.test_mode === review.test_only;
  const approved = bound(approval) && approval.schema === 'blueprint.dispatch.approval.v1' &&
    approval.approved === true && approval.approval === 'APPROVE THIS SIGNED VERSION FOR PROCESS SERVER';
  const sent = approved && bound(receipt) && receipt.delivery_verified === true &&
    !!receipt.sent_message_id && Number.isFinite(Date.parse(receipt.sent_at));
  return {...review, dispatch_approved: !!approved, sent: !!sent,
    state: sent ? (review.test_only ? 'test_delivery_verified' : 'delivery_verified') :
      approved ? 'dispatch_approved_delivery_not_verified' : review.state,
    delivery: sent ? receipt : null,
    process_server_order_placed: !!sent && !review.test_only,
    presentation: sent ? (review.test_only ?
      'The exact signed PDF was delivered to the owned test email. No real process-server order was placed.' :
      'Delivery of this exact signed PDF to the approved process server was verified.') :
      approved ? 'Dispatch approval is saved, but delivery is not verified. Reconcile the existing claim; do not send again blindly.' : review.presentation};
}
module.exports={applyDispatchReceipts};
