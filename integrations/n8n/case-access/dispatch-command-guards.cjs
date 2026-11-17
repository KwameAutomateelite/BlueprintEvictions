function validateDispatchAction(a,c,m,file,sha){
 const f=c.fields||{};const to=m.test_mode?c.policy.process_server_override:c.intended_to;
 if(a.action!=='approve'||a.approval!=='APPROVE THIS SIGNED VERSION FOR PROCESS SERVER'||typeof a.approval_quote!=='string'||!a.approval_quote.trim()||/\b(?:do not|don't|not approved|cancel|decline)\b/i.test(a.approval_quote))throw Error('Explicit approval of this signed version is required');
 if(!c.policy.reviewer_senders?.includes(a.reviewer_email))throw Error('Business reviewer identity required');
 if(f.Status!=='With Process Server (N)'||f['Dispatched On'])throw Error('Case is not awaiting dispatch approval or is already dispatched');
 if(m.signature_status!=='complete'||m.record_id!==c.id||m.folder_id!==f['Case Folder ID']||m.test_mode!==c.policy.test_mode||file.id!==m.file_id||file.parentReference?.id!==m.folder_id||!file.file||sha!==m.sha256)throw Error('Verified current signed package required');
 if(a.signature_request_id!==m.signature_request_id||a.signed_sha256!==m.sha256||a.recipient!==String(to).toLowerCase()||a.test_mode!==m.test_mode)throw Error('The approved signature, PDF, recipient or test/live mode changed; review again');
 if(!/^[-A-Za-z0-9_.:]{8,120}$/.test(a.request_id||''))throw Error('Stable approval request ID required');
 if(c.route!=='email')throw Error('This county requires a manual portal order; no paid order was placed');
 return true;
}
module.exports={validateDispatchAction};