const assert=require('node:assert/strict');
const {applyDispatchReceipts}=require('./dispatch-review-receipts.cjs');
const {validateDispatchAction}=require('./dispatch-command-guards.cjs');
const review={record_id:'owned-case',folder_id:'case-folder',signature_request_id:'signature',signed_sha256:'a'.repeat(64),recipient:'owner@example.test',test_only:true,state:'signed_awaiting_dispatch_approval'};
const approval={schema:'blueprint.dispatch.approval.v1',approved:true,record_id:review.record_id,folder_id:review.folder_id,signature_request_id:review.signature_request_id,signed_sha256:review.signed_sha256,recipient:review.recipient,test_mode:true,approval:'APPROVE THIS SIGNED VERSION FOR PROCESS SERVER'};
const receipt={record_id:review.record_id,case_folder_id:review.folder_id,signature_request_id:review.signature_request_id,sha256:review.signed_sha256,recipient:review.recipient,test_mode:true,delivery_verified:true,sent_message_id:'sent-message',sent_at:'2026-11-17T22:38:00Z'};
assert.equal(applyDispatchReceipts(review,approval,receipt).state,'test_delivery_verified');
assert.equal(applyDispatchReceipts(review,approval,receipt).process_server_order_placed,false);
assert.equal(applyDispatchReceipts(review,null,receipt).sent,false);
assert.equal(applyDispatchReceipts(review,approval,null).state,'dispatch_approved_delivery_not_verified');
for(const [key,value]of Object.entries({record_id:'another-case',case_folder_id:'another-folder',signature_request_id:'old-signature',sha256:'b'.repeat(64),recipient:'other@example.test',test_mode:false,sent_message_id:null,sent_at:'invalid',delivery_verified:false})){
 assert.equal(applyDispatchReceipts(review,approval,{...receipt,[key]:value}).sent,false,key);
}
const caseRecord={id:review.record_id,fields:{Status:'With Process Server (N)','Case Folder ID':review.folder_id},policy:{reviewer_senders:['reviewer@example.test'],process_server_override:review.recipient,test_mode:true},route:'email',intended_to:'server@example.test'};
const manifest={signature_status:'complete',record_id:review.record_id,folder_id:review.folder_id,signature_request_id:review.signature_request_id,file_id:'signed-file',sha256:review.signed_sha256,test_mode:true};
const file={id:'signed-file',parentReference:{id:review.folder_id},file:{}};
const action={action:'approve',approval:'APPROVE THIS SIGNED VERSION FOR PROCESS SERVER',approval_quote:'Send this exact signed test notice to my owned test email.',reviewer_email:'reviewer@example.test',signature_request_id:review.signature_request_id,signed_sha256:review.signed_sha256,recipient:review.recipient,test_mode:true,request_id:'owned-dispatch-acceptance'};
assert(validateDispatchAction(action,caseRecord,manifest,file,manifest.sha256));
for(const [key,value]of Object.entries({approval:'yes',approval_quote:'Do not send this notice',reviewer_email:'unknown@example.test',signature_request_id:'old',signed_sha256:'b'.repeat(64),recipient:'server@example.test',test_mode:false,request_id:''})){
 assert.throws(()=>validateDispatchAction({...action,[key]:value},caseRecord,manifest,file,manifest.sha256),key);
}
assert.throws(()=>validateDispatchAction(action,{...caseRecord,route:'card'},manifest,file,manifest.sha256));
assert.throws(()=>validateDispatchAction(action,{...caseRecord,fields:{...caseRecord.fields,'Dispatched On':'2026-11-17'}},manifest,file,manifest.sha256));
assert.throws(()=>validateDispatchAction(action,caseRecord,manifest,{...file,parentReference:{id:'wrong-folder'}},manifest.sha256));
console.log('Signed dispatch approval, receipt binding and truthful test-delivery state checks passed.');
