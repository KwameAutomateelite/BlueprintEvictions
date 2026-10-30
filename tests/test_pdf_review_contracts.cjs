const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const dir = path.join(__dirname, '../integrations/n8n/exact-pdf-review');
const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
const hash = b => crypto.createHash('sha256').update(b).digest('hex');
const pdf = Buffer.from('%PDF-fictional-review-contract-fixture');
const sourceHash = hash(Buffer.from('fictional Word'));
const snapshot = () => ({fields:{Record_ID:'QA-ONLY'},renderedPayload:{notice_type:'residential',day_count:3,total_amount_due:'1000'},context:{record_id:'QA-ONLY',onedrive_file_id:'WORD-1',case_folder_id:'CASE-1',runtime_policy:{enabled:true,reviewer_email:'reviewer@example.invalid',reviewer_senders:['reviewer@example.invalid'],signing_enabled:false},attachments_required:['TPO_Required_Attachment.pdf']}});
const result = () => ({is_standard:true,source_docx_sha256:sourceHash,template_version:'residential-original-pdf-2026-10-27',pdf_base64:pdf.toString('base64'),pdf_sha256:hash(pdf),case_fields:{notice_date:'10/30/26',tenant_names:'Fictional Tenant',property_address_street:'123 Example St',property_address_city:'Sunnyvale, CA 94086',county:'SANTA CLARA',payment_recipient_name:'Fictional Owner',payment_phone:'(408) 555-0100',total_amount_due:'1250.25',amounts_due:[{due_date:'October 2026',amount:'1250.25'}],signer_name:'Pat Example',signer_title:'Property Manager'}});
async function run(file, input, refs={}, binary=pdf) {
  const code=fs.readFileSync(path.join(dir,file),'utf8');
  const $=name=>({first:()=>{if(!(name in refs))throw Error('Node not executed');return {json:refs[name]};},get item(){if(!(name in refs))throw Error('Node not executed');return {json:refs[name]};}});
  const helpers={getBinaryDataBuffer:async()=>binary,prepareBinaryData:async(b,name,mimeType)=>({data:b.toString('base64'),fileName:name,mimeType})};
  return new AsyncFunction('$input','$','require','Buffer',code).call({helpers},{first:()=>({json:input})},$,require,Buffer);
}

test('context uses the same bound case, Word file, folder and authorized reviewer',async()=>{
 const s=snapshot(); const x=(await run('validate-context.js',s))[0].json;
 assert.equal(x.document_id,'WORD-1');assert.equal(x.folder_id,'CASE-1');assert.equal(x.needs_preview,true);
 s.context.runtime_policy.reviewer_email='unverified@example.invalid';
 await assert.rejects(run('validate-context.js',s),/verified case/);
 s.context.runtime_policy.enabled=false;
 assert.equal((await run('validate-context.js',s))[0].json.needs_preview,false);
});

test('a Word file from a different case folder is rejected',async()=>{
 const refs={'Validate Review Context':{document_id:'WORD-1',folder_id:'CASE-1'}};
 const m={id:'WORD-1',parentReference:{id:'CASE-1'},name:'notice.docx',webUrl:'https://example.invalid/word',eTag:'v1'};
 assert.equal((await run('verify-word.js',m,refs))[0].json.id,'WORD-1');
 m.parentReference.id='DIFFERENT-CASE';await assert.rejects(run('verify-word.js',m,refs),/outside this case/);
});

test('PDF preparation binds the saved Word source, PDF bytes and template version',async()=>{
 const refs={'Prepare Saved Word Snapshot':{expected_sha256:sourceHash},'Validate Review Context':{snapshot:snapshot()}};
 const r=result();const out=(await run('prepare-pdf.js',r,refs))[0];
 assert.deepEqual(Buffer.from(out.binary.data.data,'base64'),pdf);
 assert.match(out.json.name,/3-Day Notice - 123 Example St - 2026-10-30 - REVIEW/);
 await assert.rejects(run('prepare-pdf.js',{...r,source_docx_sha256:'0'.repeat(64)},refs),/does not match/);
 await assert.rejects(run('prepare-pdf.js',{...r,pdf_base64:Buffer.from('tampered').toString('base64')},refs),/does not match/);
});

test('October 28 format is accepted without weakening Word/PDF hash checks', async()=>{
 const r=result();r.template_version='residential-original-pdf-2026-10-28-format3';
 const nodes={'Prepare Saved Word Snapshot':{expected_sha256:sourceHash},'Validate Review Context':{snapshot:snapshot()}};
 const out=await run('prepare-pdf.js',r,nodes);
 assert.equal(out[0].json.name,'3-Day Notice - 123 Example St - 2026-10-30 - REVIEW - '+sourceHash.slice(0,16)+'.pdf');
 const bad={...r,source_docx_sha256:'wrong'};
 await assert.rejects(run('prepare-pdf.js',bad,nodes),/does not match/);
});

test('saved reviewer snapshot keeps Word for edits and carries current printed values',async()=>{
 const s=snapshot();const r=result();
 const refs={'Verify Saved Word Location':{id:'WORD-1',eTag:'v1',webUrl:'https://example.invalid/word'},'Validate Review Context':{snapshot:s},'Render Saved Word as Exact PDF':r,'Verify Saved PDF Location':{id:'PDF-1',name:'readable-review.pdf',webUrl:'https://example.invalid/pdf'},'Recheck Saved Word Metadata':{id:'WORD-1',eTag:'v1'}};
 const out=(await run('finish-preview.js',{id:'CASE-1',folder:{},webUrl:'https://example.invalid/case-folder'},refs))[0].json;
 assert.equal(out.context.onedrive_file_id,'WORD-1');
 assert.equal(out.context.word_edit_link,'https://example.invalid/word');
 assert.equal(out.context.onedrive_share_link,'https://example.invalid/pdf');
 assert.equal(out.context.review_pdf_file_id,'PDF-1');
 assert.equal(out.context.signer_name,'Pat Example');
 assert.equal(out.renderedPayload.total_amount_due,'1250.25');
 assert.equal(out.fields.TOTAL_AMOUNT_DUE,'1250.25');
 assert.equal(out.fields.Record_ID,'QA-ONLY');
 assert.deepEqual(out.context.attachments_required,s.context.attachments_required);
 assert.equal(out.context.runtime_policy.signing_enabled,false);
 assert.equal(out.context.case_folder_url,'https://example.invalid/case-folder');
 refs['Recheck Saved Word Metadata'].eTag='v2';
 await assert.rejects(run('finish-preview.js',{id:'CASE-1',folder:{},webUrl:'https://example.invalid/case-folder'},refs),/Word changed/);
});

test('approval uses exactly the reviewed PDF and holds changed Word, PDF or missing review',async()=>{
 const r=result();const s=snapshot();Object.assign(s.context,{review_format:'pdf',review_pdf_file_id:'PDF-1',review_pdf_sha256:r.pdf_sha256,review_pdf_template_version:r.template_version});
 const refs={'Render Approved Saved Word as Exact PDF':r,'Build Approval Email':{reviewSnapshot:s},'Capture Approved Word Snapshot':{docx_sha256:sourceHash}};
 const approved=(await run('approved-pdf.js',{},refs))[0];
 assert.deepEqual(Buffer.from(approved.binary.data.data,'base64'),pdf);
 await assert.rejects(run('approved-pdf.js',{},refs,Buffer.from('changed shared PDF')),/fresh PDF review/i);
 await assert.rejects(run('approved-pdf.js',{}, {...refs,'Render Approved Saved Word as Exact PDF':{...r,pdf_sha256:'0'.repeat(64)}}),/fresh PDF review/i);
 await assert.rejects(run('approved-pdf.js',{}, {...refs,'Capture Approved Word Snapshot':{docx_sha256:'0'.repeat(64)}}),/fresh PDF review/i);
 const noReview=snapshot();
 await assert.rejects(run('approved-pdf.js',{}, {...refs,'Build Approval Email':{reviewSnapshot:noReview}}),/fresh PDF review/i);
});

test('reviewer permission verification never treats an unrelated person as the reviewer',async()=>{
 const refs={'Validate Review Context':{reviewer:'reviewer@example.invalid'}};
 assert.equal((await run('verify-access.js',{value:[{roles:['read'],invitation:{email:'reviewer@example.invalid'}}]},refs))[0].json.readable,true);
 await assert.rejects(run('verify-access.js',{value:[{roles:['read'],invitation:{email:'other@example.invalid'}}]},refs),/cannot open/);
});

test('manual Word refresh requires an authenticated current-case reviewer reply',async()=>{
 const s=snapshot();Object.assign(s.context,{review_format:'pdf'});
 Object.assign(s,{recordId:'QA-ONLY',revision_channel:'email',reply_text:'refresh pdf!',verifyOk:true,tokenMatches:true,
   staleReview:false,timedOut:false,replyMessageId:'FICTIONAL-REPLY',verifiedFrom:'reviewer@example.invalid'});
 const out=(await run('refresh-request.js',s))[0].json;
 assert.equal(out.context.onedrive_file_id,'WORD-1');assert.equal(out.documentRevision,true);
 assert.equal(out.reviewReplyMessageId,'FICTIONAL-REPLY');
 for(const patch of [{verifyOk:false},{tokenMatches:false},{staleReview:true},{timedOut:true},
   {recordId:'OTHER-CASE'},{verifiedFrom:'other@example.invalid'},{reply_text:'YES'},{replyMessageId:null}]) {
   await assert.rejects(run('refresh-request.js',{...s,...patch}),/verified reply/);
 }
 s.context.runtime_policy.enabled=false;
 await assert.rejects(run('refresh-request.js',s),/verified reply/);
});

test('QC filenames identify the case/address/date, sanitize path characters and retain version identity',async()=>{
 const r=result();const s=snapshot();s.context.case_name='Example Owner v Tenant';
 s.context.runtime_policy.test_mode=true;
 const refs={'Prepare Saved Word Snapshot':{expected_sha256:sourceHash},'Validate Review Context':{snapshot:s}};
 const a=(await run('prepare-pdf.js',r,refs))[0];
 assert.match(a.json.name,/^TEST ONLY - 3-Day Notice - Example Owner v Tenant - 123 Example St - 2026-10-30 - REVIEW - /);
 assert.ok(a.json.name.endsWith(sourceHash.slice(0,16)+'.pdf'));
 assert.deepEqual(Buffer.from(a.binary.data.data,'base64'),pdf);
 s.context.case_name='../Case <name>: / \\ ? * | '+ 'Very Long Name '.repeat(40);
 r.case_fields.property_address_street='123 Example St / Unit 2';
 const b=(await run('prepare-pdf.js',r,refs))[0];
 assert.ok(b.json.name.length<=180);assert.match(b.json.name,/2026-10-30 - REVIEW/);assert.doesNotMatch(b.json.name,/[<>:"/\\|?*\x00-\x1f]/);
 assert.ok(b.json.name.endsWith(sourceHash.slice(0,16)+'.pdf'));
 assert.equal((await run('prepare-pdf.js',r,refs))[0].json.name,b.json.name);
});

test('a folder link from a different case cannot be returned as verified',async()=>{
 const refs={'Verify Saved Word Location':{id:'WORD-1',eTag:'v1'},'Recheck Saved Word Metadata':{id:'WORD-1',eTag:'v1'},'Validate Review Context':{snapshot:snapshot()}};
 await assert.rejects(run('finish-preview.js',{id:'OTHER-CASE',folder:{},webUrl:'https://example.invalid/other'},refs),/folder link could not be verified/);
});
