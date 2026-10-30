const test=require('node:test'),assert=require('node:assert/strict'),fs=require('fs'),path=require('path');
const P=path.join(__dirname,'../integrations/n8n/grok-notice');
const {caseContext,resolveRentYear,noticePayload}=require(path.join(P,'case-inputs.cjs'));
const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
function record(){return {id:'rec12345678901234',fields:{'Case Name':'Fictional Owner v Fictional Tenant','Notice Type':'3-day to pay or quit',
 'Tenant Name(s)':'Fictional Tenant',"Tenant's Property Address":'123 Example Lane, Cupertino, CA 95014',
 "Property Owner's Name":'Fictional Owner','Landlord Address':'456 Sample Ave, Cupertino, CA 95014',
 'Landlords Preferred Phone':'4085550100','First Name':'Fictional','Last Name':'Signer','Signer Title':'Property Manager',
 'Rent':1000,'Rent Due Day':'1','Itemized Rent Due':months.map(m=>m+' 2025 $1,000.00').join('; ')}};}
function prepare(rec){const policy={enabled:true,reviewer_email:'reviewer@example.com',reviewer_senders:['reviewer@example.com']};
 const values={'Read Bound Case':rec,'Load Current Reviewer Policy':{_noticePolicy:policy},'Validate Case Tool':{request_id:'fictional-test-12'},'Require Case Folder ID':{folder_id:'FICTIONAL_FOLDER'}};
 return new Function('$','caseContext','resolveRentYear','noticePayload',fs.readFileSync(path.join(P,'prepare-draft.js'),'utf8'))(
  name=>({first:()=>({json:values[name]})}),caseContext,resolveRentYear,noticePayload)[0].json;}
test('current case parser supplies all twelve rent entries and a separate designated signer',()=>{
 const result=prepare(record());assert.equal(result.payload.amounts_due.length,12);assert.equal(result.payload.total_amount_due,'$12,000.00');
 assert.equal(result.payload.signer_name,'Fictional Signer');assert.equal(result.payload.landlord_name,'Fictional Owner');
 assert.equal(result.payload.signer_title,'Property Manager');assert.equal(result.payload.service_date,undefined);
});
test('missing signer stops before file creation, never substitutes owner',()=>{
 const r=record();delete r.fields['Signer Title'];assert.throws(()=>prepare(r),/designated signer role/);
 delete r.fields['First Name'];assert.throws(()=>prepare(r),/designated signer first name/);
});
test('missing debt year cannot silently use this year',()=>{
 const r=record();r.fields['Itemized Rent Due']='January $1000; February $1000';assert.throws(()=>prepare(r),/year/);
});
test('stated total conflict holds instead of publishing an inaccurate notice',()=>{
 const r=record();r.fields['Itemized Rent Due']+='; Total $99';assert.throws(()=>prepare(r),/Draft needs/);
});
test('other notice families cannot enter standardized residential draft tool',()=>{
 for(const value of ['Commercial 3-day to pay or quit','3-day perform covenants or quit','30-day termination','5-day pay rent or quit']){
 const r=record();r.fields['Notice Type']=value;assert.throws(()=>prepare(r),/supports residential/);}
});
test('ready receipt requires saved PDF and actual content hashes',()=>{
 const code=fs.readFileSync(path.join(P,'review-receipt.js'),'utf8');const run=s=>new Function('$input',code)({first:()=>({json:s})});
 assert.throws(()=>run({context:{}}),/No ready PDF/);
 const out=run({context:{record_id:'rec12345678901234',onedrive_file_id:'WORD',review_pdf_file_id:'PDF',word_version:'etag1',
 case_folder_url:'https://example.com/address-folder',review_format:'pdf',onedrive_share_link:'https://example.com/review.pdf',word_edit_link:'https://example.com/edit.docx',
 review_pdf_sha256:'a'.repeat(64),review_docx_sha256:'b'.repeat(64)},renderedPayload:{amounts_due:Array(12).fill({})}})[0].json;
 assert.equal(out.state,'draft_ready_for_review');assert.equal(out.url,'https://example.com/review.pdf');assert.equal(out.rent_entry_count,12);
 for(const key of ['approved','signed','sent','notification_sent'])assert.equal(out[key],false);
 assert.equal(out.requires_fresh_approval,true);
});

test('QC receipt requires the verified address-folder link and returns three distinct destinations',()=>{
 const code=fs.readFileSync(path.join(P,'review-receipt.js'),'utf8');const run=s=>new Function('$input',code)({first:()=>({json:s})});
 const c={record_id:'rec12345678901234',case_name:'Example v Example',property_address:'123 Example St',case_folder_id:'CASE',
 onedrive_file_id:'WORD',review_pdf_file_id:'PDF',word_version:'etag1',review_format:'pdf',
 onedrive_share_link:'https://example.invalid/review.pdf',word_edit_link:'https://example.invalid/edit.docx',case_folder_url:'https://example.invalid/case-folder',
 review_pdf_sha256:'a'.repeat(64),review_docx_sha256:'b'.repeat(64),review_pdf_name:'Case - REVIEW.pdf',word_edit_name:'draft.docx'};
 const receipt=run({context:c})[0].json;
 assert.equal(receipt.pdf_url,c.onedrive_share_link);assert.equal(receipt.word_edit_url,c.word_edit_link);assert.equal(receipt.folder_url,c.case_folder_url);
 assert.equal(receipt.pdf_name,'Case - REVIEW.pdf');assert.match(receipt.presentation,/three clickable links/);
 delete c.case_folder_url;assert.throws(()=>run({context:c}),/No ready PDF/);
});
