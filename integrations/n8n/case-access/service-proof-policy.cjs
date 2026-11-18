function premises(value) {
 const aliases={st:'street',ave:'avenue',av:'avenue',rd:'road',dr:'drive',blvd:'boulevard',ln:'lane',ct:'court',pl:'place',ter:'terrace',cir:'circle',pkwy:'parkway',apt:'unit',apartment:'unit',ste:'suite',california:'ca'};
 let text=String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/^\s*\(tpo\)\s*/i,'').toLowerCase().replace(/#/g,' unit ').replace(/[^a-z0-9\s-]/g,' ').trim().replace(/\s+/g,' ');
 const match=text.match(/\s+(\d{5})(?:-\d{4})?$/),zip=match?.[1]||'';
 if(match)text=text.slice(0,match.index).trim();
 text=text.split(/\s+/).filter(Boolean).map(x=>aliases[x]||x).join(' ').replace(/\bunit unit\b/g,'unit');
 return {text,zip};
}
function samePremises(a,b) {
 const x=premises(a),y=premises(b);
 return !!x.text&&x.text===y.text&&(!x.zip||!y.zip||x.zip===y.zip);
}
function caseName(value) {
 return String(value||'').toLowerCase().replace(/\bvs?\.?\b/g,'v').replace(/[^a-z0-9]/g,'');
}
function selectServiceCase(records,name,address) {
 if(!caseName(name)||!premises(address).text)return {ok:false,reason:'missing_case_identity',candidates:[]};
 const candidates=(records||[]).filter(r=>caseName(r.caseName||r.fields?.['Case Name'])===caseName(name)&&samePremises(r.address||r.fields?.["Tenant's Property Address"],address));
 if(candidates.length!==1)return {ok:false,reason:candidates.length?'ambiguous':'no_match',candidates};
 return {ok:true,record:candidates[0]};
}
function validServiceDate(value) {
 if(typeof value!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(value))return false;
 const d=new Date(value+'T00:00:00Z');return Number.isFinite(d.getTime())&&d.toISOString().slice(0,10)===value;
}
function supportedCureDeadline(type) {
 const t=String(type||'').replace(/[()-]/g,' ');
 return /\b3\s+day\b/i.test(t)&&/\b(?:pay|perform|correct|cure)\b/i.test(t)&&/\bor\s+quit\b/i.test(t)&&!/\bunconditional\b/i.test(t);
}
function serviceSnapshot(record) {
 const f=record.fields||{};return {record_id:record.id,fields:Object.fromEntries(['Case Folder ID','Case Name',"Tenant's Property Address",'Notice Type','Status','Notice Served','Notice Expires','Notice F/U Date','Latest Notice Link'].map(k=>[k,f[k]??null]))};
}
function verifyServiceApproval(approval,proposal,record,policy) {
 if(!approval||approval.action!=='approve'||approval.approval!=='APPROVE THIS PROOF OF SERVICE'||!policy?.reviewer_senders?.includes(approval.reviewer_email)||typeof approval.approval_quote!=='string'||!approval.approval_quote.trim()||/\b(?:do not|don't|hold|reject|cancel)\b/i.test(approval.approval_quote))throw Error('Explicit authorized proof-of-service approval is required');
 if(approval.proposal_sha256!==proposal.proposal_sha256||approval.proof_sha256!==proposal.proof_sha256||approval.served_date!==proposal.served_date||approval.test_mode!==proposal.test_mode||record.id!==proposal.record_id||record.fields?.['Case Folder ID']!==proposal.folder_id)throw Error('Approved proof, date, case, folder or mode changed');
 if(!validServiceDate(proposal.served_date)||!Number.isFinite(Number(proposal.confidence))||Number(proposal.confidence)<70||Number(proposal.confidence)>100)throw Error('A valid evidence-backed service date is required');
 if(JSON.stringify(serviceSnapshot(record))!==JSON.stringify(proposal.case_snapshot))throw Error('Case changed after review; prepare a fresh proof card');
 return true;
}
function verifiedProofResumeUrl(proposal) {
 const id=String(proposal?.execution_id||'');
 const expected='https://blueprintevictions.app.n8n.cloud/webhook-waiting/'+id;
 if(!/^\d+$/.test(id)||typeof proposal?.resume_url!=='string'||!(proposal.resume_url===expected||proposal.resume_url.startsWith(expected+'?')))throw Error('Proof resume route invalid');
 return proposal.resume_url;
}
function validateServiceClientUpdate(proposal) {
 if(proposal?.write_mode!=='full'||!supportedCureDeadline(proposal.notice_type))throw Error('The deadline for this notice needs office review before a client update.');
 if(proposal.service_method!=='posting and mailing')throw Error('The existing client email describes posting and mailing. Office review is needed before another service method can change status.');
 return true;
}
module.exports={premises,samePremises,selectServiceCase,validServiceDate,supportedCureDeadline,serviceSnapshot,verifyServiceApproval,verifiedProofResumeUrl,validateServiceClientUpdate};
