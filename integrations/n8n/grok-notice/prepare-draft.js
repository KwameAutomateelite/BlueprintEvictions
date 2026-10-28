const rec = $('Read Bound Case').first().json;
const loaded = $('Load Current Reviewer Policy').first().json;
const fields = {...rec.fields, id:rec.id, _noticePolicy:loaded._noticePolicy};
if (!loaded._noticePolicy?.enabled) throw Error('Notice creation is held by the current reviewer policy.');
const type = String(fields['Notice Type'] || '').toLowerCase();
if (/commercial|perform|covenant|termination|warning/.test(type) ||
    !/(?:\b3[ -]?day|\bthree\b)/.test(type)) {
  throw Error('This tool supports residential three-day pay-rent-or-quit notices only.');
}
let ctx = resolveRentYear(caseContext(fields),rec);
const issues = [];
const need = (v,label) => {if (!String(v || '').trim()) issues.push(label);};
need(ctx.tenant.names,'tenant names');
need(ctx.property.address,'property street address');
need(ctx.property.city_state_zip,'property city, state and ZIP');
need(ctx.property.county,'county');
need(ctx.landlord.full_name,'property owner / payment recipient');
need(ctx.landlord.address,'housing provider / payment address');
need(ctx.landlord.phone,'housing provider phone');
need(fields['First Name'],'designated signer first name');
need(fields['Last Name'],'designated signer last name');
need(fields['Signer Title'],'designated signer role');
if (ctx.financials.money?.status !== 'ok') issues.push(ctx.financials.money?.question || 'unambiguous unpaid rent entries');
const rows=ctx.financials.itemized?.rows || [];
if (!rows.length || rows.some(r=>!r.isMonth || !r.month?.year)) {
  issues.push('explicit unpaid periods and years; no guessed rent dates');
}
if (ctx.property.county_unknown) issues.push('confirmed county');
if (issues.length) throw Error('Draft needs these case details: '+[...new Set(issues)].join('; ')+'. No notice was created.');
const payload=noticePayload(ctx);
payload.notice_date=new Intl.DateTimeFormat('en-US',{timeZone:'America/Los_Angeles',month:'long',day:'numeric',year:'numeric'}).format(new Date());
if (payload._notice_input_issues?.length) throw Error(payload._notice_input_issues.join('; '));
const request=$('Validate Case Tool').first().json;
return [{json:{payload,record_id:rec.id,folder_id:$('Require Case Folder ID').first().json.folder_id,
  policy:loaded._noticePolicy,request_id:request.request_id,
  name:'3-Day Notice - Grok Draft - '+request.request_id+'.docx'}}];
