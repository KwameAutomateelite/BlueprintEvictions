// Read supporting files whenever an actual case fact needs research. A fully
// populated current order with one year-less rent month only needs current mail.
const x=$json,f=x.fields||{};
if(!/^rec[A-Za-z0-9]{14}$/.test(x.record_id||'')||!x.folder_id||!x.listing_complete)throw Error('Complete current case inventory is required for research');
if(/^(?:Served|Archived|Closed|Cancelled|Expired|Continue)/i.test(String(f.Status||'')))throw Error('Existing served or historical case is read-only; no automatic redraft');
if(String(f['Notice Type']||'').trim()!=='3-Day Notice to Pay Rent or Quit')throw Error('Approved pay-rent-or-quit template does not cover this notice type');
const raw=String(f['Itemized Rent Due']||'');
const month='(?:January|February|March|April|May|June|July|August|September|October|November|December)';
const simple=new RegExp('^\\s*'+month+"(?:['’]s?)?\\s*(?:rent\\s*)?[:=]?\\s*\\$\\s*\\d[\\d,]*(?:\\.\\d{1,2})?\\s*$",'i').test(raw);
const has=k=>String(f[k]||'').trim().length>0;
const address=k=>has(k)&&/^\s*\d/.test(f[k])&&/\b(?:CA|California)[,\s]+\d{5}(?:-\d{4})?\s*$/i.test(f[k]);
const complete=address("Tenant's Property Address")&&address('Landlord Address')
  &&['Tenant Name(s)',"Property Owner's Name",'First Name','Last Name','Signer Title'].every(has)
  &&String(f['Landlords Preferred Phone']||f['Phone Number']||'').replace(/\D/g,'').length>=10;
const rent_only_gap=simple&&complete&&/^20\d{2}-\d{2}-\d{2}/.test(String(f['Order Date']||''));
return [{json:{...x,rent_only_gap,research_path:rent_only_gap?'current-order-and-current-mail':'case-assets-and-current-mail'}}];
