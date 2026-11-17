// A contemporaneous client statement can complete a month-only current order.
// A received date alone, an old notice, or a monthly rate cannot establish debt.
function rentEvidence({value,current,quote,support,source,orderDate}) {
  const text = [quote,...support.map(x=>x.quote)].join(' ');
  const amounts = s => [...String(s).matchAll(/\$\s*(\d[\d,]*(?:\.\d{1,2})?)|\b(\d[\d,]*\.\d{2})\b/g)]
    .map(m=>Math.round(Number((m[1]||m[2]).replace(/,/g,''))*100));
  const years = [...new Set(value.match(/\b20\d{2}\b/g)||[])];
  const knownYears = new Set(current.match(/\b20\d{2}\b/g)||[]);
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const months = [...new Set(monthNames.filter(m=>new RegExp('\\b'+m+'\\b','i').test(value)))];
  let period_basis = null;
  const rentQuotes=[quote,...support.map(e=>e.quote)].map(q=>String(q).split(/\bOn\s+(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b|\bFrom:\s/i)[0]);
  const quotedYear=y=>rentQuotes.some(q=>/\b(?:rent|unpaid|owed)\b/i.test(q)
    && (new RegExp('\\brent\\s+year\\s*[:=]?\\s*'+y+'\\b','i').test(q)
      || months.some(m=>new RegExp('\\b'+m+'\\b[^\\n.!?]{0,35}\\b'+y+'\\b|\\b'+y+'\\b[^\\n.!?]{0,15}\\b'+m+'\\b','i').test(q))));
  const missingYears = years.filter(y=>!quotedYear(y)&&!knownYears.has(y));
  if(missingYears.length) {
    const order = String(orderDate||'').match(/^(20\d{2})-(\d{2})-\d{2}/);
    const sent = String(source.date||'').match(/^(20\d{2})-(\d{2})-\d{2}/);
    const dayGap = (Date.parse(orderDate)-Date.parse(source.date))/86400000;
    const month = months.length===1 ? monthNames.indexOf(months[0])+1 : 0;
    const plain = String(source.text||'').replace(/<br\s*\/?>|<\/(?:p|div|tr|li)>/gi,'\n').replace(/<[^>]+>/g,' ');
    const newest = plain.split(/\bOn\s+(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b|\bFrom:\s|_{8,}/i)[0];
    const key = s=>String(s).toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    const isCurrentStatement = source.source_kind==='email' && order && sent && month
      && years.length===1 && years[0]===order[1] && sent[1]===order[1]
      && Number(order[2])===month && Number(sent[2])===month
      && dayGap>=-1 && dayGap<=31 && key(newest).includes(key(quote))
      && new RegExp('\\b'+months[0]+'\\b','i').test(quote)
      && new RegExp('\\b'+months[0]+'\\b','i').test(current)
      && /\brent\b/i.test(quote) && /\b(?:late|unpaid|not\s+paid|not\s+been\s+paid|hasn.t\s+been\s+paid|haven.t\s+(?:received|been\s+paid))\b/i.test(quote)
      && !/\b(?:last\s+year|previous|prior|historical|paid\s+in\s+full|already\s+paid|received\s+(?:the\s+)?rent)\b/i.test(quote);
    const recorded = amounts(current), proposed = amounts(value);
    if(!isCurrentStatement || !recorded.length || !proposed.length || proposed.some(v=>!recorded.includes(v)))
      return {error:'rent year requires explicit evidence or a matching current-month client statement and current order'};
    period_basis = {kind:'current-order-and-contemporaneous-unpaid-rent-statement',
      source_id:source.source_id,source_date:source.date,order_date:orderDate,month,year:Number(years[0])};
  }
  const supported = new Set(amounts(text));
  // Current order amounts remain authoritative; this only completes their year.
  if(period_basis) for(const v of amounts(current)) supported.add(v);
  const proposed = amounts(value);
  if(!proposed.length || proposed.some(v=>!supported.has(v))) return {error:'rent amount absent from verified quotations/current order'};
  // Month names also need quoted support; dates/headers never establish a month.
  if(months.some(m=>!new RegExp('\\b'+m+'\\b','i').test(text))) return {error:'unpaid month absent from verified quotations'};
  return {period_basis};
}

const ctx=$('Build Source Fact Request').first().json;
const raw=$json.choices?.[0]?.message?.content||'';let result;
try{const fenced=raw.match(/```(?:json)?\s*([\s\S]*?)```/i);result=JSON.parse(fenced?fenced[1]:raw.trim());}catch(e){throw Error('Case source research did not return readable evidence');}
const allowed=new Set(['Landlord Address',"Tenant's Property Address",'Landlords Preferred Phone','Tenant Name(s)',"Property Owner's Name",'First Name','Last Name','Signer Title','Itemized Rent Due','Case Name']);
const fields={...ctx.current_fields},provenance=[],rejected=[];
const norm=s=>String(s||'').normalize('NFKD').toLowerCase().replace(/<[^>]+>/g,' ').replace(/&quot;|&#34;/gi,'\"').replace(/&apos;|&#39;/gi,"'").replace(/&nbsp;|&#160;/gi,' ').replace(/[^a-z0-9]+/g,' ').trim();
const completeAddress=s=>/\b(?:CA|California)[,\s]+\d{5}(?:-\d{4})?\s*$/i.test(s);
for(const x of result.candidates||[]){
 const src=ctx.sources.find(s=>s.source_id===x.source_id),current=String(fields[x.field]||'').trim(),value=String(x.value||'').trim(),quote=String(x.quote||'');
 let rent_period_basis=null;
 const reject=reason=>rejected.push({field:x.field,reason});
 if(!allowed.has(x.field)||!value||!src||quote.length<8||!norm(src.text).includes(norm(quote))){reject('unsupported source quotation');continue;}
 if((result.conflicts||[]).some(c=>String(c).toLowerCase().includes(x.field.toLowerCase()))){reject('conflicting source');continue;}
 const isAddress=/Address$/.test(x.field);
 const support=(x.supporting_evidence||[]).filter(e=>ctx.sources.some(s=>s.source_id===e.source_id&&e.quote?.length>=8&&norm(s.text).includes(norm(e.quote))));
 const explicitCorrection=support.some(e=>{const m=ctx.sources.find(s=>s.source_id===e.source_id);return m.source_kind==='email'&&Date.parse(m.date)>=Date.parse(ctx.current_fields['Order Date']||'9999')&&[String(ctx.current_fields.Email||'').toLowerCase(),'info@blueprintevictions.com','orders@blueprintevictions.com'].includes(String(m.from||'').toLowerCase())&&/\b(?:please|correct|updated?|should|include|add|change)\b/i.test(e.quote);});
 const incomplete=isAddress?!completeAddress(current):x.field==='Itemized Rent Due'?(!/\b20\d{2}\b/.test(current)||/AG\$|\.SEP\$|\.OCT\./i.test(current)):!current;
 if(!incomplete&&!explicitCorrection){if(x.field!=='Case Name'&&norm(value)!==norm(current))result.conflicts=[...(result.conflicts||[]),'Source suggests a different '+x.field+' but no newer explicit correction is verified.'];reject('current confirmed field preserved');continue;}
 if(isAddress){
  if(!completeAddress(value)){reject('proposed address remains incomplete');continue;}
  const number=(current.match(/\b\d+/)||[])[0],newNumber=(value.match(/\b\d+/)||[])[0];
  const unit=s=>(s.match(/(?:\b(?:apt\.?|unit|suite|ste)\s*|#\s*)([a-z0-9-]+)/i)||[])[1]?.toLowerCase();
  if(number&&number!==newNumber||unit(current)&&unit(current)!==unit(value)){reject('current address number or unit changed');continue;}
  if(!norm(quote).includes(norm(value))){reject('complete address is not quoted');continue;}
 }else if(x.field==='Itemized Rent Due'){
  if((src.source_kind!=='email'&&!explicitCorrection)||!(/\b20\d{2}\b/.test(value))||!/(January|February|March|April|May|June|July|August|September|October|November|December)/i.test(value)){reject('unpaid rent not supported by current email');continue;}
  const checked=rentEvidence({value,current,quote,support,source:src,orderDate:ctx.current_fields['Order Date']});
  if(checked.error){reject(checked.error);continue;}
  rent_period_basis=checked.period_basis;
 }else if(x.field==='Tenant Name(s)'){
  if(value.split(/,|\band\b/i).map(s=>s.trim()).filter(Boolean).some(v=>!norm(quote).includes(norm(v)))){reject('tenant name absent from quotation');continue;}
 }else if(!norm(quote).includes(norm(value))){reject('value not present in quotation');continue;}
 fields[x.field]=value;provenance.push({field:x.field,value,previous:current,source_id:x.source_id,source_url:src.url,quote:x.quote,version:src.version||null,sha256:src.sha256||null,supporting_evidence:support,...(rent_period_basis?{period_basis:rent_period_basis}:{})});
}
const h=result.case_hold,hs=ctx.sources.find(s=>s.source_id===h?.source_id);
const case_hold=h?.active===true&&hs?.source_kind==='email'&&Date.parse(hs.date)>=Date.parse(ctx.current_fields['Order Date']||'9999')&&['info@blueprintevictions.com','orders@blueprintevictions.com','am@blueprintevictions.com'].includes(String(hs.from||'').toLowerCase())&&h.quote?.length>=8&&norm(hs.text).includes(norm(h.quote))?{active:true,reason:h.reason,source_id:hs.source_id,source_url:hs.url,quote:h.quote}:{active:false};
return [{json:{state:'case_source_research_complete',record_id:ctx.record_id,resolved_fields:fields,research_provenance:provenance,unresolved:result.unresolved||[],conflicts:result.conflicts||[],rejected,source_count:ctx.sources.length,source_asset_count:ctx.sources.filter(s=>s.source_kind==='case_document').length,mail_coverage:ctx.mail_coverage,case_hold,case_facts_written:false}}];
