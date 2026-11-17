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
module.exports = {rentEvidence: rentEvidence};
