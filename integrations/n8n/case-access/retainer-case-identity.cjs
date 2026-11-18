// Resolve identity using the full premises, including unit, and current
// reciprocal Airtable links. Never select by name or newest date alone.
function address(value) {
  const aliases={st:'street',ave:'avenue',av:'avenue',rd:'road',dr:'drive',blvd:'boulevard',ln:'lane',ct:'court',pl:'place',apt:'unit',apartment:'unit',ste:'suite'};
  let s=String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/^\s*\(tpo\)\s*/i,'').toLowerCase().replace(/#/g,' unit ').replace(/[^a-z0-9\s]/g,' ').trim().replace(/\s+/g,' ');
  const m=s.match(/\s+(\d{5})(?:\s+\d{4})?$/),zip=m?.[1]||'';
  if(m)s=s.slice(0,m.index).trim();
  s=s.replace(/\s+(ca|california)$/,'');
  return {body:s.split(' ').filter(Boolean).map(t=>aliases[t]||t).join(' ').replace(/\bunit unit\b/g,'unit'),zip};
}
function sameAddress(a,b){const x=address(a),y=address(b);return !!x.body&&x.body===y.body&&(!x.zip||!y.zip||x.zip===y.zip);}
function unitParts(value){
  const a=address(value),pattern=/\b(?:unit|suite)\s+[a-z0-9-]+\b|\badu\b/g,matches=[...a.body.matchAll(pattern)];
  return {...a,unit:matches[0]?.[0]||null,ambiguous:matches.length>1,base:a.body.replace(pattern,' ').replace(/\s+/g,' ').trim()};
}
function sameBaseAddress(a,b){const x=unitParts(a),y=unitParts(b);return !x.ambiguous&&!y.ambiguous&&!!x.base&&x.base===y.base&&(!x.zip||!y.zip||x.zip===y.zip);}
// An added unit may refine an incomplete legacy address only when both current
// tables identify exactly one reciprocal case pair. Never ignore a conflicting unit.
function linkedUnitRefinement(v,notices,evictions){
  const norm=s=>String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const document=unitParts(v.property_address);if(!document.unit||document.ambiguous)return null;
  const matches=r=>norm(r.fields?.['Case Name'])===norm(v.case_name)&&sameBaseAddress(r.fields?.["Tenant's Property Address"],v.property_address);
  const ns=(notices.records||[]).filter(matches),es=(evictions.records||[]).filter(matches);
  if(ns.length!==1||es.length!==1)return null;
  const n=ns[0],e=es[0];
  if(unitParts(n.fields["Tenant's Property Address"]).unit||unitParts(e.fields["Tenant's Property Address"]).unit)return null;
  if(!/^[A-Za-z0-9!_-]+$/.test(n.fields['Case Folder ID']||'')||!(n.fields['ELG Evictions']||[]).includes(e.id)||!(e.fields['Preceding Notice']||[]).includes(n.id))return null;
  return {notice:n,eviction:e.id};
}
function selectRetainerCase(v,notices,evictions,test=false,testBindings={}){
  if(notices.offset||evictions.offset)throw Error('Case matches truncated; no filing.');
  const norm=s=>String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const matches=r=>norm(r.fields?.['Case Name'])===norm(v.case_name)&&sameAddress(r.fields?.["Tenant's Property Address"],v.property_address);
  const candidates=(notices.records||[]).filter(matches).filter(r=>/^[A-Za-z0-9!_-]+$/.test(r.fields?.['Case Folder ID']||''));
  const linked=[];
  for(const e of (evictions.records||[]).filter(matches)){
    const ids=e.fields?.['Preceding Notice']||[];
    for(const n of candidates)if(ids.includes(n.id)&&(n.fields?.['ELG Evictions']||[]).includes(e.id))linked.push({notice:n,eviction:e.id});
  }
  let chosen,eviction_id=null,unit_refinement=false;
  if(linked.length===1){chosen=linked[0].notice;eviction_id=linked[0].eviction;}
  else if(linked.length>1)throw Error('Multiple linked eviction cases match; no filing.');
  else if(candidates.length===1)chosen=candidates[0];
  else if(candidates.length>1)throw Error('Multiple notice folders match without a unique linked eviction; no filing.');
  else{
    const refined=linkedUnitRefinement(v,notices,evictions);
    if(refined){chosen=refined.notice;eviction_id=refined.eviction;unit_refinement=true;}
    else{
      const direct=(evictions.records||[]).filter(matches).filter(r=>/^[A-Za-z0-9!_-]+$/.test(r.fields?.['Case Folder ID']||''));
      if(direct.length===1)chosen={...direct[0],table:'tblSZE5jXc4zaqVAh'};
      else throw Error('No unique case with a verified folder binding; no filing.');
    }
  }
  if(test){
    const b=testBindings[chosen.id];
    if(b){if(norm(chosen.fields['Case Name'])!==norm(b.case_name)||!sameAddress(chosen.fields["Tenant's Property Address"],b.address)||chosen.fields['Case Folder ID']!==b.folder_id||chosen.fields["Property Owner's Name"]!==b.owner)throw Error('Owned retainer test binding changed; no filing.');}
    else if(chosen.id!=='recxLwYEXtb2ZrWen')throw Error('Test email cannot access a live case.');
  }
  return {record_id:chosen.id,table:chosen.table||'tblZbUGz8OTvFNh9i',folder_id:chosen.fields['Case Folder ID'],case_name:chosen.fields['Case Name'],address:chosen.fields["Tenant's Property Address"],document_address:v.property_address,unit_refinement,linked_eviction_id:eviction_id,identity_basis:unit_refinement?'unique_reciprocal_case_and_added_document_unit':eviction_id?'full_premises_and_reciprocal_eviction_notice_link':'unique_full_premises_bound_case'};
}
if(typeof module!=='undefined')module.exports={selectRetainerCase,sameAddress,sameBaseAddress,linkedUnitRefinement};
