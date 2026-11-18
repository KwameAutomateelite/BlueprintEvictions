function validateRetainerEvidence(value) {
  if(!value||typeof value!=='object'||Array.isArray(value)||!['retainer','notice','other'].includes(value.document_type))throw Error('PDF classification unreadable');
  if(value.document_type!=='retainer')return {retainer:false,reason:'not-a-retainer'};
  const c=value.confidence;
  if(typeof c!=='number'||!Number.isFinite(c)||c<0||c>100)throw Error('Signed retainer confidence unreadable');
  const percent=c<=1?c*100:c;
  if(value.signed!==true||percent<95)throw Error('Signed retainer identity needs review');
  if(typeof value.case_name!=='string'||!value.case_name.trim()||value.case_name.length>200||typeof value.property_address!=='string'||!value.property_address.trim()||value.property_address.length>300||typeof value.evidence!=='string'||!value.evidence.trim()||value.evidence.length>10000)throw Error('Signed retainer identity or evidence unreadable');
  return {...value,retainer:true,confidence_pct:percent};
}
if(typeof module!=='undefined')module.exports={validateRetainerEvidence};
