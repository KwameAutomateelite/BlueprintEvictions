// Mailbox eligibility is separate from document classification and case approval.
function hasProofIntent(message) {
  const text=String(message.subject||'')+' '+String(message.body?.content||message.bodyPreview||'').replace(/<[^>]+>/g,' ');
  return /\bproof\s+(?:of\s+)?service\b|\bp\.?o\.?s\.?\b|\bnotice\s+(?:was\s+)?served\b|\bservice\s+(?:completed|complete)\b/i.test(text);
}
function mayContainProof(message) {
  return message.hasAttachments===true||hasProofIntent(message);
}
function shouldNotifyProofProblem(result) {
  return result.silent!==true&&!['not_a_pos','bad_subject'].includes(result.code);
}
function shouldNotifyProofGateProblem(result,message) {
  return result.gate_ok===false&&result.code!=='not_a_pos'&&hasProofIntent(message);
}
module.exports={hasProofIntent,mayContainProof,shouldNotifyProofProblem,shouldNotifyProofGateProblem};
