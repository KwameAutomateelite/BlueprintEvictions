const {test}=require('node:test');const assert=require('node:assert/strict');const {createHash}=require('node:crypto');
const hash=b=>createHash('sha256').update(b).digest('hex');
test('edit normalization retains original source binding and verifies corrected bytes',async()=>{
 const {normalizeWordLayout}=await import('../revision-service/normalize-word.mjs');
 const old=Buffer.from('PKold'),fixed=Buffer.from('PKcorrected');
 const input={source_sha256:'a'.repeat(64),docx_sha256:hash(old),docx_base64:old.toString('base64'),applied_changes:1};
 const result=await normalizeWordLayout(input,{fetchImpl:async(url,request)=>{
  assert.equal(url,'https://web-production-56acb.up.railway.app/normalize-saved-notice-word');
  assert.equal(JSON.parse(request.body).expected_sha256,hash(old));
  return {ok:true,json:async()=>({is_standard:true,source_docx_sha256:hash(old),word_template_version:'residential-pay-or-quit-2026-10-30-word-layout4',docx_base64:fixed.toString('base64'),docx_sha256:hash(fixed),printed_signer:{name:'Test Owner',title:'Property Owner'}})};
 }});
 assert.equal(result.source_sha256,input.source_sha256);assert.equal(result.docx_sha256,hash(fixed));assert.equal(result.applied_changes,1);
});
test('unavailable normalizer or altered response cannot save an unverified Word',async()=>{
 const {normalizeWordLayout}=await import('../revision-service/normalize-word.mjs');
 const source=Buffer.from('PKold'),input={docx_sha256:hash(source),docx_base64:source.toString('base64')};
 await assert.rejects(()=>normalizeWordLayout(input,{fetchImpl:async()=>({ok:false})}));
 await assert.rejects(()=>normalizeWordLayout(input,{fetchImpl:async()=>({ok:true,json:async()=>({is_standard:false})})}));
 await assert.rejects(()=>normalizeWordLayout({...input,docx_sha256:'0'.repeat(64)},{fetchImpl:async()=>{throw Error('Should not call')}}));
});
