import {createHash} from 'node:crypto';
const ENDPOINT='https://web-production-56acb.up.railway.app/normalize-saved-notice-word';
export async function normalizeWordLayout(result,{fetchImpl=fetch}={}) {
 const source=Buffer.from(result.docx_base64||'','base64');
 if(source.toString('base64')!==result.docx_base64 || createHash('sha256').update(source).digest('hex')!==result.docx_sha256)throw Error('Patched source binding failed');
 const response=await fetchImpl(ENDPOINT,{method:'POST',headers:{'Content-Type':'application/json'},
   body:JSON.stringify({docx_base64:result.docx_base64,expected_sha256:result.docx_sha256}),signal:AbortSignal.timeout(15000)});
 if(!response.ok)throw Error('Standard Word layout correction requires review');
 const body=await response.json();
 const bytes=Buffer.from(body.docx_base64||'','base64');
 if(body.is_standard!==true || body.source_docx_sha256!==result.docx_sha256 ||
    body.word_template_version!=='residential-pay-or-quit-2026-10-30-word-layout4' ||
    bytes.toString('base64')!==body.docx_base64 || bytes.subarray(0,2).toString()!=='PK' ||
    createHash('sha256').update(bytes).digest('hex')!==body.docx_sha256 ||
    !body.printed_signer?.name || !body.printed_signer?.title)throw Error('Normalized Word binding failed');
 return {...result,docx_base64:body.docx_base64,docx_sha256:body.docx_sha256,
   word_template_version:body.word_template_version,printed_signer:body.printed_signer};
}
