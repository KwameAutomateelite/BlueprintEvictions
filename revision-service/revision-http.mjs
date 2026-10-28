import {createServer} from 'node:http';import {timingSafeEqual,createHash} from 'node:crypto';

export function revisionServer({secret,worker}) {
 if(typeof secret!=='string'||secret.length<32||typeof worker!=='function')throw Error('Authenticated revision configuration required');
 const expected=Buffer.from(secret);
 return createServer(async(req,res)=>{
  const reply=(status,body)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(body))};
  if(req.method==='GET'&&req.url==='/health')return reply(200,{status:'ok',revision_enabled:true,mode:'persistent-document-transform',revision_version:'fixed-form-2026-10-26',rent_row_limit:1000});
  if(req.method!=='POST'||req.url!=='/events/notice-revision')return reply(404,{error:'not_found'});
  const given=Buffer.from(String(req.headers['x-maddy-event-key']||''));
  if(given.length!==expected.length||!timingSafeEqual(given,expected))return reply(401,{error:'unauthorized'});
  try {
   let size=0;const chunks=[];for await(const b of req){size+=b.length;if(size>15*1024*1024)return reply(413,{error:'request_too_large'});chunks.push(b)}
   const input=JSON.parse(Buffer.concat(chunks).toString());
   if(!/^rec[A-Za-z0-9]{14}$/.test(input.record_id||'')||!['inspect','patch'].includes(input.mode)||typeof input.document_id!=='string'||!input.document_id||typeof input.source_etag!=='string'||!input.source_etag||typeof input.docx_base64!=='string'||!/^[a-f0-9]{64}$/.test(input.expected_sha256))return reply(400,{error:'bound_document_required'});
   const bytes=Buffer.from(input.docx_base64,'base64');
   if(bytes.toString('base64')!==input.docx_base64||createHash('sha256').update(bytes).digest('hex')!==input.expected_sha256)return reply(409,{error:'source_integrity_mismatch'});
   const result=await worker({mode:input.mode,bytes,expected_sha256:input.expected_sha256,changes:input.changes});
   if(result.source_sha256!==input.expected_sha256)throw Error('Source binding failed');
   const output={...result,record_id:input.record_id,document_id:input.document_id,source_etag:input.source_etag,test_mode:input.test_mode===true};
   if(output.bytes){output.docx_base64=output.bytes.toString('base64');delete output.bytes;output.state='revised_requires_review'}
   return reply(200,output);
  }catch{return reply(409,{state:'held',reason:'document_revision_requires_review'})}
 });
}
