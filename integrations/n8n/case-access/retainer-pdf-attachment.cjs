function prepareRetainerPdf(response,maxBytes=20*1024*1024) {
  if(response['@odata.nextLink'])throw Error('Incomplete attachment list');
  const attachments=(response.value||[]).filter(x=>!x.isInline);
  if(attachments.length!==1||attachments[0]['@odata.type']!=='#microsoft.graph.fileAttachment'||!(/\.pdf$/i.test(attachments[0].name||'')))throw Error('Need one signed agreement PDF; review required');
  const a=attachments[0],bytes=Buffer.from(a.contentBytes||'','base64');
  if(bytes.length<5||bytes.length>maxBytes||bytes.subarray(0,5).toString()!=='%PDF-')throw Error('Invalid PDF or larger than the supported 20 MiB signed-retainer limit');
  const sha=require('crypto').createHash('sha256').update(bytes).digest('hex');
  return {attachment_id:a.id,pdf_base64:a.contentBytes,original_name:a.name,sha256:sha,size:bytes.length,stored_name:'Signed-Retainer-'+sha+'.pdf'};
}
if(typeof module!=='undefined')module.exports={prepareRetainerPdf};
