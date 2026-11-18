function selectThreadPdf(command, messages, attachmentSets) {
  if (!command.id || !command.conversationId || !Number.isFinite(Date.parse(command.receivedDateTime))) throw new Error('Message context missing');
  const eligible=new Map(messages.filter(m=>m.conversationId===command.conversationId && Number.isFinite(Date.parse(m.receivedDateTime)) && Date.parse(m.receivedDateTime)<=Date.parse(command.receivedDateTime)).map(m=>[m.id,m]));
  const unique=new Map();
  const current=attachmentSets.filter(set=>set.messageId===command.id&&(set.value||[]).some(a=>a['@odata.type']==='#microsoft.graph.fileAttachment'&&!a.isInline&&/\.pdf$/i.test(a.name||'')));
  // A fresh proof attachment belongs to this reply; older signed packets must not compete with it.
  for(const set of (current.length?current:attachmentSets)){
    if(!eligible.has(set.messageId))continue;
    for(const a of set.value||[]){
      if(a['@odata.type']!=='#microsoft.graph.fileAttachment' || a.isInline || !/\.pdf$/i.test(a.name||''))continue;
      if(!a.contentBytes)throw new Error('PDF bytes not retrieved');
      // Copies of the same PDF in a conversation count once; names alone are insufficient.
      if(!unique.has(a.contentBytes))unique.set(a.contentBytes,{...a,sourceMessageId:set.messageId});
    }
  }
  if(unique.size!==1)return {ok:false,reason:unique.size?'multiple_pdfs_require_selection':'no_pdf_in_thread',candidates:[...unique.values()].map(a=>({name:a.name,sourceMessageId:a.sourceMessageId,attachmentId:a.id}))};
  return {ok:true,pdf:[...unique.values()][0]};
}


module.exports={selectThreadPdf};
