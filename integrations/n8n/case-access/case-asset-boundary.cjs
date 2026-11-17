function verifyCaseAsset(m,c,request,parent){
if(m.id!==request.item_id||!m.file||m.folder||!m.eTag||!(m.size>0))throw Error('Requested asset is not a current case file');
const direct=m.parentReference?.id===c.folder_id;
const review=parent?.id===m.parentReference?.id&&!!parent.folder&&parent.parentReference?.id===c.folder_id&&/^Maddy retainer review - [a-f0-9]{64}$/i.test(parent.name||'')&&/^retainer-draft\.pdf$/i.test(m.name||'');
if(!direct&&!review)throw Error('Requested asset is outside this case or its verified retainer review folder');
if(m.size>30*1024*1024)throw Error('Asset requires a larger-document research read; it is not missing');
return {...m,verified_asset_parent_id:m.parentReference.id,retainer_review_child:!direct};
}

module.exports={verifyCaseAsset};
