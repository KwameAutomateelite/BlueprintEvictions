"""Offline builders. Inputs are fresh workflow exports; no API calls or secrets."""
from copy import deepcopy
from pathlib import Path
import importlib.util
import json
import uuid

P=Path(__file__).parent
spec=importlib.util.spec_from_file_location('pdf_review_builder',P.parent/'exact-pdf-review/build.py')
b=importlib.util.module_from_spec(spec);spec.loader.exec_module(b)
BASE=b.BASE

REVISION_PLAN_DESCRIPTION = (
    'JSON STRING using one supported plan shape. Rent-only: '
    '{"state":"rent_changes","rows":[{"row_index":0,"new_amount":"1100.50",'
    '"instruction_quote":"verbatim reviewer clause naming the month and amount"}]}. '
    'row_index is ZERO-BASED in the latest read_case_notice rent_tables[0].rows, '
    'not a paragraph index. Do not propose a total edit: the tool recalculates it. '
    'Phone-only: use instruction "Change the landlord phone number from OLD to NEW. '
    'Keep everything else unchanged." and plan_json '
    '{"state":"held","question":"Use the deterministic phone parser"}; '
    'the parser updates both bound phone fields. Other text: '
    '{"state":"changes","changes":[{"old":"exact unique saved text",'
    '"new":"replacement","instruction_quote":"verbatim reviewer clause"}]}. '
    'Never use generic text changes for money. For mixed rent and phone requests, '
    'make separate calls with only the authorized clause for that edit and reread '
    'the saved Word/version between calls. Never use operations/edits/replacements '
    'as top-level plan keys. A rejected edit has not saved; clarify the contract '
    'rather than guessing another shape.'
)

def revision_contract(before):
    """Document the existing deployed edit protocol without changing its guards."""
    w=deepcopy(before)
    n=next(n for n in w['nodes'] if n['name']=='revise_case_notice')
    n['parameters']['workflowInputs']['value']['plan_json']=(
        "={{ $fromAI('plan_json', '"+REVISION_PLAN_DESCRIPTION+"', 'string') }}")
    suffix=(' EDIT CONTRACT: Read the current Word first. Split mixed rent and phone '
            'requests into separate calls and reread after each save. Follow the exact '
            'plan_json schema; rent row indexes are zero-based and totals are automatic. '
            'Show the last successful exact PDF after all requested corrections.')
    if suffix not in n['parameters']['description']:
        n['parameters']['description']+=suffix
    return w

def js(file): return (P/file).read_text()
def call(name,wid):
    return b.node(name,'executeWorkflow',{'source':'database','workflowId':{'__rl':True,'value':wid,'mode':'id'},
      'mode':'each','options':{'waitForSubWorkflow':True}},1.3)
def chain(nodes):
    return {a['name']:{'main':[[b.edge(z['name'])]]} for a,z in zip(nodes,nodes[1:])}
def copied(source,names):
    out=[]
    for name in names:
        n=deepcopy(next(n for n in source['nodes'] if n['name']==name));n['id']=str(uuid.uuid4());out.append(n)
    return out
def shape(name,nodes,connections=None):
    for i,n in enumerate(nodes):n['position']=[i%6*260,i//6*240]
    return {'name':name,'nodes':nodes,'connections':connections or chain(nodes),
      'settings':{'executionOrder':'v1','callerPolicy':'workflowsFromSameOwner'}}
PREFIX=['Called By Grok Tool','Validate Case Tool','Read Bound Case','Require Case Folder ID',
        'Read Exact Bound Folder','Return Verified Case']

def review(read,policy_id,pdf_id):
    nodes=copied(read,PREFIX+['Read Current Document','Require Bound Training Document'])
    nodes[1]['parameters']['jsCode']="""const b=$json;
if(!/^rec[A-Za-z0-9]{14}$/.test(b.record_id||'') || !/^[A-Za-z0-9_-]{10,80}$/.test(b.item_id||'') ||
 typeof b.expected_version!=='string'||!b.expected_version)throw Error('Exact case, Word ID and latest version are required.');
return [{json:b}];"""
    guard=nodes[-1]['parameters']['jsCode']
    # Inspection remains read-only elsewhere; this action requires a version.
    if 'expected_version' not in guard:
        guard=guard.replace('return [{json:r}];',"if(r.eTag!==input.expected_version)throw Error('Word changed. Read the latest version before preparing its PDF.');return [{json:r}];")
    nodes[-1]['parameters']['jsCode']=guard
    nodes += [b.code('Policy Case Input',text="return [{json:{record_id:$('Read Bound Case').first().json.id}}];"),
      call('Load Current Reviewer Policy',policy_id),b.code('Prepare Review Snapshot',text=js('prepare-review.js')),
      call('Prepare Exact PDF Review',pdf_id),b.code('Return Ready Review Receipt',text=js('review-receipt.js'))]
    return shape('Blueprint Case Notice Exact PDF Review',nodes)

def draft(read,policy_id,pdf_id):
    nodes=copied(read,PREFIX)
    nodes[1]['parameters']['jsCode']="""const b=$json;
if(!/^rec[A-Za-z0-9]{14}$/.test(b.record_id||''))throw Error('Exact notice case ID required.');
if(!/^[a-z0-9-]{8,48}$/.test(b.request_id||''))throw Error('Stable request_id of 8 to 48 lowercase letters, digits or hyphens required. Reuse it on retries.');
return [{json:{record_id:b.record_id,request_id:b.request_id}}];"""
    drive=next(n for n in read['nodes'] if n['name']=='Read Current Document')['credentials']
    # The production parser contains no n8n references after pure-function adaptation.
    parser=js('case-inputs.cjs').replace('module.exports={caseContext,resolveRentYear,noticePayload};','')
    fileopt={'response':{'response':{'responseFormat':'file'}}}
    nodes += [b.code('Policy Case Input',text="return [{json:{record_id:$('Read Bound Case').first().json.id}}];"),
      call('Load Current Reviewer Policy',policy_id),b.code('Prepare Verified Draft',text=parser+'\n'+js('prepare-draft.js')),
      b.http('Generate Standard Word',BASE+'/generate-notice',method='POST',sendBody=True,specifyBody='json',
        jsonBody='={{ JSON.stringify($json.payload) }}',options=fileopt),
      b.code('Capture Generated Word',text="""const raw=await this.helpers.getBinaryDataBuffer(0,'data');
if(raw.length<1000||raw.length>4000000)throw Error('Unexpected generated Word size');
return [{json:{...$('Prepare Verified Draft').first().json,word_base64:raw.toString('base64'),
word_sha256:require('crypto').createHash('sha256').update(raw).digest('hex'),size:raw.length}}];"""),
      b.http('Reserve New Draft Upload',"={{ 'https://graph.microsoft.com/v1.0/me/drive/items/'+$json.folder_id+':/'+encodeURIComponent($json.name)+':/createUploadSession' }}",drive,
        method='POST',sendBody=True,specifyBody='json',jsonBody="={{ JSON.stringify({item: {'@microsoft.graph.conflictBehavior':'fail',name:$json.name} }) }}"),
      b.code('Prepare New Draft Upload',text=r"""const s=$('Capture Generated Word').first().json;
if(!/^https:\/\//.test($json.uploadUrl||''))throw Error('Upload session unavailable; no overwrite attempted.');
return [{json:{upload_url:$json.uploadUrl,size:s.size},binary:{data:await this.helpers.prepareBinaryData(Buffer.from(s.word_base64,'base64'),s.name,'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}}];"""),
      b.http('Save New Draft Without Overwrite','={{ $json.upload_url }}',method='PUT',sendBody=True,contentType='binaryData',inputDataFieldName='data',sendHeaders=True,
        headerParameters={'parameters':[{'name':'Content-Range','value':"={{ 'bytes 0-'+($json.size-1)+'/'+$json.size }}"},{'name':'Content-Type','value':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}]}),
      b.code('Require New Draft Location',text="""const expected=$('Capture Generated Word').first().json,r=$json;
if(r.parentReference?.id!==expected.folder_id||r.name!==expected.name||!r.file||!r.eTag||r.size!==expected.size)throw Error('New Word save not verified. Do not claim ready.');return [{json:r}];"""),
      b.http('Read Saved New Draft Bytes',"={{ 'https://graph.microsoft.com/v1.0/me/drive/items/'+$json.id+'/content' }}",drive,options=fileopt),
      b.code('Verify New Draft Bytes',text="""const raw=await this.helpers.getBinaryDataBuffer(0,'data');
if(require('crypto').createHash('sha256').update(raw).digest('hex')!==$('Capture Generated Word').first().json.word_sha256)throw Error('Saved Word bytes differ; do not claim ready.');return [{json:$('Require New Draft Location').first().json}];"""),
      b.http('Grant Reviewer Word Editing',"={{ 'https://graph.microsoft.com/v1.0/me/drive/items/'+$json.id+'/invite' }}",drive,method='POST',sendBody=True,specifyBody='json',
        jsonBody="={{ JSON.stringify({recipients:[{email:$('Prepare Verified Draft').first().json.policy.reviewer_email}],requireSignIn:true,sendInvitation:false,roles:['write'],retainInheritedPermissions:true}) }}"),
      b.code('Verify Reviewer Editing',text="""const reviewer=$('Prepare Verified Draft').first().json.policy.reviewer_email.toLowerCase();
if(!$json.value?.some(p=>p.roles?.includes('write')&&(p.grantedToV2?.user?.email||p.grantedTo?.user?.email||'').toLowerCase()===reviewer))throw Error('Word saved but reviewer edit access unverified.');return [{json:$('Require New Draft Location').first().json}];"""),
      b.http('Recheck New Draft Metadata',"={{ 'https://graph.microsoft.com/v1.0/me/drive/items/'+$json.id }}",drive),
      b.code('Require Bound Training Document',text="""const old=$('Require New Draft Location').first().json;
if($json.id!==old.id||$json.eTag!==old.eTag||$json.parentReference?.id!==old.parentReference.id)throw Error('New Word changed during preparation. Fresh review required.');return [{json:$json}];"""),
      b.code('Prepare Review Snapshot',text=js('prepare-review.js')),
      call('Prepare Exact PDF Review',pdf_id),b.code('Return Ready Review Receipt',text=js('review-receipt.js'))]
    con=chain(nodes)
    nodes[-2]['onError']='continueErrorOutput'
    con['Prepare Exact PDF Review']['main'].append([b.edge('Return Saved Word PDF Hold')])
    nodes.append(b.code('Return Saved Word PDF Hold',text="""const w=$('Require Bound Training Document').first().json;
return [{json:{state:'word_saved_pdf_not_ready',case_id:$('Read Bound Case').first().json.id,item_id:w.id,version:w.eTag,word_edit_url:w.webUrl,
saved:true,approved:false,sent:false,pdf_ready:false,recovery:'Call prepare_case_notice_review using this exact case, item and version. Do not repeat drafting or claim PDF ready.'}}];"""))
    return shape('Blueprint Case Residential Three Day Draft',nodes,con)

def revise(before,review_id):
    w=deepcopy(before)
    old='Return Indexed Edit Receipt'
    add=[b.condition('Edited Residential Three Day',"={{ /(?:\\b3[ -]?day|\\bthree\\b)/i.test(String($('Read Bound Case').first().json.fields?.['Notice Type']||'')) && !/commercial|perform|warning|termination/i.test(String($('Read Bound Case').first().json.fields?.['Notice Type']||'')) }}"),
      b.code('Prepare Edited PDF Request',text="const r=$('Return Indexed Edit Receipt').first().json;return [{json:{record_id:r.case_id,item_id:r.item_id,expected_version:r.version}}];"),
      call('Prepare Edited Exact PDF',review_id),
      b.code('Return Edited PDF Receipt',text="return [{json:{...$('Return Indexed Edit Receipt').first().json,...$json,word_saved:true,requires_fresh_approval:true}}];"),
      b.code('Return Existing Other Notice Receipt',text="return [{json:$('Return Indexed Edit Receipt').first().json}];"),
      b.code('Return Edited PDF Hold',text="return [{json:{...$('Return Indexed Edit Receipt').first().json,state:'word_saved_pdf_not_ready',pdf_ready:false,recovery:'The edit was saved. Call prepare_case_notice_review with the returned case_id, item_id and version. Do not repeat the edit or claim the PDF is ready.'}}];")]
    add[2]['onError']='continueErrorOutput'
    w['nodes']+=add
    b.connect(w['connections'],old,add[0]['name'])
    w['connections'][add[0]['name']]={'main':[[b.edge(add[1]['name'])],[b.edge(add[4]['name'])]]}
    b.connect(w['connections'],add[1]['name'],add[2]['name'])
    w['connections'][add[2]['name']]={'main':[[b.edge(add[3]['name'])],[b.edge(add[5]['name'])]]}
    return w

def mcp(before,draft_id,review_id):
    w=deepcopy(before)
    template=next(n for n in w['nodes'] if n['name']=='read_case_notice')
    for name,wid,args,description in [
      ('draft_case_notice',draft_id,{'record_id':'Exact current BE Notices Airtable record ID; use get_case_documents first',
       'request_id':'Stable 8-48 lowercase letters/digits/hyphens; reuse on retries, never choose a new ID to bypass a duplicate'},
       'Create a NEW residential three-day pay-rent-or-quit notice from verified current intake case facts using the standard exact PDF generator, including long rent lists. Saves a new Word edit copy and matching PDF in its bound case folder. Does not send messages, approve or sign. Missing signer/period/year/amount facts produce a specific hold; never guess. Show PDF first and Word for edits only when state is draft_ready_for_review. A Slack/email notice-created message alone is not proof. On word_saved_pdf_not_ready call prepare_case_notice_review; do not draft again. Older notice migration is out of scope.'),
      ('prepare_case_notice_review',review_id,{'record_id':'Exact BE Notices Airtable record ID','item_id':'Exact current notice Word ID in its verified case folder',
       'expected_version':'Exact current eTag returned by read_case_notice'},
       'Explicitly prepare/save the exact standardized PDF from the CURRENT saved Word, including long itemized rent lists. Use after manual Word edits or to recover a saved draft whose PDF failed. Read current Word first and pass its exact version. Returns verified PDF for Anne-Michelle review, plus Word for edits. Say draft ready only for draft_ready_for_review. No message, approval, signing, status change or client delivery occurs.')]:
        n=deepcopy(template);n['id']=str(uuid.uuid4());n['name']=name
        n['parameters'].update(name=name,description=description,workflowId={'__rl':True,'mode':'id','value':wid})
        n['parameters']['workflowInputs']={'mappingMode':'defineBelow','value':{k:"={{ $fromAI('"+k+"', '"+v+"', 'string') }}" for k,v in args.items()},
          'schema':[{'id':k,'displayName':k,'required':True,'type':'string','display':True} for k in args],
          'matchingColumns':[],'attemptToConvertTypes':False,'convertFieldsToString':False}
        w['nodes'].append(n);w['connections'][name]=deepcopy(w['connections']['read_case_notice'])
    for n in w['nodes']:
        if n['name']=='read_case_notice':n['parameters']['description']+=' For the correct PDF review copy after edits, call prepare_case_notice_review with this exact item/version; do not present Word as the final PDF.'
        if n['name']=='revise_case_notice':n['parameters']['description']+=' Residential three-day edits automatically prepare the matching exact PDF. Present its PDF and Word links only when draft_ready_for_review. If word_saved_pdf_not_ready, retry only prepare_case_notice_review, never the edit. This cannot approve or send.'
        if n['name']=='get_case_documents':n['parameters']['description']+=' For a NEW residential three-day notice use draft_case_notice. Never claim notice creation from an email or Slack status alone; obtain a verified ready receipt with PDF and Word links.'
    return revision_contract(w)
