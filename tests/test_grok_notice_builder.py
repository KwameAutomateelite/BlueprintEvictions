import importlib.util
import json
import subprocess
from pathlib import Path

P=Path(__file__).resolve().parents[1]/'integrations/n8n/grok-notice'
spec=importlib.util.spec_from_file_location('grok_builder',P/'build.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)

def fake_read():
    names=m.PREFIX+['Read Current Document','Require Bound Training Document']
    return {'nodes':[{'id':n,'name':n,'type':'n8n-nodes-base.code','parameters':{'jsCode':
      "const r=$json,input=$('Validate Case Tool').first().json;return [{json:r}];"},'credentials':{'ref':'unchanged'}} for n in names]}

def test_draft_writes_only_new_files_and_never_sends_or_changes_case():
    w=m.draft(fake_read(),'POLICY','PDF');nodes={n['name']:n for n in w['nodes']}
    assert not any('openai' in n['type'].lower() or 'slack' in n['type'].lower() for n in w['nodes'])
    reserve=nodes['Reserve New Draft Upload']['parameters']
    assert "conflictBehavior':'fail'" in reserve['jsonBody']
    assert '/createUploadSession' in reserve['url']
    upload=nodes['Save New Draft Without Overwrite']
    assert 'credentials' not in upload
    assert all('api.airtable.com' not in n['parameters'].get('url','') or n['parameters'].get('method','GET')=='GET' for n in w['nodes'])
    assert nodes['Prepare Exact PDF Review']['onError']=='continueErrorOutput'
    assert w['connections']['Prepare Exact PDF Review']['main'][1][0]['node']=='Return Saved Word PDF Hold'

def test_refresh_requires_exact_version_and_leaves_inspection_separate():
    before=fake_read();w=m.review(before,'POLICY','PDF')
    guard=next(n for n in w['nodes'] if n['name']=='Require Bound Training Document')['parameters']['jsCode']
    assert 'r.eTag!==input.expected_version' in guard
    assert before==fake_read()
    assert all(n['parameters'].get('method','GET')=='GET' for n in w['nodes'] if n['type'].endswith('httpRequest'))

def test_edit_routes_only_residential_three_day_to_pdf_and_preserves_existing_nodes():
    old={'nodes':[{'id':'same','name':'Return Indexed Edit Receipt','parameters':{'jsCode':'original'}}],
         'connections':{},'settings':{'executionOrder':'v1'}}
    w=m.revise(old,'REVIEW')
    assert w['nodes'][:1]==old['nodes'] and w['settings']==old['settings']
    expr=w['nodes'][1]['parameters']['conditions']['conditions'][0]['leftValue'][3:-2]
    code="""const expr=EXPR;let type;const $=()=>({first:()=>({json:{fields:{'Notice Type':type}}})});
const vals=['3-day to pay or quit','three day notice to pay rent or quit','Commercial 3-day pay rent or quit','5-day pay rent or quit','30-day termination'];
process.stdout.write(JSON.stringify(vals.map(t=>{type=t;return eval(expr)})));""".replace('EXPR',json.dumps(expr))
    result=subprocess.run(['node','-e',code],text=True,capture_output=True,check=True)
    assert json.loads(result.stdout)==[True,True,False,False,False]
    assert w['connections']['Prepare Edited Exact PDF']['main'][1][0]['node']=='Return Edited PDF Hold'

def test_revision_contract_documents_guarded_shapes_without_changing_dispatch():
    before={'nodes':[{'name':'revise_case_notice','credentials':{'id':'keep'},
        'parameters':{'description':'Existing guarded editor',
            'workflowId':{'value':'EXISTING_EDITOR'},
            'workflowInputs':{'value':{'record_id':'keep', 'plan_json':'old'}}}}],
        'connections':{'unchanged':{}},'settings':{'unchanged':True}}
    after=m.revision_contract(before)
    node=after['nodes'][0]
    expression=node['parameters']['workflowInputs']['value']['plan_json']
    assert 'rent_changes' in expression and 'row_index' in expression
    assert 'ZERO-BASED' in expression and 'not a paragraph index' in expression
    assert 'recalculates' in expression and 'deterministic phone parser' in expression
    assert 'separate calls' in expression and 'instruction_quote' in expression
    assert node['parameters']['workflowId']==before['nodes'][0]['parameters']['workflowId']
    assert node['credentials']==before['nodes'][0]['credentials']
    assert after['connections']==before['connections'] and after['settings']==before['settings']
    assert before['nodes'][0]['parameters']['workflowInputs']['value']['plan_json']=='old'
    assert m.revision_contract(after)==after
