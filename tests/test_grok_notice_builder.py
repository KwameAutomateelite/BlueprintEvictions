import importlib.util
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
