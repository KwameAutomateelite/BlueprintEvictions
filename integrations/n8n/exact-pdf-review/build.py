"""Build review changes from a freshly read workflow; performs no API writes."""
from copy import deepcopy
import json
from pathlib import Path
import uuid

P = Path(__file__).parent
BASE = "https://web-production-56acb.up.railway.app"


def node(name, kind, parameters, version=2):
    return {"id": str(uuid.uuid4()), "name": name, "type": "n8n-nodes-base." + kind,
            "typeVersion": version, "position": [0, 0], "parameters": parameters}


def code(name, filename=None, text=None):
    return node(name, "code", {"jsCode": text if text is not None else (P / filename).read_text()})


def http(name, url, credentials=None, **parameters):
    p = {"url": url, "options": {}, **parameters}
    n = node(name, "httpRequest", p, 4.2)
    if credentials:
        p.update(authentication="predefinedCredentialType", nodeCredentialType="microsoftOneDriveOAuth2Api")
        n["credentials"] = deepcopy(credentials)
    return n


def condition(name, expression):
    return node(name, "if", {"conditions": {"options": {"version": 2, "caseSensitive": True, "typeValidation": "strict"},
        "conditions": [{"id": str(uuid.uuid4()), "leftValue": expression, "rightValue": True,
                        "operator": {"type": "boolean", "operation": "true", "singleValue": True}}],
        "combinator": "and"}, "options": {}}, 2.2)


def edge(name):
    return {"node": name, "type": "main", "index": 0}


def connect(connections, source, *targets):
    connections[source] = {"main": [[edge(x) for x in targets]]}


def child(before):
    credentials = next(n for n in before["nodes"] if n["name"] == "Read Approved Word Metadata")["credentials"]
    file_options = {"response": {"response": {"responseFormat": "file"}}}
    nodes = [
        node("Review Snapshot Input", "executeWorkflowTrigger", {"inputSource": "passthrough"}, 1.1),
        code("Validate Review Context", "validate-context.js"),
        condition("Reviewer Policy Enabled", "={{ $json.needs_preview === true }}"),
        http("Read Saved Word Metadata", "={{ 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent($json.document_id) }}", credentials),
        code("Verify Saved Word Location", "verify-word.js"),
        http("Read Saved Word Bytes", "={{ 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent($json.id) + '/content' }}", credentials, options=file_options),
        code("Prepare Saved Word Snapshot", "prepare-source.js"),
        http("Render Saved Word as Exact PDF", BASE + "/render-saved-notice-pdf", method="POST", sendBody=True, specifyBody="json", jsonBody="={{ JSON.stringify($json) }}"),
        condition("Original Residential Form", "={{ $json.is_standard === true }}"),
        code("Prepare Exact PDF", "prepare-pdf.js"),
        http("Save Exact Review PDF", "={{ 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent($('Validate Review Context').first().json.folder_id) + ':/' + encodeURIComponent($json.name) + ':/content' }}", credentials,
             method="PUT", sendBody=True, contentType="binaryData", inputDataFieldName="data", sendHeaders=True,
             headerParameters={"parameters": [{"name": "Content-Type", "value": "application/pdf"}]}),
        code("Verify Saved PDF Location", "verify-pdf-saved.js"),
        http("Read Saved PDF Bytes", "={{ 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent($json.id) + '/content' }}", credentials, options=file_options),
        code("Verify Saved PDF Bytes", "verify-pdf-bytes.js"),
        http("Grant Reviewer PDF Read Access", "={{ 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent($json.id) + '/invite' }}", credentials,
             method="POST", sendBody=True, specifyBody="json",
             jsonBody="={{ JSON.stringify({recipients:[{email:$('Validate Review Context').first().json.reviewer}],requireSignIn:true,sendInvitation:false,roles:['read'],retainInheritedPermissions:true}) }}"),
        code("Verify PDF Reviewer Access", "verify-access.js"),
        http("Recheck Saved Word Metadata", "={{ 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent($('Validate Review Context').first().json.document_id) }}", credentials),
        code("Return PDF Review Snapshot", "finish-preview.js"),
        code("Return Other Review Snapshot", text="""const s=$('Validate Review Context').first().json.snapshot;
let r;try{r=$('Render Saved Word as Exact PDF').first().json;}catch{}
if(r?.is_standard===false && (s.context?.review_format==='pdf' ||
 (s.renderedPayload?.notice_type==='residential' && s.renderedPayload?.day_count===3)))
 throw Error('Residential Word form lost its fields. Regenerate a fresh PDF review; nothing approved.');
return [{json:s}];"""),
    ]
    connections = {}
    main_names = [n["name"] for n in nodes[:-1]]
    for a, b in zip(main_names, main_names[1:]):
        connect(connections, a, b)
    connections["Reviewer Policy Enabled"]["main"].append([edge("Return Other Review Snapshot")])
    connections["Original Residential Form"]["main"].append([edge("Return Other Review Snapshot")])
    for i, n in enumerate(nodes):
        n["position"] = [(i % 6) * 250, (i // 6) * 220]
    return {"name": "Maddy Exact PDF Review from Saved Word", "nodes": nodes, "connections": connections,
            "settings": {"executionOrder": "v1", "callerPolicy": "workflowsFromSameOwner"}}


def parent(before, child_id):
    result = deepcopy(before)
    credentials = next(n for n in before["nodes"] if n["name"] == "Read Approved Word Metadata")["credentials"]
    wrappers = ["Bundle For Review", "Update Revision Context", "Saved Edit: Advance Review Snapshot"]
    # Keep the public node names for downstream expressions. Their Word-only
    # builders become sources; the same named outputs now carry the PDF snapshot.
    for name in wrappers:
        original = next(n for n in result["nodes"] if n["name"] == name)
        source_name = name + " - Word Editing Source"
        original["name"] = source_name
        outgoing = deepcopy(result["connections"].pop(name))
        for c in result["connections"].values():
            for branches in c.values():
                for branch in branches:
                    for item in branch:
                        if item["node"] == name:
                            item["node"] = source_name
        wrapper = node(name, "executeWorkflow", {"source": "database", "workflowId": {"__rl": True, "value": child_id, "mode": "id"},
            "mode": "each", "options": {"waitForSubWorkflow": True}}, 1.3)
        wrapper["position"] = [original["position"][0] + 200, original["position"][1]]
        source_edges = deepcopy(outgoing)
        source_edges["main"][0] = [edge(name)]
        result["connections"][source_name] = source_edges
        result["connections"][name] = outgoing
        if original.get("onError"):
            wrapper["onError"] = original["onError"]
        result["nodes"].append(wrapper)
    additions = [
        http("Render Approved Saved Word as Exact PDF", BASE + "/render-saved-notice-pdf", method="POST", sendBody=True, specifyBody="json",
             jsonBody="={{ JSON.stringify({docx_base64:$json.docx_base64,expected_sha256:$json.docx_sha256}) }}"),
        code("Check Approved PDF Mode", text="""const r=$input.first().json,c=$('Build Approval Email').item.json.reviewSnapshot.context;
if(c.review_format==='pdf' && r.is_standard!==true)throw Error('Saved Word changed notice family. Fresh PDF review required.');
return [{json:r}];"""),
        condition("Approved Original Residential Form", "={{ $json.is_standard === true }}"),
        http("Read Reviewed Exact PDF Bytes", "={{ 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent($('Build Approval Email').item.json.reviewSnapshot.context.review_pdf_file_id || 'missing-review-copy') + '/content' }}", credentials,
             options={"response": {"response": {"responseFormat": "file"}}}),
        code("Use Approved Exact PDF", "approved-pdf.js"),
    ]
    result["nodes"].extend(additions)
    connect(result["connections"], "Capture Approved Word Snapshot", additions[0]["name"])
    connect(result["connections"], additions[0]["name"], additions[1]["name"])
    connect(result["connections"], additions[1]["name"], additions[2]["name"])
    result["connections"][additions[2]["name"]] = {"main": [[edge(additions[3]["name"] )], [edge("Convert Approved Saved Word")]]}
    connect(result["connections"], additions[3]["name"], additions[4]["name"])
    connect(result["connections"], additions[4]["name"], "Capture Approved PDF Snapshot")
    refresh_if = condition("PDF Refresh Requested", "={{ $json.revision_channel === 'email' && $json.context?.review_format === 'pdf' && String($json.reply_text || '').trim().replace(/[.!]+$/, '').toUpperCase() === 'REFRESH PDF' }}")
    refresh_source = code("Prepare Verified PDF Refresh", "refresh-request.js")
    refresh = node("Refresh PDF Review Snapshot", "executeWorkflow", {"source": "database",
        "workflowId": {"__rl": True, "value": child_id, "mode": "id"}, "mode": "each",
        "options": {"waitForSubWorkflow": True}}, 1.3)
    result["nodes"].extend([refresh_if, refresh_source, refresh])
    previous_revision_targets = deepcopy(result["connections"]["Revision Input"])
    connect(result["connections"], "Revision Input", refresh_if["name"])
    result["connections"][refresh_if["name"]] = {"main": [[edge(refresh_source["name"])], previous_revision_targets["main"][0]]}
    connect(result["connections"], refresh_source["name"], refresh["name"])
    connect(result["connections"], refresh["name"], "Build Approval Email")
    email = next(n for n in result["nodes"] if n["name"] == "Build Approval Email")
    js = email["parameters"]["jsCode"]
    old = "'<p>The ' + esc(noticeType || 'notice') + ' is drafted and saved. Open the link to review it in place - it is the live document, so any edit you make there is the version that goes out.</p>',"
    assert old in js
    js = js.replace(old, """ctx.review_format === 'pdf'
    ? '<p>Your notice is ready. Review and approve the PDF below. Keep the Word copy for edits; after changes I will return a fresh PDF for approval.</p>'
    : """ + old)
    old = "'<p><b>Review the notice:</b><br><a href=\"' + esc(link) + '\">' + esc(link) + '</a></p>',"
    assert old in js
    js = js.replace(old, """'<p><b>' + (ctx.review_format === 'pdf' ? 'Review and approve this PDF:' : 'Review the notice:') + '</b><br><a href="' + esc(link) + '">' + esc(link) + '</a></p>',
  ctx.review_format === 'pdf' ? '<p><a href="' + esc(ctx.word_edit_link) + '">Open Word for edits</a>. After saving edits, reply <b>REFRESH PDF</b> to receive an updated approval copy. Approve that updated PDF before it is sent.</p>' : '',""")
    js = js.replace("const totalBlock = !src.documentRevision && totalUsed > 0", "const totalBlock = (!src.documentRevision || ctx.review_format === 'pdf') && totalUsed > 0")
    email["parameters"]["jsCode"] = js
    return result


if __name__ == "__main__":
    import sys
    before = json.loads(Path(sys.argv[1]).read_text())
    output = Path(sys.argv[2]); output.mkdir(exist_ok=True)
    (output / "preview-child-candidate.json").write_text(json.dumps(child(before)))
    if len(sys.argv) == 4:
        (output / "workflow-proposed.json").write_text(json.dumps(parent(before, sys.argv[3])))
