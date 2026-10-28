"""Check that PDF wiring preserves existing policy, edits and delivery controls."""
from copy import deepcopy
import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location(
    "pdf_review_builder", ROOT / "integrations/n8n/exact-pdf-review/build.py"
)
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)


def before():
    names = ["Bundle For Review", "Update Revision Context", "Saved Edit: Advance Review Snapshot"]
    nodes = [dict(id=name, name=name, type="n8n-nodes-base.code", parameters={"jsCode": "return [];"},
                  position=[100, 100], onError="continueErrorOutput") for name in names]
    nodes += [dict(id="metadata", name="Read Approved Word Metadata", credentials={"example": {"id": "fake"}}),
              dict(id="email", name="Build Approval Email", parameters={"jsCode": "\n".join([
                  "'<p>The ' + esc(noticeType || 'notice') + ' is drafted and saved. Open the link to review it in place - it is the live document, so any edit you make there is the version that goes out.</p>',",
                  "'<p><b>Review the notice:</b><br><a href=\"' + esc(link) + '\">' + esc(link) + '</a></p>',",
                  "const totalBlock = !src.documentRevision && totalUsed > 0",
              ])})]
    connections = {name: {"main": [[builder.edge("Build Approval Email")], [builder.edge("Existing Hold")]]}
                   for name in names}
    connections["Initial Word Saved"] = {"main": [[builder.edge(names[0])]]}
    connections["Revision Input"] = {"main": [[builder.edge("Use Saved Document Editing")]]}
    connections["Freeze Approved Delivery"] = {"main": [[builder.edge("Existing Signing Gate")]]}
    return dict(nodes=nodes, connections=connections, settings={"executionOrder": "v1", "errorWorkflow": "fake"},
                staticData={"node:Status to AM Review": {"lastTimeChecked": "unchanged"}})


def test_wraps_word_sources_without_changing_original_code_or_error_routes():
    original = before()
    preserved = deepcopy(original)
    result = builder.parent(original, "CHILD-ID")
    assert original == preserved
    assert len(result["nodes"]) == len(original["nodes"]) + 11
    by_name = {n["name"]: n for n in result["nodes"]}
    for old in original["nodes"][:3]:
        renamed = deepcopy(old)
        renamed["name"] += " - Word Editing Source"
        assert by_name[renamed["name"]] == renamed
        wrapper = by_name[old["name"]]
        assert wrapper["parameters"]["workflowId"]["value"] == "CHILD-ID"
        assert wrapper["parameters"]["options"]["waitForSubWorkflow"] is True
        assert wrapper["onError"] == old["onError"]
        assert result["connections"][old["name"]] == original["connections"][old["name"]]
        assert result["connections"][renamed["name"]]["main"][1] == [builder.edge("Existing Hold")]
    assert result["connections"]["Initial Word Saved"]["main"][0][0]["node"].endswith("Word Editing Source")
    assert result["settings"] == original["settings"]
    assert result["staticData"] == original["staticData"]
    assert result["connections"]["Freeze Approved Delivery"] == original["connections"]["Freeze Approved Delivery"]


def test_saved_edits_refresh_pdf_and_approval_cannot_bypass_pdf_checks():
    original = before()
    result = builder.parent(original, "CHILD-ID")
    nodes = {n["name"]: n for n in result["nodes"]}
    connections = result["connections"]
    assert connections["Revision Input"]["main"][0] == [builder.edge("PDF Refresh Requested")]
    assert connections["PDF Refresh Requested"]["main"][1] == original["connections"]["Revision Input"]["main"][0]
    assert connections["Refresh PDF Review Snapshot"]["main"][0] == [builder.edge("Build Approval Email")]
    assert connections["Capture Approved Word Snapshot"]["main"][0] == [builder.edge("Render Approved Saved Word as Exact PDF")]
    assert connections["Approved Original Residential Form"]["main"] == [
        [builder.edge("Read Reviewed Exact PDF Bytes")], [builder.edge("Convert Approved Saved Word")]]
    assert connections["Use Approved Exact PDF"]["main"][0] == [builder.edge("Capture Approved PDF Snapshot")]
    assert "Open Word for edits" in nodes["Build Approval Email"]["parameters"]["jsCode"]
    assert "REFRESH PDF" in nodes["Build Approval Email"]["parameters"]["jsCode"]
    # Current loop item must be selected, rather than the first review iteration.
    assert "$('Build Approval Email').item" in nodes["Read Reviewed Exact PDF Bytes"]["parameters"]["url"]
    assert ".first()" not in nodes["Use Approved Exact PDF"]["parameters"]["jsCode"]
