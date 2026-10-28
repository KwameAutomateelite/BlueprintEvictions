"""The original PDF, not a Word recreation, is the layout authority."""
from io import BytesIO
from pathlib import Path
import sys

import pytest
from pypdf import PdfReader

sys.path.insert(0, str(Path(__file__).parents[1]))
sys.path.insert(0, str(Path(__file__).parent))
from residential_pdf import MASTER, PDF_TEMPLATE_VERSION, render_residential_pdf
from residential_notice import NoticeInputError
from test_residential_standard import case


def widgets(reader):
    for page_no, page in enumerate(reader.pages, 1):
        for ref in page["/Annots"]:
            widget = ref.get_object()
            field = widget.get("/Parent", widget).get_object()
            if field.get("/FT") == "/Tx":
                yield page_no, str(field["/T"]), widget, field


@pytest.mark.parametrize("rows", [1, 3, 6])
def test_original_pages_coordinates_and_canonical_values(rows):
    original = PdfReader(MASTER)
    result = PdfReader(BytesIO(render_residential_pdf(case(rows))))
    assert len(result.pages) == 2
    for source, actual in zip(original.pages, result.pages):
        assert source.get_contents().get_data() == actual.get_contents().get_data()
        assert list(source.mediabox) == list(actual.mediabox) == [0, 0, 612, 792]
    source_widgets = list(widgets(original))
    actual_widgets = list(widgets(result))
    assert len(actual_widgets) == len(source_widgets) == 27
    canonical = result.get_fields()
    for (sp, sn, sw, _), (ap, an, aw, af) in zip(source_widgets, actual_widgets):
        assert (sp, sn, list(sw["/Rect"])) == (ap, an, list(aw["/Rect"]))
        assert canonical[an]["/V"] == af["/V"]
        assert aw["/AP"]["/N"].get_data()
    assert canonical["TOTAL RENT DUE"]["/V"] == f"${rows * 1000:,.2f}"
    assert canonical["Text5"]["/V"] == "Sam Example, Property Manager"
    assert canonical["NAME OF PERSON SERVED"]["/V"] == "10/26/26"
    assert canonical["PLAINTIFF"]["/V"] != canonical["Text5"]["/V"]
    assert canonical["cc"]["/V"] == ""
    for i in range(rows + 1, 7):
        assert canonical[f"Month{i}"]["/V"] == canonical[f"AmountDue{i}"]["/V"] == ""
    assert result.metadata["/NoticeTemplateVersion"] == PDF_TEMPLATE_VERSION


def test_both_linked_contact_appearances_update_and_can_differ():
    data = case()
    data.update(housing_provider_street="987 Different Street", housing_provider_city="San Jose, CA 95125", housing_provider_phone="(408) 555-0199")
    result = PdfReader(BytesIO(render_residential_pdf(data)))
    fields = result.get_fields()
    assert fields["PLAINTIFF STREET ADDRESS"]["/V"] == "456 Example Blvd, Suite 200"
    assert fields["PLAINTIFF STREET ADDRESS (provider)"]["/V"] == "987 Different Street"
    assert fields["PLAINTIFF CITY/STATE/ZIP (provider)"]["/V"] == "San Jose, CA 95125"
    for page_no, name, widget, field in widgets(result):
        assert fields[name]["/V"] == field["/V"]
        if name.startswith("PLAINTIFF STREET ADDRESS"):
            expected = data["payment_address_street"] if page_no == 1 else data["housing_provider_street"]
            assert expected.encode().hex().encode() in widget["/AP"]["/N"].get_data()


@pytest.mark.parametrize("change", [
    {"total_amount_due": "1"}, {"signer_name": ""},
    {"property_address_city": "W" * 60},
])
def test_invalid_or_overflow_inputs_hold_instead_of_reflow(change):
    data = case()
    data.update(change)
    with pytest.raises(NoticeInputError):
        render_residential_pdf(data)


def test_no_replaced_reference_values_in_any_written_object():
    reader = PdfReader(BytesIO(render_residential_pdf(case(6))))
    from pypdf.generic import IndirectObject
    for generation, objects in reader.xref.items():
        if generation != 0:
            continue
        for object_id in objects:
            if object_id == 0:
                continue
            obj = reader.get_object(IndirectObject(object_id, generation, reader))
            raw = obj.get_data() if hasattr(obj, "get_data") else repr(obj).encode()
            for forbidden in (b"750 Calla", b"Leilei", b"Hernandez", b"Chao Pang"):
                assert forbidden not in raw


def test_pdf_api_and_existing_docx_contract(monkeypatch):
    monkeypatch.setenv("DROPBOX_SIGN_API_KEY", "local-test-no-network")
    monkeypatch.setenv("AIRTABLE_API_KEY", "local-test-no-network")
    from fastapi.testclient import TestClient
    import main
    from unittest.mock import patch
    # Other notice/signature routes call this existing synchronous helper.
    # The additive endpoint must not shadow or change its calling convention.
    with patch.object(main, "fill_template", return_value="/tmp/fictional.docx"), patch.object(main, "convert_docx_to_pdf", return_value="/tmp/fictional.pdf"), patch.object(main.os, "unlink"):
        assert main.generate_notice_pdf("commercial", {}) == "/tmp/fictional.pdf"
    client = TestClient(main.app)
    data = case()
    data["attachments_required"] = ["TPO_Required_Attachment.pdf"]
    response = client.post("/generate-notice-pdf", json=data)
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/pdf"
    assert response.headers["x-notice-template-version"] == PDF_TEMPLATE_VERSION
    assert response.headers["x-attachments-needed"] == "TPO_Required_Attachment.pdf"
    assert len(PdfReader(BytesIO(response.content)).pages) == 2
    assert client.post("/generate-notice", json=data).content.startswith(b"PK")
    data["total_amount_due"] = "1"
    assert client.post("/generate-notice-pdf", json=data).status_code == 422
    data["notice_type"] = "commercial"
    assert client.post("/generate-notice-pdf", json=data).status_code == 422
