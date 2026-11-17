"""Case details that fit the actual master must survive Word/PDF generation."""
from io import BytesIO
import re

from docx import Document
from pypdf import PdfReader
import pytest

from residential_notice import NoticeInputError, fields_for_case, render_residential_notice
from residential_pdf import render_residential_pdf
from residential_word_snapshot import saved_word_case
from .test_residential_standard import case


TRUST = "David P. Smith and Mona-Lisa M. Smith Inter Vivos Trust dated November 6, 2000"
STREET = "c/o Shipping Xpress, Attn: Blueprint Evictions, 1060 Lincoln Ave, Suite 20-1022"


def source_case(rows=2):
    data = case(rows)
    data.update(landlord_name=TRUST, payment_recipient_name=TRUST,
                payment_address_street=STREET,
                payment_address_city="San Jose, CA 95125")
    return data


def widgets(pdf):
    for page_number, page in enumerate(PdfReader(BytesIO(pdf)).pages):
        for ref in page.get("/Annots", []):
            widget = ref.get_object()
            parent = widget.get("/Parent", widget).get_object()
            if widget.get("/Subtype") == "/Widget" and parent.get("/V"):
                yield page_number, str(parent["/T"]), str(parent["/V"]), widget


def test_complete_trust_and_care_of_address_survive_both_formats():
    data = source_case()
    word = render_residential_notice(data)
    extracted = saved_word_case(word)
    assert extracted["payment_recipient_name"] == TRUST
    assert extracted["payment_address_street"] == STREET
    assert extracted["housing_provider_street"] == STREET
    assert extracted["housing_provider_city"] == "San Jose, CA 95125"
    doc = Document(BytesIO(word))
    street = next(p for p in doc.paragraphs if any(
        b.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name") == "BE_PROVIDER_STREET"
        for b in p._p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart")))
    assert street.text.count("\n") == 1
    assert all(r.font.name == "Times New Roman" and r.font.size.pt == 12 for r in street.runs if r.text)
    pdf = render_residential_pdf(data)
    assert len(PdfReader(BytesIO(pdf)).pages) == 2
    values = [value for _, _, value, _ in widgets(pdf)]
    assert TRUST in values and STREET in values
    assert any(value.replace("\n", " ") == STREET for value in values)
    for _, _, _, widget in widgets(pdf):
        assert re.search(r"/(?:TiRo|TiBo) 12 Tf", str(widget["/DA"]))


def test_wrapped_provider_moves_city_and_phone_together_without_moving_signature():
    base = list(widgets(render_residential_pdf(case(2))))
    wrapped = list(widgets(render_residential_pdf(source_case())))
    def get(items, name):
        return next(w for page, n, _, w in items if page == 1 and n.split(" (provider)")[0] == name)
    for name in ("PLAINTIFF CITY/STATE/ZIP", "Text6"):
        a, b = get(base, name)["/Rect"], get(wrapped, name)["/Rect"]
        assert float(b[1]) == pytest.approx(float(a[1])-14)
        assert float(b[3]) == pytest.approx(float(a[3])-14)
    assert get(base, "Text5")["/Rect"] == get(wrapped, "Text5")["/Rect"]
    w = get(wrapped, "PLAINTIFF STREET ADDRESS")
    city = get(wrapped, "PLAINTIFF CITY/STATE/ZIP")
    assert float(w["/Rect"][1]) > float(city["/Rect"][3])-2


def test_provider_wrap_also_survives_rent_continuation_and_saved_word_refresh():
    data = source_case(12)
    word = render_residential_notice(data)
    extracted = saved_word_case(word)
    assert len(extracted["amounts_due"]) == 12
    assert extracted["housing_provider_street"] == STREET
    pdf = render_residential_pdf(extracted)
    assert any(value.replace("\n", " ") == STREET for _, _, value, _ in widgets(pdf))
    assert len(PdfReader(BytesIO(pdf)).pages) >= 2


@pytest.mark.parametrize("street", ["Very long mailing address "*20, "W"*100])
def test_more_than_two_provider_lines_still_requires_review(street):
    data = source_case(); data["housing_provider_street"] = street
    with pytest.raises(NoticeInputError, match="Housing-provider street"):
        fields_for_case(data)
