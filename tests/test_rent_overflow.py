"""Long schedules must retain every row, column, paragraph and approval byte."""
import base64
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import re
import sys
import zipfile

import pytest
from lxml import etree
from pypdf import PdfReader

sys.path[:0] = [str(Path(__file__).parents[1]), str(Path(__file__).parents[1]/"revision-service"), str(Path(__file__).parent)]
from notice_display import rent_period
from residential_notice import fields_for_case, render_residential_notice, NoticeInputError
from residential_pdf import MASTER, render_residential_pdf, _render_original
from residential_pdf_flow import plan_pages, BODY_BOTTOM, ROW_LEADING, ROW_PADDING
from residential_word_snapshot import render_saved_word_pdf
from standard_form import inspect_standard
from docx_patch import patch_docx
from test_residential_standard import case


def long_case(count):
    data = case(count)
    data["amounts_due"] = [{"due_date":f"{i%12+1}/1/{2020+i//12}","amount":"1,000"} for i in range(count)]
    return data


def static_text(reader):
    text = "".join(page.extract_text() for page in reader.pages)
    # Column headings repeat when the rent table continues onto a new page.
    for heading in ["TIME PERIOD DUE", "AMOUNT DUE"]:
        text = text.replace(heading, "")
    return re.sub(r"\s+", "", text)


@pytest.mark.parametrize("count", [7, 12, 24, 60, 120])
def test_all_entries_stay_in_two_columns_with_one_reconciled_total(count):
    data = long_case(count)
    reader = PdfReader(BytesIO(render_residential_pdf(data)))
    fields = reader.get_fields()
    assert len([k for k in fields if re.fullmatch(r"AmountDue\d+",k)]) == count
    for number,row in enumerate(data["amounts_due"],1):
        assert fields[f"Month{number}"]["/V"] == rent_period(row["due_date"])
        assert fields[f"AmountDue{number}"]["/V"] == "$1,000.00"
    assert fields["TOTAL RENT DUE"]["/V"] == f"${count*1000:,.2f}"
    assert static_text(reader) == static_text(PdfReader(MASTER))
    locations = {}
    for page_number,page in enumerate(reader.pages):
        for ref in page["/Annots"]:
            widget = ref.get_object()
            field = widget.get("/Parent",widget).get_object()
            name = str(field["/T"])
            locations[name] = (page_number,list(map(float,widget["/Rect"])))
            assert fields[name].get("/V", "") == field.get("/V", "")
            assert widget["/AP"]["/N"].get_data()
    for number in range(1,count+1):
        period,amount = locations[f"Month{number}"], locations[f"AmountDue{number}"]
        assert period[0] == amount[0]
        assert period[1][1::2] == amount[1][1::2]
        assert period[1][2] < amount[1][0]
    assert locations[f"Month{count}"][0] == locations["TOTAL RENT DUE"][0]
    assert locations["TOTAL RENT DUE"][1][3] < locations[f"Month{count}"][1][1]
    assert locations["Text5"][0] == locations["NAME OF PERSON SERVED"][0]
    assert locations["Text5"][0] == locations["PLAINTIFF STREET ADDRESS (provider)"][0]


@pytest.mark.parametrize("count", [7, 12, 24, 60, 1000])
def test_measured_page_plan_has_no_overlapping_blocks_or_footer_intrusions(count):
    data = long_case(count)
    pages = plan_pages(fields_for_case(data),count)
    row_numbers = []
    for page in pages:
        bottom = 0
        for item in page:
            height = item["height"] if item["kind"]=="row" else item["bottom"]-item["top"]
            assert item["at"] >= bottom
            bottom = item["at"]+height
            assert bottom <= BODY_BOTTOM
            if item["kind"]=="row": row_numbers.append(item["number"])
    assert row_numbers == list(range(1,count+1))


def test_long_period_wraps_in_its_own_column_without_clipping_or_smaller_type():
    data = long_case(12)
    label = "March 1 through March 31, 2021 unpaid rent for the dwelling under the rental agreement"
    data["amounts_due"][2]["due_date"] = label
    pdf = PdfReader(BytesIO(render_residential_pdf(data)))
    fields = pdf.get_fields()
    labels = [fields[k]["/V"] for k in fields if k=="Month3" or k.startswith("Month3 (")]
    assert " ".join(labels) == label
    assert fields["AmountDue3"]["/V"] == "$1,000.00"
    for page in pdf.pages:
        for ref in page["/Annots"]:
            widget=ref.get_object()
            if str(widget.get("/T", "")).startswith(("Month","AmountDue")):
                assert widget["/DA"] == "/TiRo 12 Tf 0 g"
                assert widget["/Ff"] == 0  # Native PDFKit must not regenerate a multiline AP.
                assert b"/TiRo 12 Tf" in widget["/AP"]["/N"].get_data()
    assert static_text(pdf) == static_text(PdfReader(MASTER))


def test_extremely_long_period_can_continue_across_pages_without_losing_lines():
    data = long_case(1)
    data["amounts_due"][0]["due_date"] = "January monthly rent " * 150
    reader = PdfReader(BytesIO(render_residential_pdf(data)))
    fields=reader.get_fields()
    labels=[f["/V"] for k,f in fields.items() if k=="Month1" or k.startswith("Month1 (")]
    assert " ".join(labels) == data["amounts_due"][0]["due_date"].strip()
    assert len([k for k in fields if k.startswith("AmountDue")]) == 1
    assert static_text(reader) == static_text(PdfReader(MASTER))


def test_saved_word_edit_preserves_all_twelve_rows_and_refreshes_the_pdf():
    data=long_case(12)
    word=render_residential_notice(data)
    root=etree.fromstring(zipfile.ZipFile(BytesIO(word)).read("word/document.xml"))
    table=inspect_standard(root)["rent_tables"][0]
    assert len(table["rows"]) == 12
    first=table["rows"][0];total=table["total"]
    edited=patch_docx(word,[{"old":first["text"],"new":"$1,250.25","paragraph_index":first["paragraph_index"]},
                            {"old":total["text"],"new":"$12,250.25","paragraph_index":total["paragraph_index"]}])
    result=render_saved_word_pdf(edited,sha256(edited).hexdigest())
    assert len(result["case_fields"]["amounts_due"]) == 12
    assert result["case_fields"]["total_amount_due"] == "$12,250.25"
    assert base64.b64decode(result["pdf_base64"]) == render_residential_pdf(result["case_fields"])
    assert render_saved_word_pdf(edited,sha256(edited).hexdigest()) == result


@pytest.mark.parametrize("count", [1,3,6])
def test_existing_short_form_bytes_do_not_change(count):
    data=case(count)
    assert render_residential_pdf(data) == _render_original(fields_for_case(data))


def test_resource_limit_never_silently_truncates_rows():
    with pytest.raises(NoticeInputError,match="no entries may be truncated"):
        render_residential_pdf(long_case(1001))
