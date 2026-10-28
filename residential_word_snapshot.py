"""Render the saved Word editing copy into the original, fixed-layout PDF.

Read only named form fields. Reject changed boilerplate, lost bindings,
unreconciled rent or an incorrect source hash. No network or signing occurs.
"""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, BadZipFile
import re

from docx import Document

from residential_notice import MASTER as WORD_MASTER, NoticeInputError
from residential_pdf import PDF_TEMPLATE_VERSION, render_residential_pdf

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FIELD_NAMES = {
    "TENANT_NAMES", "PROPERTY_STREET", "PROPERTY_CITY", "COUNTY",
    "PAYMENT_RECIPIENT", "PAYMENT_STREET", "PAYMENT_CITY", "PAYMENT_PHONE",
    "NOTICE_DATE", "SIGNER_NAME_AND_TITLE", "PROVIDER_STREET", "PROVIDER_CITY", "PROVIDER_PHONE",
    *{f"RENT_ROW_{i}" for i in range(1, 7)},
}


def _text(p):
    return "".join((n.text or "") if n.tag == W + "t" else " " if n.tag == W + "br" else "" for n in p.iter())


def _bindings(doc):
    result = {}
    for i, p in enumerate(doc.element.iter(W + "p")):
        for bookmark in p.iter(W + "bookmarkStart"):
            name = bookmark.get(W + "name", "")
            if name.startswith("BE_"):
                key = name[3:]
                if key in result:
                    raise NoticeInputError("Repeated saved Word form field; fresh review required")
                result[key] = (i, p)
    return result


def _total(doc):
    matches = []
    for table in doc.element.iter(W + "tbl"):
        rows = table.findall(W + "tr")
        if len(rows) == 1:
            cells = rows[0].findall(W + "tc")
            if len(cells) == 2 and _text(cells[0]).strip() == "TOTAL AMOUNT DUE:":
                paragraphs = list(cells[1].iter(W + "p"))
                if len(paragraphs) == 1:
                    matches.append(paragraphs[0])
    if len(matches) != 1:
        raise NoticeInputError("Missing or repeated saved Word total; fresh review required")
    return matches[0]


def saved_word_case(content):
    """Return None for other notice families; malformed standard forms fail closed."""
    try:
        with ZipFile(BytesIO(content)) as z:
            if len(z.infolist()) > 1000 or sum(i.file_size for i in z.infolist()) > 20_000_000:
                raise NoticeInputError("Saved Word file exceeds notice size limits")
        doc = Document(BytesIO(content))
    except (BadZipFile, KeyError, ValueError) as exc:
        raise NoticeInputError("Invalid saved Word document") from exc
    bindings = _bindings(doc)
    identifier = doc.core_properties.identifier or ""
    if not bindings and not identifier.startswith("residential-pay-or-quit-"):
        return None
    if set(bindings) != FIELD_NAMES or not identifier.startswith("residential-pay-or-quit-"):
        raise NoticeInputError("Saved residential Word form lost its field bindings; fresh review required")
    master = Document(WORD_MASTER)
    expected = _bindings(master)
    a, b = list(master.element.iter(W + "p")), list(doc.element.iter(W + "p"))
    if len(a) != len(b) or any(bindings[k][0] != expected[k][0] for k in FIELD_NAMES):
        raise NoticeInputError("Saved Word form structure changed; regenerate the review copy")
    total = _total(doc)
    allowed = {i for i, _ in bindings.values()} | {b.index(total)}
    if any(_text(x) != _text(y) for i, (x, y) in enumerate(zip(a, b)) if i not in allowed):
        raise NoticeInputError("Saved Word boilerplate changed; human review required")
    values = {k: _text(p).strip() for k, (_, p) in bindings.items()}
    if not values["TENANT_NAMES"].startswith("To: ") or not values["COUNTY"].startswith("COUNTY OF "):
        raise NoticeInputError("Saved Word form labels changed")
    match = re.fullmatch(r"Date:\s*([0-9/]+)\s*\(Original signed by agent/owner\)", values["NOTICE_DATE"])
    if not match:
        raise NoticeInputError("Saved notice date or signature wording changed")
    signer, comma, role = values["SIGNER_NAME_AND_TITLE"].partition(",")
    if not comma or not signer.strip() or not role.strip():
        raise NoticeInputError("Saved notice needs its designated signer name and title")
    if not values["PROVIDER_PHONE"].startswith("Phone: "):
        raise NoticeInputError("Saved housing-provider phone label changed")
    rows, blank = [], False
    for i in range(1, 7):
        p = bindings[f"RENT_ROW_{i}"][1]
        tokens = [""]
        for n in p.iter():
            if n.tag == W + "t":
                tokens[-1] += n.text or ""
            elif n.tag == W + "tab" and n.getparent().tag == W + "r":
                tokens.append("")
        if len(tokens) != 2:
            raise NoticeInputError("Saved rent period and amount must remain separate columns")
        if not any(tokens):
            blank = True
        elif blank or not all(tokens):
            raise NoticeInputError("Saved rent row is incomplete or follows an empty row")
        else:
            rows.append({"due_date": tokens[0], "amount": tokens[1]})
    return {
        "notice_type": "residential", "day_count": 3,
        "tenant_names": values["TENANT_NAMES"][4:],
        "property_address_street": values["PROPERTY_STREET"],
        "property_address_city": values["PROPERTY_CITY"],
        "county": values["COUNTY"][10:],
        "amounts_due": rows, "total_amount_due": _text(total),
        "notice_date": match[1],
        "landlord_name": values["PAYMENT_RECIPIENT"],
        "payment_recipient_name": values["PAYMENT_RECIPIENT"],
        "payment_address_street": values["PAYMENT_STREET"],
        "payment_address_city": values["PAYMENT_CITY"],
        "payment_phone": values["PAYMENT_PHONE"],
        "signer_name": signer.strip(), "signer_title": role.strip(),
        "housing_provider_street": values["PROVIDER_STREET"],
        "housing_provider_city": values["PROVIDER_CITY"],
        "housing_provider_phone": values["PROVIDER_PHONE"][7:],
    }


def render_saved_word_pdf(content, expected_sha256):
    digest = sha256(content).hexdigest()
    if digest != expected_sha256:
        raise NoticeInputError("Saved Word source hash changed; fresh review required")
    case = saved_word_case(content)
    if case is None:
        return {"is_standard": False, "source_docx_sha256": digest}
    pdf = render_residential_pdf(case)
    import base64
    return {"is_standard": True, "source_docx_sha256": digest,
            "template_version": PDF_TEMPLATE_VERSION,
            "case_fields": case,
            "pdf_base64": base64.b64encode(pdf).decode("ascii"),
            "pdf_sha256": sha256(pdf).hexdigest()}
