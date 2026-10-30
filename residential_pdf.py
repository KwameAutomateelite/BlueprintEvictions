"""Fill Anne-Michelle's original PDF; paginate only an overflowing rent list.

PDF coordinates are points from the lower-left corner (72 points/inch).
For a fitting schedule only variable appearances and the approved heading/date
typography change. An overflowing
schedule uses measured continuation pages with the source's paragraph fragments.
"""
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject, BooleanObject, DecodedStreamObject, DictionaryObject,
    NameObject, TextStringObject, FloatObject, NumberObject, ByteStringObject, ContentStream,
)

from residential_notice import NoticeInputError, REFERENCE_SHA256, fields_for_case

PDF_TEMPLATE_VERSION = "residential-original-pdf-2026-10-28-format3"
MASTER_SHA256 = "cf06f18af3bc79720a91a99afbcd84c5214aa613aa63608feb10508ff0ef989a"
MASTER = Path(__file__).parent / "templates/residential-pay-or-quit-original.pdf"
WIDTHS = json.loads((MASTER.parent / "residential-pdf-font-widths.json").read_text())

# Keys include the page because the original links payment and provider
# contacts together. Separate logical fields allow distinct contacts without
# moving either widget. The source's misleading 'NAME OF PERSON SERVED' is
# the date next to the signature, not a tenant-name field.
FIELD_MAP = {
    (1, "DEFENDANT"): "TENANT_NAMES",
    (1, "DEFENDANT STREET ADDRESS"): "PROPERTY_STREET",
    (1, "DEFENDANT CITY/STATE/ZIP"): "PROPERTY_CITY",
    (1, "COUNTY OF"): "COUNTY",
    (1, "TOTAL RENT DUE"): "TOTAL_AMOUNT",
    (1, "PLAINTIFF"): "PAYMENT_RECIPIENT",
    (1, "PLAINTIFF STREET ADDRESS"): "PAYMENT_STREET",
    (1, "PLAINTIFF CITY/STATE/ZIP"): "PAYMENT_CITY",
    (1, "Text6"): "PAYMENT_PHONE",
    (2, "NAME OF PERSON SERVED"): "NOTICE_DATE",
    (2, "Text5"): "SIGNER_NAME_AND_TITLE",
    (2, "PLAINTIFF STREET ADDRESS"): "PROVIDER_STREET",
    (2, "PLAINTIFF CITY/STATE/ZIP"): "PROVIDER_CITY",
    (2, "Text6"): "PROVIDER_PHONE",
    (2, "cc"): "CC",
    **{(1, f"Month{i}"): f"PERIOD_{i}" for i in range(1, 7)},
    **{(1, f"AmountDue{i}"): f"AMOUNT_{i}" for i in range(1, 7)},
}


def _bytes(text):
    try:
        return text.encode("cp1252")
    except UnicodeEncodeError as exc:
        raise NoticeInputError("A case field contains characters unavailable in the original form font; human review required") from exc


def _width(text, font):
    return sum(WIDTHS[font][code] for code in _bytes(text)) * 12 / 1000


def _lines(text, width, font, multiline):
    # Exact source font advances for the fixed case fields. Rent period text
    # uses the separate paginator when it exceeds the original rent boxes.
    if not multiline:
        if _width(text, font) > width - 4:
            raise NoticeInputError("A case field exceeds its original PDF box; human review required")
        return [text]
    lines = [""]
    for word in text.split():
        if _width(word, font) > width - 4:
            raise NoticeInputError("A case field contains a word wider than its original PDF box")
        candidate = (lines[-1] + " " + word).strip()
        if _width(candidate, font) > width - 4:
            if len(lines) == 2:
                raise NoticeInputError("A case field exceeds the original two-line PDF box; human review required")
            lines.append(word)
        else:
            lines[-1] = candidate
    return lines


def _appearance(writer, widget, field, text, fonts):
    x0, y0, x1, y1 = map(float, widget["/Rect"])
    width, height = x1 - x0, y1 - y0
    da = str(widget.get("/DA", field.get("/DA")))
    font = re.search(r"/(TiRo|TiBo) 12 Tf", da).group(1)
    multiline = bool(int(widget.get("/Ff", field.get("/Ff", 0))) & 4096)
    lines = _lines(text, width, font, multiline)
    old_ap = widget.get("/AP", {}).get("/N")
    old_ap = old_ap.get_object() if old_ap else None
    # Reuse the source baseline exactly, including its fractional points.
    # Blank source boxes use the same original Acrobat 12pt Times formula.
    baseline = height - 15.392 if multiline else height / 2 - 4.092
    if old_ap is not None:
        match = re.search(rb"([-\d.]+) ([-\d.]+) Td", old_ap.get_data())
        if match:
            baseline = float(match[2])
    if field.get("/T") == "NAME OF PERSON SERVED":
        # Align the entered date with both the fixed Date: caption and the
        # signature caption, whose approved baseline is y=296.57 points.
        baseline = 296.57 - y0
    parts = [f"/Tx BMC\nq\n1 1 {width-2:.6f} {height-2:.6f} re W n\nBT\n/{font} 12 Tf\n0 g\n"]
    for i, line in enumerate(lines):
        # Q=1 means centered within this fixed box, not centered on the page.
        x = (width - _width(line, font)) / 2 if widget.get("/Q", field.get("/Q")) == 1 else 2
        y = baseline - i * 14.4
        parts.append(f"1 0 0 1 {x:.6f} {y:.6f} Tm\n<{_bytes(line).hex()}> Tj\n")
    parts.append("ET\nQ\nEMC\n")
    ap = DecodedStreamObject()
    ap.set_data("".join(parts).encode("ascii"))
    ap.update({
        NameObject("/Type"): NameObject("/XObject"),
        NameObject("/Subtype"): NameObject("/Form"),
        NameObject("/FormType"): NumberObject(1),
        NameObject("/BBox"): ArrayObject([FloatObject(v) for v in (0, 0, width, height)]),
        NameObject("/Resources"): DictionaryObject({NameObject("/Font"): fonts}),
    })
    widget[NameObject("/AP")] = DictionaryObject({NameObject("/N"): writer._add_object(ap)})


def _render_original(values, template_path=MASTER):
    source = Path(template_path).read_bytes()
    if sha256(source).hexdigest() != MASTER_SHA256:
        raise RuntimeError("Original approved PDF changed; do not substitute a reconstructed form")
    values = dict(values)
    values["TENANT_NAMES"] = " ".join(filter(None, (values["TENANT_NAMES"], values["TENANT_NAMES_CONTINUED"])))
    values["SIGNER_NAME_AND_TITLE"] = " ".join(filter(None, (values["SIGNER_NAME_AND_TITLE"], values["SIGNER_CONTINUED"])))
    values["CC"] = ""
    reader = PdfReader(BytesIO(source))
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    if len(writer.pages) != 2:
        raise RuntimeError("The approved PDF must have exactly two pages")
    form = writer._root_object["/AcroForm"]
    fonts = form["/DR"]["/Font"]
    embedded = PdfReader(MASTER.parent / "residential-input-fonts.pdf")
    for key in ("/TiRo", "/TiBo"):
        fonts[NameObject(key)] = embedded.pages[0]["/Resources"]["/Font"][key].clone(writer)
    # Correct only this heading's size/underline, leaving the approved body,
    # page geometry and remaining fixed artwork untouched.
    page = writer.pages[0]
    page["/Resources"]["/Font"][NameObject("/AMBold12")] = fonts["/TiBo"]
    stream = ContentStream(page.get_contents(), writer)
    heading = False
    changed = 0
    operations = []
    for operands, op in stream.operations:
        if op == b"Td" and len(operands) == 2 and abs(float(operands[0])-179.905)<.001 and abs(float(operands[1])-469.549)<.001:
            heading = True
        if heading and op == b"Tj":
            operations.extend([([NameObject("/AMBold12"), NumberObject(12)], b"Tf"), ([NumberObject(0)], b"Tw"),
                               ([ByteStringObject(b"TIME PERIOD DUE")], b"Tj"),
                               ([NameObject("/C2_0"), FloatObject(11.52)], b"Tf")])
            heading = False
            changed += 1
        else:
            operations.append((operands, op))
    if changed != 1:
        raise RuntimeError("Approved rent heading could not be identified")
    stream.operations = operations
    heading_stream = DecodedStreamObject()
    end = 179.905 + _width("TIME PERIOD DUE", "TiBo")
    heading_stream.set_data(stream.get_data() + f"\nq 0 G 0.6 w 179.905 467.2 m {end:.6f} 467.2 l S Q\n".encode())
    page[NameObject("/Contents")] = writer._add_object(heading_stream)
    fields, seen = form["/Fields"], set()
    updated = {}
    for page_no, page in enumerate(writer.pages, 1):
        for ref in page["/Annots"]:
            widget = ref.get_object()
            if widget.get("/Subtype") != "/Widget":
                continue
            parent = widget.get("/Parent")
            parent = parent.get_object() if parent else widget
            name = str(parent.get("/T"))
            if parent.get("/FT") == "/Btn":
                continue
            key = FIELD_MAP[(page_no, name)]
            seen.add((page_no, name))
            text = values[key]
            field = parent
            parent_id = id(parent)
            if parent_id in updated and updated[parent_id] != text:
                # Only split a linked contact when the case explicitly supplies
                # different payment and provider details. Otherwise retain the
                # original field tree and appearances exactly.
                field = DictionaryObject({k: v for k, v in parent.items() if k != "/Kids"})
                field[NameObject("/Kids")] = ArrayObject([ref])
                field[NameObject("/T")] = TextStringObject(name + " (provider)")
                field_ref = writer._add_object(field)
                widget[NameObject("/Parent")] = field_ref
                parent[NameObject("/Kids")] = ArrayObject([kid for kid in parent["/Kids"] if kid != ref])
                fields.append(field_ref)
            field[NameObject("/V")] = TextStringObject(text)
            field[NameObject("/DV")] = TextStringObject(text)
            updated[parent_id] = text
            # Every variable uses an embedded, actual Times New Roman face at
            # 12pt, including a value identical to the original reference.
            font = "TiBo" if key == "TOTAL_AMOUNT" else "TiRo"
            field[NameObject("/DA")] = TextStringObject(f"/{font} 12 Tf 0 g")
            widget[NameObject("/DA")] = TextStringObject(f"/{font} 12 Tf 0 g")
            _appearance(writer, widget, field, text, fonts)
    if seen != set(FIELD_MAP):
        raise RuntimeError("Original PDF field map is incomplete")
    form[NameObject("/Fields")] = fields
    form[NameObject("/NeedAppearances")] = BooleanObject(False)
    writer.add_metadata({"/NoticeTemplateVersion": PDF_TEMPLATE_VERSION, "/NoticeTemplateSHA256": REFERENCE_SHA256})
    out = BytesIO()
    writer.write(out)
    # Clone only the reachable final tree. Replaced appearances and detached
    # source parents must not leave the reference case's values in the file.
    final = PdfWriter()
    final.clone_document_from_reader(PdfReader(BytesIO(out.getvalue())))
    final.add_metadata({"/NoticeTemplateVersion": PDF_TEMPLATE_VERSION, "/NoticeTemplateSHA256": REFERENCE_SHA256})
    out = BytesIO()
    final.write(out)
    return out.getvalue()


def render_residential_pdf(data, template_path=MASTER):
    values = fields_for_case(data)
    rows = len(data["amounts_due"])
    # The six original boxes have slightly different measured widths. Keep
    # short notices byte-for-byte on that original path; long labels or lists
    # use the same artwork and paragraph fragments with measured pagination.
    from residential_word_layout import needs_flow
    overflow = needs_flow(values, rows)
    if not overflow:
        return _render_original(values, template_path)
    base_values = dict(values)
    for i in range(1, 7):
        base_values[f"PERIOD_{i}"] = base_values[f"AMOUNT_{i}"] = ""
    from residential_pdf_flow import render_flowing_pdf
    return render_flowing_pdf(_render_original(base_values, template_path), values, rows)
