"""Measured continuation layout for rent that outgrows the original six boxes.

Reuse the original PDF's artwork, text operators, fonts and paragraph line
breaks. Only the variable rent table grows. Whole boilerplate paragraphs and
the complete signature/contact block move together when a page fills up.
All geometry is in PDF points; top coordinates increase down the page.
"""
from io import BytesIO
from math import floor

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, ContentStream, DecodedStreamObject,
    DictionaryObject, FloatObject, NameObject, NumberObject, TextStringObject,
    BooleanObject)

from residential_pdf import PDF_TEMPLATE_VERSION, _bytes, _width
from residential_notice import REFERENCE_SHA256

RENDERER_VERSION = "rent-pagination-2026-10-27"
PAGE_WIDTH, PAGE_HEIGHT = 612, 792
BODY_TOP, BODY_BOTTOM = 98.0, 754.0
ROW_LEADING, ROW_PADDING = 14.4, 2.5
PERIOD_X, PERIOD_WIDTH = 176.631, 164.8
AMOUNT_X, AMOUNT_WIDTH = 357.864, 107.57


def wrap_period(text):
    """Wrap to the period column at the original 12pt font, never shrink it."""
    width = PERIOD_WIDTH - 4
    lines, current = [], ""
    for word in text.split():
        candidate = (current + " " + word).strip()
        if _width(candidate, "TiRo") <= width:
            current = candidate
            continue
        if current:
            lines.append(current)
            current = ""
        for char in word:
            if _width(current + char, "TiRo") > width:
                lines.append(current)
                current = ""
            current += char
    if current:
        lines.append(current)
    return lines or [""]


def plan_pages(values, row_count):
    pages = [[{"kind": "segment", "source": 0, "top": 92, "bottom": 312,
               "at": 92, "fields": True, "label": "premises and introduction"}]]
    cursor = 312.0

    def new_page():
        nonlocal cursor
        pages.append([])
        cursor = BODY_TOP

    def segment(source, top, bottom, gap=0, label="", fields=True):
        nonlocal cursor
        height = bottom - top
        if cursor + gap + height > BODY_BOTTOM:
            new_page()
            gap = 0
        pages[-1].append({"kind": "segment", "source": source, "top": top,
                         "bottom": bottom, "at": cursor+gap, "fields": fields, "label": label})
        cursor += gap + height

    def heading():
        segment(0, 312, 327.5, label="rent column headings", fields=False)

    heading()
    cursor += 1
    for number in range(1, row_count+1):
        lines = wrap_period(values[f"PERIOD_{number}"])
        reserve = 22 if number == row_count else 0  # final row + total stay together
        height = len(lines)*ROW_LEADING + ROW_PADDING
        if (cursor + height + reserve > BODY_BOTTOM and
                height + reserve <= BODY_BOTTOM-BODY_TOP-16.5):
            new_page()
            heading()
            cursor += 1
        part = 1
        while lines:
            capacity = floor((BODY_BOTTOM-cursor-ROW_PADDING-reserve)/ROW_LEADING)
            if capacity < 1:
                new_page()
                heading()
                cursor += 1
                continue
            shown, lines = lines[:capacity], lines[capacity:]
            h = len(shown)*ROW_LEADING + ROW_PADDING
            pages[-1].append({"kind": "row", "number": number, "part": part,
                             "lines": shown, "amount": values[f"AMOUNT_{number}"] if part == 1 else "",
                             "at": cursor, "height": h})
            cursor += h
            part += 1
            if lines:
                new_page()
                heading()
                cursor += 1
    segment(0, 412, 432, gap=2, label="total after final rent entry")
    for source, top, bottom, gap, label in [
        (0, 440, 481, 8, "payment or possession paragraph"),
        (0, 494, 535, 13, "service-date paragraph"),
        (0, 544, 630, 9, "payment recipient and contact"),
        (0, 630, 658, 0, "payment business hours"),
        (0, 671, 754, 13, "payment conditions paragraph"),
        (1, 99, 127, 16, "forfeiture paragraph"),
        (1, 141, 223, 14, "California heading and credit paragraph"),
        (1, 237, 333, 14, "abandoned property paragraph"),
        (1, 348, 389, 15, "unlawful detainer paragraph"),
        (1, 403, 444, 14, "superseded notices paragraph"),
        (1, 472, 602, 28, "signature and housing provider contact"),
    ]:
        segment(source, top, bottom, gap, label)
    return pages


def _multiply(m, n):
    a,b,c,d,e,f = m; g,h,i,j,k,l = n
    return [a*g+c*h, b*g+d*h, a*i+c*j, b*i+d*j, a*k+c*l+e, b*k+d*l+f]


def _segment_stream(page, reader, top, bottom):
    """Remove out-of-band text too: clipping alone leaves duplicate hidden text."""
    stream = ContentStream(page.get_contents(), reader)
    ctm = [1,0,0,1,0,0]; tm = [1,0,0,1,0,0]; line = tm[:]; stack = []
    kept = []
    for operands, op in stream.operations:
        v = operands
        if op == b"q": stack.append(ctm[:])
        elif op == b"Q": ctm = stack.pop()
        elif op == b"cm": ctm = _multiply(ctm, list(map(float,v)))
        elif op == b"BT": tm = [1,0,0,1,0,0]; line = tm[:]
        elif op == b"Tm": tm = list(map(float,v)); line = tm[:]
        elif op in (b"Td", b"TD"):
            line = _multiply(line, [1,0,0,1,float(v[0]),float(v[1])]); tm = line[:]
        if op in (b"Tj", b"TJ"):
            y = ctm[1]*tm[4] + ctm[3]*tm[5] + ctm[5]
            if not top <= PAGE_HEIGHT-y < bottom:
                continue
        kept.append((operands, op))
    stream.operations = kept
    return stream.get_data()


def render_flowing_pdf(base_pdf, values, row_count):
    source = PdfReader(BytesIO(base_pdf))
    writer = PdfWriter()
    source_form = source.trailer["/Root"]["/AcroForm"]
    fonts = source_form["/DR"]["/Font"].clone(writer)
    fields, used_names, forms = ArrayObject(), set(), {}
    form = DictionaryObject({NameObject("/Fields"): fields,
        NameObject("/DR"): DictionaryObject({NameObject("/Font"): fonts}),
        NameObject("/NeedAppearances"): BooleanObject(False)})
    writer._root_object[NameObject("/AcroForm")] = writer._add_object(form)

    def add_widget(page, widget, name, dy=0):
        parent = widget.get("/Parent", widget).get_object()
        result = DictionaryObject()
        for obj in (parent, widget):
            for k,v in obj.items():
                if k not in ("/Parent", "/Kids", "/P", "/Rect", "/T"):
                    result[NameObject(k)] = v.clone(writer)
        if name in used_names:
            name += " (provider)"
        if name in used_names:
            raise RuntimeError("Repeated continuation PDF field name")
        used_names.add(name)
        result[NameObject("/T")] = TextStringObject(name)
        result[NameObject("/Rect")] = ArrayObject([FloatObject(float(x)+(dy if i%2 else 0)) for i,x in enumerate(widget["/Rect"])])
        result[NameObject("/P")] = page.indirect_reference
        ref = writer._add_object(result)
        page["/Annots"].append(ref); fields.append(ref)

    def original_segment(page, content, source_index, top, bottom, at, include_fields=True):
        key = source_index, top, bottom
        if key not in forms:
            original = source.pages[source_index]
            obj = DecodedStreamObject()
            obj.set_data(_segment_stream(original, source, top, bottom))
            obj.update({NameObject("/Type"):NameObject("/XObject"), NameObject("/Subtype"):NameObject("/Form"),
                NameObject("/BBox"):ArrayObject([FloatObject(x) for x in [0,0,PAGE_WIDTH,PAGE_HEIGHT]]),
                NameObject("/Resources"):original["/Resources"].clone(writer)})
            forms[key] = writer._add_object(obj)
        name = NameObject("/Part"+str(len(page["/Resources"]["/XObject"])))
        page["/Resources"]["/XObject"][name] = forms[key]
        dy = top-at
        content.append(f"q 0 {PAGE_HEIGHT-at-(bottom-top):.6f} {PAGE_WIDTH} {bottom-top:.6f} re W n 1 0 0 1 0 {dy:.6f} cm {name} Do Q\n")
        if include_fields:
            for ref in source.pages[source_index].get("/Annots", []):
                widget = ref.get_object()
                if top <= PAGE_HEIGHT-float(widget["/Rect"][3]) < bottom:
                    parent = widget.get("/Parent",widget).get_object()
                    add_widget(page, widget, str(parent["/T"]), dy)

    def row_field(page, name, text, lines, x, width, at, height):
        ap = DecodedStreamObject()
        text_ops = [f"/Tx BMC q 1 1 {width-2:.6f} {height-2:.6f} re W n BT /TiRo 12 Tf 0 g\n"]
        for i,line in enumerate(lines):
            text_ops.append(f"1 0 0 1 2 {height-11-i*ROW_LEADING:.6f} Tm <{_bytes(line).hex()}> Tj\n")
        text_ops.append("ET Q EMC\n")
        ap.set_data("".join(text_ops).encode("ascii"))
        ap.update({NameObject("/Type"):NameObject("/XObject"),NameObject("/Subtype"):NameObject("/Form"),
            NameObject("/FormType"):NumberObject(1),
            NameObject("/Matrix"):ArrayObject([NumberObject(v) for v in [1,0,0,1,0,0]]),
            NameObject("/BBox"):ArrayObject([FloatObject(v) for v in [0,0,width,height]]),
            NameObject("/Resources"):DictionaryObject({NameObject("/Font"):fonts})})
        widget = DictionaryObject({NameObject("/Type"):NameObject("/Annot"),NameObject("/Subtype"):NameObject("/Widget"),
            NameObject("/FT"):NameObject("/Tx"),NameObject("/T"):TextStringObject(name),
            NameObject("/V"):TextStringObject(text),NameObject("/DV"):TextStringObject(text),
            NameObject("/DA"):TextStringObject("/TiRo 12 Tf 0 g"),NameObject("/F"):NumberObject(4),
            NameObject("/Ff"):NumberObject(4096 if len(lines)>1 else 0),NameObject("/MK"):DictionaryObject(),NameObject("/P"):page.indirect_reference,
            NameObject("/Rect"):ArrayObject([FloatObject(v) for v in [x,PAGE_HEIGHT-at-height,x+width,PAGE_HEIGHT-at]]),
            NameObject("/AP"):DictionaryObject({NameObject("/N"):writer._add_object(ap)})})
        ref=writer._add_object(widget);fields.append(ref);page["/Annots"].append(ref);used_names.add(name)

    plans = plan_pages(values, row_count)
    for page_index, placements in enumerate(plans):
        page=writer.add_blank_page(PAGE_WIDTH,PAGE_HEIGHT)
        page[NameObject("/Resources")]=DictionaryObject({NameObject("/XObject"):DictionaryObject()})
        page[NameObject("/Annots")]=ArrayObject()
        content=[]
        original_segment(page,content,0 if page_index==0 else 1,0,89,0,include_fields=page_index==0)
        original_segment(page,content,1,754,792,754,include_fields=False)
        for item in placements:
            if item["kind"]=="segment":
                original_segment(page,content,item["source"],item["top"],item["bottom"],item["at"],item["fields"])
            else:
                number,part=item["number"],item["part"]
                suffix="" if part==1 else f" (part {part})"
                # Separate line widgets avoid PDFKit's multiline appearance
                # regeneration clipping longer values. Canonical values and
                # visible appearances agree for every line in every viewer.
                for line_index, line in enumerate(item["lines"]):
                    line_suffix = suffix + (f" (line {line_index+1})" if line_index else "")
                    row_field(page,f"Month{number}"+line_suffix,line,[line],PERIOD_X,PERIOD_WIDTH,
                              item["at"]+line_index*ROW_LEADING,ROW_LEADING+ROW_PADDING)
                if item["amount"]:
                    row_field(page,f"AmountDue{number}",item["amount"],[item["amount"]],AMOUNT_X,AMOUNT_WIDTH,item["at"],ROW_LEADING+ROW_PADDING)
        stream=DecodedStreamObject();stream.set_data("".join(content).encode("ascii"))
        page[NameObject("/Contents")]=writer._add_object(stream)
    writer.add_metadata({"/NoticeTemplateVersion":PDF_TEMPLATE_VERSION,"/NoticeTemplateSHA256":REFERENCE_SHA256,
                         "/NoticeRentRenderer":RENDERER_VERSION,"/NoticeRentRows":str(row_count)})
    out=BytesIO();writer.write(out)
    return out.getvalue()
