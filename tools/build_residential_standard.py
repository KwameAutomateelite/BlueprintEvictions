"""Build the blank, editable Word master from the approved October 2026 form.

The retained legacy package supplies only its existing Blueprint artwork and
footer. All case slots are blank placeholders, never values from the example.
Run locally when the approved master changes; generation does not run this tool.
"""
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "templates/3-Day Notice - BLUEPRINT - residential - NEW BRANDING_071125_v2.docx"
TARGET = ROOT / "templates/residential-pay-or-quit-2026-10.docx"


def build():
    doc = Document(SOURCE)
    artwork = deepcopy(doc.paragraphs[0]._p)
    footer_part = doc.sections[0].footer.part
    footer_run = deepcopy(footer_part.element.find(".//" + qn("w:r")))
    for blip in footer_run.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
        rid = blip.get(qn("r:embed"))
        blip.set(qn("r:embed"), doc.part.relate_to(footer_part.rels[rid].target_part, RT.IMAGE))
    sect = deepcopy(doc.sections[-1]._sectPr)
    for node in list(doc._element.body):
        doc._element.body.remove(node)
    doc._element.body.append(sect)
    for ref in list(sect.findall(qn("w:footerReference"))):
        sect.remove(ref)
    pg = sect.find(qn("w:pgMar"))
    pg.set(qn("w:bottom"), "720")
    # Page geometry/artwork/footer are inherited from the source Word package.
    doc._element.body.insert(0, artwork)
    # The reference footer is artwork anchored to its first page. Put that
    # artwork on the retained page anchor so Word/LibreOffice agree even when
    # a footer's tiny inherited paragraph height is collapsed.
    artwork.append(deepcopy(footer_run))
    from docx.text.paragraph import Paragraph
    ap = Paragraph(artwork, doc._body)
    ap.paragraph_format.space_before = Pt(0)
    ap.paragraph_format.space_after = Pt(0)
    ap.paragraph_format.line_spacing = Pt(0.05)
    cursor = 0.05

    def block(text, y, *, x=72, width=480, size=11.5, leading=13.75,
              bold=False, bold_prefix=None, underline=False, align=None,
              slot=None, new_page=False, border=False):
        nonlocal cursor
        if new_page:
            p = doc.add_paragraph()
            p.paragraph_format.page_break_before = True
            p.paragraph_format.line_spacing = Pt(0.05)
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
            p.add_run().font.size = Pt(0.5)
            cursor = 0.05
        gap = y - cursor
        if gap < -0.1:
            raise ValueError(f"Overlapping master block: {slot or text[:30]} ({gap})")
        if gap > 0:
            spacer = doc.add_paragraph()
            sp = spacer.paragraph_format
            sp.space_before = sp.space_after = Pt(0)
            sp.line_spacing = Pt(gap)
            sp.keep_with_next = False
            spacer.add_run().font.size = Pt(0.5)
        p = doc.add_paragraph(style="Normal")
        pf = p.paragraph_format
        pf.left_indent = Pt(x - 54)
        pf.right_indent = Pt(594 - x - width)
        pf.first_line_indent = Pt(0)
        pf.space_before = pf.space_after = Pt(0)
        pf.line_spacing = Pt(leading)
        pf.keep_together = True
        pf.keep_with_next = False
        pf.widow_control = False
        p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.LEFT
        if slot:
            bookmark = OxmlElement("w:bookmarkStart")
            bookmark.set(qn("w:id"), str(len(doc.paragraphs)))
            bookmark.set(qn("w:name"), "BE_" + slot)
            p._p.append(bookmark)
        chunks = [(text, bold)]
        if bold_prefix and text.startswith(bold_prefix):
            chunks = [(bold_prefix, True), (text[len(bold_prefix):], False)]
        for value, weight in chunks:
            r = p.add_run(value)
            r.bold = weight
            r.underline = underline
            r.font.name = "Times New Roman"
            r.font.size = Pt(size)
            r.font.color.rgb = RGBColor(0, 0, 0)
        if slot:
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), bookmark.get(qn("w:id")))
            p._p.append(end)
        if border:
            borders = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            for key, val in {"val": "single", "sz": "12", "space": "0", "color": "000000"}.items():
                bottom.set(qn("w:" + key), val)
            borders.append(bottom)
            p._p.get_or_add_pPr().append(borders)
        cursor = y + leading * (text.count("\n") + 1)
        return p

    block("THREE (3) DAY NOTICE TO PAY OR QUIT", 100, size=16, leading=20,
          bold=True, x=72, width=480, align=WD_ALIGN_PARAGRAPH.CENTER)
    block("To: {{TENANT_NAMES}}\n{{TENANT_NAMES_CONTINUED}}", 130, size=12, leading=14, slot="TENANT_NAMES")
    block("Tenants in possession of the premises herein referred to are situated at:",
          158, leading=14)
    block("{{PROPERTY_STREET}}", 180, x=106, width=445, size=12, leading=14, slot="PROPERTY_STREET")
    block("{{PROPERTY_CITY}}", 194, x=106, width=445, size=12, leading=14, slot="PROPERTY_CITY")
    block("COUNTY OF {{COUNTY}}", 212, x=108, width=445, size=12, leading=14, slot="COUNTY")
    block('[hereinafter called the "Premises"] and all notices thereafter:', 227, leading=14)
    block("PLEASE TAKE NOTICE that pursuant to the terms of tenancy for the above-referenced\n"
          "agreement by which you hold the possession of the Premises described in this Notice, there is\n"
          "now due, unpaid, and delinquent rent in the following amounts as specified below:",
          254, leading=15.8, bold_prefix="PLEASE TAKE NOTICE")
    p = block("TIME PERIOD DUE\tAMOUNT DUE", 311, x=180, width=280,
              size=12, leading=14, bold=True, underline=True)
    p.paragraph_format.tab_stops.add_tab_stop(Pt(306))
    for i in range(1, 7):
        p = block("{{PERIOD_%s}}\t{{AMOUNT_%s}}" % (i, i), 327 + (i - 1) * 14,
                  x=180, width=280, size=12, leading=14, slot="RENT_ROW_%s" % i)
        p.paragraph_format.tab_stops.add_tab_stop(Pt(306))
    p = block("", 414, x=180, width=292, size=12, leading=16)
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width, table.columns[1].width = Pt(168), Pt(118)
    props = table._tbl.tblPr
    indent = OxmlElement("w:tblInd")
    indent.set(qn("w:w"), "2540")
    indent.set(qn("w:type"), "dxa")
    props.append(indent)
    table.rows[0].height = Pt(16)
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    for i, cell in enumerate(table.rows[0].cells):
        cell.width = Pt(168 if i == 0 else 118)
        cp = cell._tc.get_or_add_tcPr()
        margins = OxmlElement("w:tcMar")
        for edge in ("top", "bottom", "left", "right"):
            e = OxmlElement("w:" + edge)
            e.set(qn("w:w"), ("140" if i else "40") if edge in ("left", "right") else "0")
            e.set(qn("w:type"), "dxa")
            margins.append(e)
        cp.append(margins)
        borders = OxmlElement("w:tcBorders")
        for edge in ("top", "bottom", "left", "right"):
            e = OxmlElement("w:" + edge)
            e.set(qn("w:val"), "single" if i else "nil")
            e.set(qn("w:sz"), "8")
            e.set(qn("w:color"), "000000")
            borders.append(e)
        cp.append(borders)
        para = cell.paragraphs[0]
        pf = para.paragraph_format
        pf.space_before = pf.space_after = Pt(0)
        pf.line_spacing = Pt(14)
        pf.keep_with_next = False
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = para.add_run("TOTAL AMOUNT DUE:" if i == 0 else "{{TOTAL_AMOUNT}}")
        run.font.name, run.font.size, run.bold = "Times New Roman", Pt(12), i == 0
    p._p.addprevious(table._tbl)
    p._p.getparent().remove(p._p)
    block("You are further notified that within three (3) days after service of this notice on you, not including\n"
          "Saturdays, Sundays or other judicial holidays, you must pay the total amount of rent stated in this\n"
          "Notice in full or quit the Premises, meaning to deliver up possession of the Premises and vacate.",
          439, x=73.7, width=482, leading=13.75)
    service = block("This notice was served to you on _______________. You have three (3) days after the date the\n"
          "notice was served to you, excluding Saturdays, Sundays, and other judicial holidays, to pay the\n"
          "rent due or vacate the property.", 492, x=73.7, width=482, leading=13.75, bold=True)
    before, after = service.runs[0].text.split("days after", 1)
    service.runs[0].text = before
    for value, under in [("days after", True), (after, False)]:
        r = service.add_run(value)
        r.font.name, r.font.size, r.bold, r.underline = "Times New Roman", Pt(11.5), True, under
    block("Payment of the above total amount due should be tendered to, or possession of the premises\n"
          "delivered to:", 543.5, x=73.7, width=482, leading=13.75)
    block("{{PAYMENT_RECIPIENT}}", 571, x=157, width=420, size=12, leading=14, slot="PAYMENT_RECIPIENT")
    block("{{PAYMENT_STREET}}", 586, x=159, width=420, size=12, leading=14, slot="PAYMENT_STREET")
    block("{{PAYMENT_CITY}}", 600, x=159, width=420, size=12, leading=14, slot="PAYMENT_CITY")
    block("{{PAYMENT_PHONE}}", 616, x=160, width=415, size=12, leading=14, slot="PAYMENT_PHONE")
    block("Rent or possession of the Premises may be tendered between normal business hours of 9:00AM to\n"
          "5:00PM, Monday through Friday.", 630, x=73.7, width=482, leading=13.75)
    payment_terms = block("Only the total amount due and made prior to the expiration of this Notice may be made using cash or\n"
          "certified funds (e.g., cashier's check, money order, etc.). No other forms of payment are acceptable.\n"
          "Landlords, and the attorney on behalf of the Landlords, will not accept any payments after the\n"
          "expiration of this Notice. Further, be advised that Landlords will not accept any partial payments for the\n"
          "amounts outlined in this notice nor will they accept payments made to others ostensibly on their behalf:\n"
          "all payments must be made to the above-stated address.", 670.5, x=73.7, width=482, leading=13.75)
    final_line = "all payments must be made to the above-stated address."
    payment_terms.runs[0].text = payment_terms.runs[0].text[:-len(final_line)]
    final_run = payment_terms.add_run(final_line)
    final_run.font.name = "Times New Roman"
    final_run.font.size = Pt(11.5)
    final_run.underline = True
    block("You are further notified by this Notice the Landlords elect to and therefore do declare a forfeiture of\n"
          "the rental agreement if the rent stated in this notice is not paid in full within the 3 (three) days.",
          98, leading=13.75, new_page=True)
    block("NOTICES FROM THE STATE OF CALIFORNIA:", 140.5, x=78.4, width=475, bold=True)
    block("(A) NOTICE regarding possible credit consequences: Pursuant to Civil Code § 1785.26, you are\n"
          "hereby notified that a negative credit report reflecting on your credit record may be submitted in the\n"
          "future to a credit reporting agency if you fail to fulfill the terms of your rental obligation or if you\n"
          "default in those obligations in any way. This is the only notice that you will receive in this regard.",
          167, leading=13.75, bold_prefix="(A) NOTICE regarding possible credit consequences:")
    block('(B) NOTICE regarding abandoned property: "State law permits former tenants to reclaim\n'
          'abandoned personal property left at the former address of the tenant, subject to certain conditions.\n'
          'You may or may not be able to reclaim property without incurring additional costs, depending on\n'
          'the cost of storing the property and the length of time before it is reclaimed. In general, these costs\n'
          'will be lower the sooner you contact your former landlord after being notified that property\n'
          'belonging to you was left behind after you moved out." California Civil Codes §1946.1; §1983-\n'
          '1989 et seq.', 236, leading=13.75, bold_prefix="(B) NOTICE regarding abandoned property:")
    block("Please be advised that any tenant under the above-stated agreement that remains in the unit after the\n"
          "termination date may result in the Landlords seeking to enforce the termination in Court, in an action\n"
          "for Unlawful Detainer at which time the occupant(s) may appear and present a defense.",
          346, leading=13.75)
    block("Finally, any Notice to Pay or Quit served prior to this one is hereby superseded and canceled. Any notice\n"
          "served in conjunction with or subsequent to this one is not designed to supersede or cancel this present\n"
          "notice. You are still required to oblige by the terms outlined in this Notice.",
          401, width=486, leading=13.75)
    p = block("Date: {{NOTICE_DATE}}\t(Original signed by agent/owner)", 484.5, x=78, width=385,
              size=12, leading=14, slot="NOTICE_DATE")
    p.paragraph_format.tab_stops.add_tab_stop(Pt(194))
    # A real line below the signature caption; printed name is a separate slot.
    block("", 499, x=192, width=271, size=1, leading=0.5, border=True)
    block("{{SIGNER_NAME_AND_TITLE}}\n{{SIGNER_CONTINUED}}", 506, x=192, width=271, size=12,
          leading=14, align=WD_ALIGN_PARAGRAPH.CENTER, slot="SIGNER_NAME_AND_TITLE")
    block("Housing Provider's Address:", 540.5, x=216, width=365, size=12, leading=14)
    block("{{PROVIDER_STREET}}", 556, x=219, width=375, size=12, leading=14, slot="PROVIDER_STREET")
    block("{{PROVIDER_CITY}}", 571, x=217, width=375, size=12, leading=14, slot="PROVIDER_CITY")
    block("Phone: {{PROVIDER_PHONE}}", 585.5, x=216, width=365, size=12, leading=14, slot="PROVIDER_PHONE")
    doc.core_properties.title = "Blueprint residential three day notice to pay or quit"
    doc.core_properties.subject = "Approved layout 2026-10-26; case fields only"
    doc.core_properties.comments = "Reference PDF SHA256 0e23f57004a2f9842acaab92a334d46c031a5d09e31f4f3900b91f531418825d"
    # Drawing IDs are unique across the two retained page artworks.
    for i, prop in enumerate(doc._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'), 1):
        prop.set("id", str(i))
    doc.save(TARGET)
    return TARGET


if __name__ == "__main__":
    print(build())
