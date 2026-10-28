# Residential three-day notice: use the original PDF

`POST /generate-notice-pdf` accepts the same case JSON as `/generate-notice` and
returns an editable, two-page PDF filled from Anne-Michelle's actual supplied
PDF. The existing DOCX endpoint remains available to its current consumers.
Do not convert this PDF to Word and back; that reintroduces layout changes.

There is no page-layout reconstruction. For every field:

```
page, x_left, y_bottom, x_right, y_top = original_widget.Rect
width = x_right - x_left
text_x = x_left + 2                         # original left inset, points
centered_text_x = x_left + (width - text_width) / 2
text_y = y_bottom + original_appearance_baseline
text_width = sum(original_font_glyph_widths) * 12 / 1000
```

PDF coordinates start at the bottom left; 72 points equal one inch. The original
612 x 792 point page content streams are preserved byte-for-byte. The total stays
in its page 1 rectangle `[352.08, 363.538, 459.651, 378.11]`, regardless of rent-row
count. The signature wording and line are already printed in the original page 2
content stream; the code never redraws, repositions or types that wording.
The printed signer stays centered inside its original page 2 field.

Only case values change: tenant/premises details, county, up to six unpaid rent
periods and amounts, the checked total, payment recipient/contact, notice date,
and designated signer/name/title. Blank rent rows stay blank. The total must
equal the sum in decimal cents. Missing required inputs and overflow require
review; no shrinking, page merging, extra page or guessed signer is permitted.
The payment recipient and the designated signer remain separate inputs.

The reusable master is the supplied PDF with case values and old text
appearances blanked; static page contents and field rectangles are unchanged.
The source hash is checked by `tools/blank_residential_pdf.py`, and the resulting
blank-master hash is checked before every render. Its field tree and
appearances are retained; linked payment/provider contact widgets are separated
only if the case explicitly supplies different contacts. Updated canonical
values and widget appearances are checked in tests. A final clean clone drops
orphaned source appearances so another case does not carry old sample values.

Local use (no network, email, signing or case changes):

```
python tools/render_residential_pdf.py case.json notice.pdf
```

`POST /render-saved-notice-pdf` accepts `docx_base64` and `expected_sha256`.
It reads the named fields from the actual saved Word editing copy, checks the
unchanged boilerplate and rent columns, and returns the exact PDF, its hash and
the Word source hash. Other notice families return `is_standard: false` and keep
their existing conversion. A malformed residential form fails instead of falling
back to a layout that can move.

The approval copy is the exact PDF. Word remains the editing source. After an
edit, make a fresh PDF and obtain approval of that version. Before approval is
frozen, compare the saved review PDF's bytes and a fresh render of the current
Word against the reviewed PDF hash. A mismatch requires a fresh review; never
silently substitute the edited Word conversion for the PDF she approved.
