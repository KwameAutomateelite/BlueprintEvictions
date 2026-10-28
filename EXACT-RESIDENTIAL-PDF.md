# Residential three-day notice: use the original PDF

`POST /generate-notice-pdf` accepts the same case JSON as `/generate-notice` and
returns a PDF filled from Anne-Michelle's actual supplied PDF. A fitting schedule
keeps its original two-page layout; overflowing rent grows the table and moves
whole paragraphs onto continuation pages. The existing DOCX endpoint remains
available to its current consumers.
Do not convert this PDF to Word and back; that reintroduces layout changes.

For schedules that fit the original six boxes, every field uses:

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
count up to six fitting rows. The signature wording and line are already printed in the original page 2
content stream; the code never redraws, repositions or types that wording.
The printed signer stays centered inside its original page 2 field.

Only case values change: tenant/premises details, county, unpaid rent
periods and amounts, the checked total, payment recipient/contact, notice date,
and designated signer/name/title. Blank rent rows stay blank. The total must
equal the sum in decimal cents. Missing required inputs and overflow require
review; no guessed signer or silently truncated entries are permitted.
The payment recipient and the designated signer remain separate inputs.

## Schedules that outgrow the original space

`residential_pdf_flow.py` measures each period at the original 12pt font and wraps
it within the left column. Its amount stays in the right column. Row height is
`wrapped_line_count * 14.4 + 2.5` points. Before placing a row or paragraph:

```
if current_top + gap + block_height > 754:
    start a branded continuation page at top = 98
```

Rent column headings repeat when the table continues. The final rent row and
checked total stay together. Text that cannot fit in a single row can continue
over multiple lines/pages; its amount appears once. The original body paragraphs
move as complete vector fragments, retaining their original fonts, line breaks,
wording and horizontal positions. The signature, printed name and housing-provider
contact move as one block. Headers and footers remain outside the usable body.
Out-of-band text operators are removed as well as clipped, avoiding hidden
duplicate boilerplate. Canonical fields and their appearances remain present.

Word editing copies receive additional individually bound rent rows. Refresh
reads and reconciles all of them, including edits, before rebuilding the PDF.
The document editor also inspects every row instead of its former first six.
There is a 1,000-entry resource limit; a larger request is rejected explicitly,
never truncated. Existing short-form bytes and old reviewed hashes are unchanged.

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
