"""Case-field filling for the approved two-page residential notice master.

No paragraph insertion, layout rewrite, LLM, network call, or signing occurs here.
"""
from datetime import datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
import re
import unicodedata
from zoneinfo import ZoneInfo

from docx import Document

TEMPLATE_VERSION = "residential-pay-or-quit-2026-10-26"
REFERENCE_SHA256 = "0e23f57004a2f9842acaab92a334d46c031a5d09e31f4f3900b91f531418825d"
MASTER = Path(__file__).parent / "templates/residential-pay-or-quit-2026-10.docx"
MAX_RENT_ROWS = 6


class NoticeInputError(ValueError):
    """The case needs correction/review; do not produce a misleading notice."""


def clean(value):
    s = re.sub(r"\s+", " ", str(value or "")).strip()
    return "" if s.upper() in {"NOT_FOUND", "NOT FOUND", "N/A", "NONE", "NULL"} or re.fullmatch(r"_+", s) else s


def money(value):
    raw = clean(value)
    if not re.fullmatch(r"\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?", raw):
        raise NoticeInputError("Rent amounts must be numeric dollars and cents")
    try:
        return Decimal(raw.replace("$", "").replace(",", "")).quantize(Decimal("0.01"))
    except InvalidOperation as exc:
        raise NoticeInputError("Invalid rent amount") from exc


def text_width(text, size=12):
    # Conservative Times Roman width bounds; never shrink a case field to fit.
    # Narrow/wide groups prevent a long string of W/M from slipping through a
    # simple character-count check. Non-Latin full-width glyphs use a full em.
    total = 0
    for c in text:
        if unicodedata.east_asian_width(c) in {"W", "F"}:
            em = 1
        elif c in "MW@%":
            em = .95
        elif c in "ilI.,:;'!| ":
            em = .28
        elif c.isupper():
            em = .73
        elif c.isdigit():
            em = .5
        else:
            em = .56
        total += em * size
    return total


def fit(value, width, field):
    value = clean(value)
    if text_width(value) > width - 4:
        raise NoticeInputError(f"{field} exceeds the approved form space; human review required")
    if "{{" in value or "}}" in value:
        raise NoticeInputError(f"Invalid template text in {field}")
    return value


def two_lines(value, width, field, first_prefix=""):
    value = clean(value)
    if "{{" in value or "}}" in value:
        raise NoticeInputError(f"Invalid template text in {field}")
    words = value.split()
    lines = [""]
    for word in words:
        candidate = (lines[-1] + " " + word).strip()
        available = width - (text_width(first_prefix) if len(lines) == 1 else 0)
        if text_width(candidate) > available - 4:
            if not lines[-1] or len(lines) == 2:
                raise NoticeInputError(f"{field} exceeds the approved two-line space; human review required")
            lines.append(word)
            if text_width(word) > width - 4:
                raise NoticeInputError(f"{field} contains a word too wide for the form")
        else:
            lines[-1] = candidate
    return (lines + [""])[:2]


def notice_date(value):
    value = clean(value)
    if not value:
        d = datetime.now(ZoneInfo("America/Los_Angeles")).date()
    else:
        d = None
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%B %d, %Y", "%b %d, %Y"):
            try:
                d = datetime.strptime(value, fmt).date()
                break
            except ValueError:
                pass
        if d is None:
            raise NoticeInputError("Notice date must be a valid calendar date")
    return f"{d.month}/{d.day}/{d.year % 100:02d}"


def address_parts(data, prefix):
    if prefix == "property":
        street = clean(data.get("property_address_street"))
        city = clean(data.get("property_address_city"))
        full = clean(data.get("property_address"))
    else:
        # payment_address is the designated delivery address when supplied.
        street = clean(data.get("payment_address_street"))
        city = clean(data.get("payment_address_city"))
        full = clean(data.get("payment_address")) or clean(data.get("landlord_address"))
        if not city:
            c = clean(data.get("landlord_city"))
            s = clean(data.get("landlord_state"))
            z = clean(data.get("landlord_zip"))
            city = (c + (", " if c and s else "") + s.upper() + (" " if z and (c or s) else "") + z).strip()
    if not street:
        # Preserve unit components on the street line when splitting a full
        # address. Prefer explicit street/city inputs for unusual addresses.
        parts = [p.strip() for p in full.split(",") if p.strip()]
        if city:
            street = full
            suffix = re.escape(city).replace(r"\ ", r"\s+")
            street = re.sub(r",?\s*" + suffix + r"$", "", street, flags=re.I).strip(" ,")
        elif len(parts) >= 3:
            street = ", ".join(parts[:-2])
            city = ", ".join(parts[-2:])
        elif len(parts) == 2:
            street, city = parts
        else:
            street = full
    if prefix == "property":
        for key, label in (("apt", "Apt"), ("unit", "Unit")):
            unit = clean(data.get(key))
            if unit:
                labeled = unit if re.match(r"^(?:apt\b|apartment\b|unit\b|suite\b|ste\b|#)", unit, re.I) else f"{label} {unit}"
                # Upstream frequently supplies unit both in street and apt.
                if not re.search(r"(?:apt\.?|apartment|unit|suite|ste\.?|#)\s*" + re.escape(re.sub(r"^(?:apt\.?|apartment|unit|suite|ste\.?|#)\s*", "", labeled, flags=re.I)) + r"\b", street, re.I):
                    street += ", " + labeled
    return street, city


def fields_for_case(data):
    if data.get("notice_type") != "residential" or data.get("day_count") != 3:
        raise NoticeInputError("This master is only for residential three-day pay-or-quit notices")
    override = clean(data.get("notice_header"))
    if override and override.upper() not in {"THREE (3) DAY NOTICE TO PAY OR QUIT", "THREE (3) DAY NOTICE TO PAY RENT OR QUIT"}:
        raise NoticeInputError("The approved notice title cannot be replaced")
    tenants = two_lines(data.get("tenant_names"), 480, "Tenant names", first_prefix="To: ")
    if not tenants[0]:
        raise NoticeInputError("Tenant names are required")
    street, city = address_parts(data, "property")
    if not street or not city:
        raise NoticeInputError("Complete premises street and city/state/ZIP are required")
    county = re.sub(r"^COUNTY\s+OF\s+|\s+COUNTY$", "", clean(data.get("county") or data.get("county_name")), flags=re.I).strip()
    if not county:
        raise NoticeInputError("Premises county is required")
    rows = data.get("amounts_due") or []
    if not 1 <= len(rows) <= MAX_RENT_ROWS:
        raise NoticeInputError("The approved form has six rent rows; one to six rows required. Do not truncate or alter the form")
    f = {
        "TENANT_NAMES": tenants[0], "TENANT_NAMES_CONTINUED": tenants[1],
        "PROPERTY_STREET": fit(street, 445, "Premises street"),
        "PROPERTY_CITY": fit(city, 445, "Premises city/state/ZIP"),
        "COUNTY": fit(county.upper(), 375, "County"),
        "NOTICE_DATE": notice_date(data.get("notice_date")),
    }
    total = Decimal("0.00")
    for i in range(1, MAX_RENT_ROWS + 1):
        period = amount = ""
        if i <= len(rows):
            row = rows[i - 1]
            period = re.sub(r"^[;\s]+", "", clean(row.get("due_date")))
            if not period:
                raise NoticeInputError("Each rent row needs a period/date")
            period = fit(period, 176, "Rent period")
            value = money(row.get("amount"))
            if value <= 0:
                raise NoticeInputError("Each rent row must have a positive amount")
            total += value
            amount = fit(f"${value:,.2f}", 100, "Rent amount")
        f[f"PERIOD_{i}"] = period
        f[f"AMOUNT_{i}"] = amount
    if money(data.get("total_amount_due")) != total:
        raise NoticeInputError("Total rent does not equal the sum of the rent rows")
    f["TOTAL_AMOUNT"] = fit(f"${total:,.2f}", 108, "Total rent")
    payee = clean(data.get("payment_recipient_name")) or clean(data.get("landlord_name"))
    payee = re.sub(r"^(?:owner[’']?s?\s+name|landlord\s+name)\s*:\s*", "", payee, flags=re.I)
    if re.search(r"\b(?:company|owner[’']?s?\s+name)\s*:", payee, re.I):
        raise NoticeInputError("Payment recipient needs a verified display name; raw order-form labels require review")
    p_street, p_city = address_parts(data, "payment")
    phone = clean(data.get("payment_phone")) or clean(data.get("landlord_phone"))
    f.update({
        "PAYMENT_RECIPIENT": fit(payee, 420, "Payment recipient"),
        "PAYMENT_STREET": fit(p_street, 420, "Payment street"),
        "PAYMENT_CITY": fit(p_city, 420, "Payment city/state/ZIP"),
        "PAYMENT_PHONE": fit(phone, 415, "Payment phone"),
        "PROVIDER_STREET": fit(clean(data.get("housing_provider_street")) or p_street, 375, "Housing-provider street"),
        "PROVIDER_CITY": fit(clean(data.get("housing_provider_city")) or p_city, 375, "Housing-provider city/state/ZIP"),
        "PROVIDER_PHONE": fit(clean(data.get("housing_provider_phone")) or phone, 325, "Housing-provider phone"),
    })
    signer = clean(data.get("signer_name") or data.get("authorized_signer_name"))
    role = clean(data.get("signer_title") or data.get("authorized_signer_title"))
    # An owner/payee and their authorized signer may be different people.
    # Missing signatory information stays blank for review, never guessed.
    if role and not signer:
        raise NoticeInputError("Signer title requires a signer name")
    signer_lines = two_lines(signer + (", " + role if role else ""), 271, "Printed signer")
    f["SIGNER_NAME_AND_TITLE"], f["SIGNER_CONTINUED"] = signer_lines
    return f


def render_residential_notice(data, template_path=MASTER):
    fields = fields_for_case(data)
    doc = Document(template_path)
    seen = set()
    paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    for p in paragraphs:
        for run in p.runs:
            for key in re.findall(r"\{\{([A-Z0-9_]+)\}\}", run.text):
                if key not in fields:
                    raise RuntimeError(f"Unknown field in approved master: {key}")
                seen.add(key)
                run.text = run.text.replace("{{" + key + "}}", fields[key])
    if seen != set(fields):
        raise RuntimeError("Approved master field map is incomplete")
    doc.core_properties.identifier = TEMPLATE_VERSION
    out = BytesIO()
    doc.save(out)
    return out.getvalue()
