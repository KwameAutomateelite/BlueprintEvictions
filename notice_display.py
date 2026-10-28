"""Display rules confirmed in Anne-Michelle's October 28 working session.

Normalize presentation only. Preserve the intake record and all money values;
an informal title never supplies a missing signer or grants authority.
"""
import calendar
from datetime import datetime
import re


def compact(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def display_case(value):
    value = compact(value)
    if value.isupper() or value.islower():
        value = value.title()
    # Preserve organizational abbreviations and the standard postal label.
    value = re.sub(r"\b(llc|llp|lp|inc|ltd)\b", lambda m: m[0].upper(), value, flags=re.I)
    value = re.sub(r"\bp\.?\s*o\.?\s+box\b", "P.O. Box", value, flags=re.I)
    return value


def tenant_display(value):
    # A line break or semicolon is a known separator; never guess where two
    # people divide inside an otherwise unpunctuated name.
    parts = re.split(r"\r?\n|;", str(value or ""))
    value = ", ".join(compact(p).strip(" ,") for p in parts if compact(p))
    value = re.sub(r",\s*,", ",", value)
    value = re.sub(r"\ball other occupants\b", "All Other Occupants", value, flags=re.I)
    # Normalize each name independently, including an ALL CAPS name preceding
    # the mixed-case occupants phrase.
    parts = [display_case(p) for p in value.split(",")]
    value = ", ".join(p for p in parts if p)
    return re.sub(r"\b(?:and|And) All Other Occupants\b", "and All Other Occupants", value)


def signer_title(value):
    value = compact(value)
    if re.search(r"\b(?:child|children|son|daughter|wife|husband|spouse|niece|nephew|grandparent|grandmother|grandfather|grandchild|grandson|granddaughter|sibling|sister|brother|mother|father|parent|aunt|uncle|cousin|relative)\b", value, re.I):
        return "Authorized Agent for Property Owner"
    if value.casefold() in {"owner", "property owner"}:
        return "Property Owner"
    return display_case(value)


def rent_period(value):
    value = compact(value).lstrip("; ")
    # Explicit ranges/partial periods stay explicit. A single dated monthly
    # entry is presented as that calendar month, independently of rent-due day.
    date = None
    for fmt in ("%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y",
                "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%B %Y", "%b %Y", "%Y-%m"):
        try:
            date = datetime.strptime(value, fmt)
            break
        except ValueError:
            pass
    if date:
        return f"{date:%B} 1-{calendar.monthrange(date.year, date.month)[1]}, {date.year}"
    return value


def city_line(value):
    value = compact(value)
    m = re.fullmatch(r"(.+?)(?:\s*,\s*|\s+)([A-Za-z]{2})\s*,?\s*(\d{5}(?:-\d{4})?)", value)
    if not m:
        return ""
    return f"{display_case(m[1].strip(' ,'))}, {m[2].upper()} {m[3]}"


def split_full_address(value):
    """Separate literal street and city/state/ZIP without a ZIP lookup/guess."""
    value = compact(value)
    m = re.fullmatch(r"(.+?)(?:\s*,\s*|\s+)([A-Za-z]{2})\s*,?\s*(\d{5}(?:-\d{4})?)", value)
    if not m:
        return "", ""
    before, state, zipcode = m.groups()
    parts = [p.strip() for p in before.strip(" ,").split(",") if p.strip()]
    if len(parts) >= 2:
        street, city = ", ".join(parts[:-1]), parts[-1]
    else:
        # The call's PO Box parsing failure: PO BOX 12345 SAN JOSE, CA 95152.
        box = re.fullmatch(r"(P\.?\s*O\.?\s+BOX\s+[A-Za-z0-9-]+)\s+(.+)", before.strip(" ,"), re.I)
        if not box:
            return "", ""
        street, city = box.groups()
    return display_case(street), f"{display_case(city)}, {state.upper()} {zipcode}"


def address_display(street, city, full):
    street, city, full = map(compact, (street, city, full))
    complete = city_line(city)
    if street and complete:
        suffix = re.escape(complete).replace(r"\ ", r"\s+")
        street = re.sub(r",?\s*" + suffix + r"$", "", street, flags=re.I).strip(" ,")
        return display_case(street), complete
    parsed = split_full_address(full or ", ".join(filter(None, (street, city))))
    if all(parsed):
        return parsed
    return display_case(street), complete


def display_inputs(data):
    data = dict(data)
    data["tenant_names"] = tenant_display(data.get("tenant_names"))
    for field in ("landlord_name", "payment_recipient_name", "signer_name", "authorized_signer_name"):
        if data.get(field):
            data[field] = display_case(data[field])
    for field in ("signer_title", "authorized_signer_title"):
        if data.get(field):
            data[field] = signer_title(data[field])
    data["amounts_due"] = [{**r, "due_date": rent_period(r.get("due_date"))} for r in data.get("amounts_due", [])]
    return data
