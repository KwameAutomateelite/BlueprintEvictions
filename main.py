import json
import logging
import os
import tempfile
from datetime import datetime, timezone
from typing import Optional

import httpx
from dropbox_sign import ApiClient, ApiException, Configuration, apis, models
from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib import colors

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Blueprint Evictions - Dropbox Sign Service")

# --- Config ---

DROPBOX_SIGN_API_KEY = os.environ["DROPBOX_SIGN_API_KEY"]
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ.get("AIRTABLE_BASE_ID", "appumajkmYLcMryFd")
AIRTABLE_TABLE_ID = os.environ.get("AIRTABLE_TABLE_ID", "tblZbUGz8OTvFNh9i")

configuration = Configuration(username=DROPBOX_SIGN_API_KEY)


# --- Models ---


class SendSignatureRequest(BaseModel):
    signer_email: str
    signer_name: str
    document_name: str
    record_id: str
    notice_type: str
    case_name: str
    fields: dict  # Pre-filled notice fields (TENANT_NAMES, PROPERTY_ADDRESS, etc.)
    file_url: Optional[str] = None  # Optional: if a PDF already exists


class SendSignatureResponse(BaseModel):
    signature_request_id: str
    status: str
    message: str


# --- Helpers ---


def generate_notice_pdf(
    notice_type: str,
    case_name: str,
    signer_name: str,
    fields: dict,
) -> str:
    """Generate a PDF from notice fields using reportlab and return the temp file path."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.close()

    doc = SimpleDocTemplate(
        tmp.name,
        pagesize=letter,
        leftMargin=1 * inch,
        rightMargin=1 * inch,
        topMargin=1 * inch,
        bottomMargin=1 * inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "NoticeTitle", parent=styles["Title"], fontSize=18, spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        "NoticeSubtitle",
        parent=styles["Normal"],
        fontSize=14,
        textColor=colors.HexColor("#555555"),
        alignment=1,
        spaceAfter=12,
    )
    meta_style = ParagraphStyle(
        "NoticeMeta",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.HexColor("#666666"),
        spaceAfter=20,
    )
    label_style = ParagraphStyle(
        "FieldLabel", parent=styles["Normal"], fontSize=11, fontName="Helvetica-Bold"
    )
    value_style = ParagraphStyle(
        "FieldValue", parent=styles["Normal"], fontSize=11
    )
    sig_style = ParagraphStyle(
        "SigLabel",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.HexColor("#666666"),
    )

    elements = []

    # Header
    elements.append(Paragraph(notice_type, title_style))
    elements.append(Paragraph(case_name, subtitle_style))

    # Divider line via a thin table
    divider = Table([[""]],  colWidths=[6.5 * inch])
    divider.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 2, colors.HexColor("#333333")),
    ]))
    elements.append(divider)
    elements.append(Spacer(1, 12))

    # Meta
    generated = datetime.now(timezone.utc).strftime("%B %d, %Y")
    elements.append(
        Paragraph(f"Generated: {generated} | Prepared for: {signer_name}", meta_style)
    )

    # Fields table
    table_data = []
    for key, value in fields.items():
        label = key.replace("_", " ").title()
        val = str(value) if value else ""
        table_data.append([
            Paragraph(label, label_style),
            Paragraph(val, value_style),
        ])

    if table_data:
        fields_table = Table(table_data, colWidths=[2.3 * inch, 4.2 * inch])
        fields_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
        ]))
        elements.append(fields_table)

    # Signature block
    elements.append(Spacer(1, 60))

    sig_line = Table([[""]], colWidths=[4 * inch])
    sig_line.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, -1), 1, colors.HexColor("#333333")),
    ]))
    elements.append(sig_line)
    elements.append(Paragraph(f"Signature — {signer_name}", sig_style))

    elements.append(Spacer(1, 30))
    elements.append(sig_line)
    elements.append(Paragraph("Date", sig_style))

    doc.build(elements)
    return tmp.name


async def download_file(url: str) -> str:
    """Download a file from URL to a temp path and return the path."""
    async with httpx.AsyncClient(follow_redirects=True, timeout=60) as client:
        resp = await client.get(url)
        resp.raise_for_status()

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.write(resp.content)
    tmp.close()
    return tmp.name


async def update_airtable_status(record_id: str, status: str) -> None:
    """Update a record's Status field in Airtable."""
    url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_ID}/{record_id}"
    headers = {
        "Authorization": f"Bearer {AIRTABLE_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {"fields": {"Status": status}}

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.patch(url, headers=headers, json=payload)
        if resp.status_code != 200:
            logger.error(
                f"Airtable update failed for {record_id}: {resp.status_code} {resp.text}"
            )
            raise HTTPException(
                status_code=502, detail=f"Airtable update failed: {resp.text}"
            )
        logger.info(f"Airtable record {record_id} updated to Status='{status}'")


# --- Endpoints ---


@app.get("/health")
async def health():
    return {"status": "ok", "service": "dropbox-sign-service"}


@app.post("/send-signature", response_model=SendSignatureResponse)
async def send_signature(req: SendSignatureRequest):
    """Generate a PDF from notice fields (or download from URL) and send to Dropbox Sign."""
    logger.info(
        f"Sending signature request: {req.document_name} to {req.signer_email}"
    )

    # Get PDF: either download from URL or generate from fields
    try:
        if req.file_url:
            file_path = await download_file(req.file_url)
        else:
            file_path = generate_notice_pdf(
                notice_type=req.notice_type,
                case_name=req.case_name,
                signer_name=req.signer_name,
                fields=req.fields,
            )
    except Exception as e:
        logger.error(f"Failed to prepare PDF: {e}")
        raise HTTPException(status_code=400, detail=f"Failed to prepare PDF: {str(e)}")

    try:
        with ApiClient(configuration) as api_client:
            signature_request_api = apis.SignatureRequestApi(api_client)

            signer = models.SubSignatureRequestSigner(
                email_address=req.signer_email,
                name=req.signer_name,
                order=0,
            )

            signing_options = models.SubSigningOptions(
                draw=True,
                type=True,
                upload=True,
                phone=False,
                default_type="draw",
            )

            data = models.SignatureRequestSendRequest(
                title=req.document_name,
                subject=f"Please sign: {req.document_name}",
                message=(
                    f"Hi {req.signer_name},\n\n"
                    f"Please review and sign the attached document: {req.document_name}.\n\n"
                    "Thank you,\nBlueprint Evictions LLC"
                ),
                signers=[signer],
                files=[open(file_path, "rb")],
                metadata={"record_id": req.record_id},
                signing_options=signing_options,
                test_mode=os.environ.get("DROPBOX_SIGN_TEST_MODE", "0") == "1",
            )

            result = signature_request_api.signature_request_send(data)
            sig_request = result.signature_request

            logger.info(
                f"Signature request created: {sig_request.signature_request_id}"
            )

            return SendSignatureResponse(
                signature_request_id=sig_request.signature_request_id,
                status="sent",
                message=f"Signature request sent to {req.signer_email}",
            )

    except ApiException as e:
        logger.error(f"Dropbox Sign API error: {e.status} {e.body}")
        raise HTTPException(
            status_code=502,
            detail=f"Dropbox Sign API error: {e.body}",
        )
    finally:
        # Clean up temp file
        try:
            os.unlink(file_path)
        except OSError:
            pass


@app.post("/signature-callback")
async def signature_callback(request: Request):
    """Webhook endpoint that Dropbox Sign calls when a document is signed.

    Dropbox Sign sends event data as form-encoded with a 'json' field.
    On event 'signature_request_all_signed', update Airtable Status to 'Signed (A)'.
    """
    form = await request.form()

    # Dropbox Sign sends the payload in a 'json' form field
    json_str = form.get("json")
    if not json_str:
        logger.warning("Callback received with no 'json' field")
        return {"status": "ignored", "reason": "no json payload"}

    payload = json.loads(json_str)
    event = payload.get("event", {})
    event_type = event.get("event_type")
    event_time = event.get("event_time")

    logger.info(f"Dropbox Sign callback: event_type={event_type} time={event_time}")

    # Respond to the callback test (Dropbox Sign sends this to verify the endpoint)
    if event_type == "callback_test":
        logger.info("Responding to callback_test with 'Hello API Event Received'")
        return "Hello API Event Received"

    # Only act on all-signed events
    if event_type != "signature_request_all_signed":
        logger.info(f"Ignoring event type: {event_type}")
        return {"status": "ignored", "event_type": event_type}

    # Extract metadata with record_id
    signature_request = payload.get("signature_request", {})
    metadata = signature_request.get("metadata", {})
    record_id = metadata.get("record_id")
    sig_request_id = signature_request.get("signature_request_id", "unknown")

    if not record_id:
        logger.error(
            f"No record_id in metadata for signature request {sig_request_id}"
        )
        return {"status": "error", "reason": "no record_id in metadata"}

    logger.info(
        f"All signed for request {sig_request_id}, updating record {record_id}"
    )

    # Update Airtable
    await update_airtable_status(record_id, "Signed (A)")

    return {
        "status": "processed",
        "event_type": event_type,
        "record_id": record_id,
        "signature_request_id": sig_request_id,
    }
