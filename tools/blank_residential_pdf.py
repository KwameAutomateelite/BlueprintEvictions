"""Build a reusable blank master from the supplied PDF, without case data."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import sys
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream, NameObject, TextStringObject

APPROVED_SOURCE_SHA256 = "0e23f57004a2f9842acaab92a334d46c031a5d09e31f4f3900b91f531418825d"


def blank_master(source):
    if sha256(source).hexdigest() != APPROVED_SOURCE_SHA256:
        raise ValueError("This is not Anne-Michelle's supplied reference PDF")
    writer = PdfWriter()
    writer.clone_document_from_reader(PdfReader(BytesIO(source)))
    for page in writer.pages:
        for ref in page["/Annots"]:
            widget = ref.get_object()
            field = widget.get("/Parent", widget).get_object()
            if field.get("/FT") != "/Tx":
                continue
            field[NameObject("/V")] = TextStringObject("")
            field[NameObject("/DV")] = TextStringObject("")
            field.pop(NameObject("/RV"), None)
            normal = widget.get("/AP", {}).get("/N")
            if normal:
                appearance = normal.get_object()
                stream = ContentStream(appearance, writer)
                for operands, operator in stream.operations:
                    if operator in (b"Tj", b"'", b'"'):
                        operands[-1] = TextStringObject("")
                    elif operator == b"TJ":
                        operands[0].clear()
                appearance.set_data(stream.get_data())
    intermediate = BytesIO()
    writer.write(intermediate)
    clean = PdfWriter()
    clean.clone_document_from_reader(PdfReader(BytesIO(intermediate.getvalue())))
    out = BytesIO()
    clean.write(out)
    return out.getvalue()


if __name__ == "__main__":
    source, output = map(Path, sys.argv[1:])
    output.write_bytes(blank_master(source.read_bytes()))
    print(sha256(output.read_bytes()).hexdigest())
