"""Local original-form filling: python tools/render_residential_pdf.py case.json notice.pdf."""
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parents[1]))
from residential_pdf import render_residential_pdf

if __name__ == "__main__":
    source, output = map(Path, sys.argv[1:])
    output.write_bytes(render_residential_pdf(json.loads(source.read_text())))
