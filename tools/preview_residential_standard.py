"""Generate fictional local review copies only; no provider requests."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
from test_residential_standard import case
from residential_notice import render_residential_notice

target = Path(sys.argv[1])
target.mkdir(parents=True, exist_ok=True)
for n in (1, 3, 6):
    data = case(n)
    if n == 6:
        data["tenant_names"] = "Alexandria Example Hernandez, Benjamin Example Hernandez, All Other Occupants"
    p = target / f"residential-standard-{n}-rent-rows.docx"
    p.write_bytes(render_residential_notice(data))
    print(p)
