"""Engine regression check (no test framework needed).

    .venv/bin/python tests/check_engine.py

Converts the synthetic fixtures and asserts on document structure, plus OCR
accuracy: the text recovered from the *scanned* PDF is compared with the text of
the *digital* original.
"""
import difflib
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import docx  # noqa: E402

from server import engine  # noqa: E402

FIX = ROOT / "tests" / "fixtures"
if not (FIX / "digital.pdf").exists():
    sys.exit("Run tests/make_fixtures.py first.")


def structure(path):
    d = docx.Document(str(path))
    paras = [p for p in d.paragraphs if p.text.strip()]
    return {
        "h1": [p.text.strip() for p in paras if p.style.name == "Heading 1"],
        "h2": [p.text.strip() for p in paras if p.style.name == "Heading 2"],
        "bullets": [p.text.strip() for p in paras if p.style.name == "List Bullet"],
        "tables": [[[c.text.strip() for c in r.cells] for r in t.rows] for t in d.tables],
        "text": "\n".join(p.text.strip() for p in paras),
    }


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    return cond


ok = True
with tempfile.TemporaryDirectory() as tmp:
    a = engine.analyze_pdf(FIX / "digital.pdf")
    b = engine.analyze_pdf(FIX / "scanned.pdf")
    print("scan detection")
    ok &= check(not a.scanned, "digital.pdf is not scanned")
    ok &= check(b.scanned_pages == [0, 1], "scanned.pdf: both pages scanned")

    dig, scan_plain, scan_ocr = (Path(tmp) / n for n in ("d.docx", "s.docx", "o.docx"))
    engine.convert(FIX / "digital.pdf", dig)
    engine.convert(FIX / "scanned.pdf", scan_plain, ocr=False)
    stats = engine.convert(FIX / "scanned.pdf", scan_ocr, ocr=True)

    expected_h2 = ["Key figures", "Station summary", "Methods", "Next steps"]
    expected_table = [["Station", "pH", "Turbidity (NTU)", "Status"], ["North 1", "7.1", "18.4", "Review"],
                      ["North 2", "7.3", "15.9", "Review"], ["Central", "7.6", "4.2", "Good"],
                      ["East", "7.4", "3.8", "Good"]]
    for label, path in (("digital", dig), ("scanned + OCR", scan_ocr)):
        s = structure(path)
        print(label)
        ok &= check(s["h1"] == ["Quarterly Field Report"], "title is Heading 1")
        ok &= check(s["h2"] == expected_h2, "section headings are Heading 2")
        ok &= check(len(s["bullets"]) == 3, "3 bulleted list items")
        ok &= check(s["tables"] == [expected_table], "station table rebuilt cell-for-cell")
        ok &= check("rainfall in April." in s["text"], "wrapped lines reflowed into one paragraph")

    print("scanned without OCR")
    plain = docx.Document(str(scan_plain))
    ok &= check(len(plain.inline_shapes) + len(plain.element.body.findall(".//{*}drawing")) > 0
                and not any(p.text.strip() for p in plain.paragraphs), "pages become images, no text")

    print("OCR accuracy")
    ratio = difflib.SequenceMatcher(None, structure(dig)["text"], structure(scan_ocr)["text"]).ratio()
    ok &= check(ratio >= 0.97, f"OCR text matches the digital original ({ratio:.1%})")
    print(f"  info  OCR stats: {stats}")

print("\nALL CHECKS PASSED" if ok else "\nSOME CHECKS FAILED")
sys.exit(0 if ok else 1)
