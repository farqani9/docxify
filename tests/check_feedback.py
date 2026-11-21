"""Structural checks over every feedback case (feedback/<case>/*.pdf).

    .venv/bin/python tests/check_feedback.py

Converts each PDF and verifies the properties fixed in response to feedback,
without relying on document content (the files stay private and out of git).
"""
import re
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from server import engine  # noqa: E402

UNMAPPED = re.compile(r'w:ascii="(Liberation[^"]*|Free(?:Sans|Serif|Mono)[^"]*|Nimbus[^"]*|TeX Gyre[^"]*|[A-Z]{6}\+[^"]*)"')


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    return cond


pdfs = sorted(p for p in (ROOT / "feedback").glob("*/*.pdf") if not p.name.startswith("_"))
if not pdfs:
    sys.exit("No feedback PDFs found.")
ok = True
with tempfile.TemporaryDirectory() as tmp:
    for pdf in pdfs:
        print(f"{pdf.parent.name}/{pdf.name}")
        out = Path(tmp) / (pdf.stem + ".docx")
        stats = engine.convert(pdf, out)
        with zipfile.ZipFile(out) as z:
            doc = z.read("word/document.xml").decode("utf-8")
            names = z.namelist()
        ok &= check(not UNMAPPED.search(doc), "fonts mapped to installed equivalents")
        ok &= check(doc.count("<w:sectPr") <= 1 or stats.get("sections_merged", 0) == 0,
                    f"continuous flow ({doc.count('<w:sectPr')} section)")
        ok &= check('w:hRule="exact"' not in doc and 'w:lineRule="exact"' not in doc,
                    "no fixed row heights / exact line spacing (no clipped text)")
        ok &= check(not re.search(r'w:\w+="-?\d+\.\d+"', doc) and not re.search(r'w:(?:color|fill)="#', doc),
                    "valid OOXML measurements and colours")
        if stats.get("header_footer"):
            hdr = any(n.startswith("word/header") for n in names)
            ftr = [n for n in names if n.startswith("word/footer")]
            footer_xml = "".join(zipfile.ZipFile(out).read(n).decode() for n in ftr)
            ok &= check(bool(hdr and ftr), "running header/footer rebuilt as Word header/footer")
            if re.search(r"\bPage\b", footer_xml):
                ok &= check("PAGE" in footer_xml, "page number is a live PAGE field")
        packed = len(re.findall(r"<w:t>[^<]*\S</w:t></w:r><w:r>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:t>\u2022", doc))
        ok &= check(packed == 0, "no bullet items packed into one paragraph")
        print(f"  info {stats}")
print("\nALL FEEDBACK CHECKS PASSED" if ok else "\nSOME FEEDBACK CHECKS FAILED")
sys.exit(0 if ok else 1)
