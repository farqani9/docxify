"""Engine regression check (no test framework needed).

    .venv/bin/python tests/check_engine.py

Converts the synthetic fixtures and asserts on document structure, plus OCR
accuracy: the text recovered from the *scanned* PDF is compared with the text of
the *digital* original. The Arabic fixtures are checked against their source text,
including right-to-left paragraph marks and the digits/Latin inside Arabic lines.
"""
import re
import difflib
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import make_fixtures as F  # noqa: E402
from server import engine  # noqa: E402

FIX = ROOT / "tests" / "fixtures"
if not all((FIX / n).exists() for n in ("digital.pdf", "arabic_digital.pdf", "arabic_shaped.pdf",
                                        "arabic_scanned.pdf", "arabic_pages.pdf", "generic.pdf")):
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

    # ---------------------------------------------------------------- Arabic
    def norm(t):
        return re.sub(r"\s+", " ", t).strip()

    def rtl_marked(path):
        d = docx.Document(str(path))
        paras = [p for p in d.paragraphs if re.search("[\u0600-\u06FF]", p.text)]
        paras += [p for t in d.tables for r in t.rows for c in r.cells for p in c.paragraphs
                  if re.search("[\u0600-\u06FF]", p.text)]
        bidi = all(p._p.pPr is not None and p._p.pPr.find(qn("w:bidi")) is not None for p in paras)
        runs = all(r._r.rPr is not None and r._r.rPr.find(qn("w:rtl")) is not None
                   for p in paras for r in p.runs if re.search("[\u0600-\u06FF]", r.text))
        sect = d.element.body.find(qn("w:sectPr"))
        return bool(paras) and bidi and runs and sect is not None and sect.find(qn("w:bidi")) is not None

    ar_dig, ar_ocr = Path(tmp) / "ad.docx", Path(tmp) / "ao.docx"
    engine.convert(FIX / "arabic_digital.pdf", ar_dig)
    ar_stats = engine.convert(FIX / "arabic_scanned.pdf", ar_ocr, ocr=True, lang="ar")

    print("Arabic digital (right-to-left text layer)")
    s = structure(ar_dig)
    truth = norm(" ".join([F.AR_TITLE, F.AR_SUBTITLE, F.AR_BODY, F.AR_HEADINGS[0], *F.AR_BULLETS, F.AR_MIXED]))
    ratio = difflib.SequenceMatcher(None, truth, norm(" ".join(s["text"].split("\n"))), autojunk=False).ratio()
    ok &= check(ratio >= 0.99, f"text in reading order ({ratio:.1%})")
    ok &= check(F.AR_MIXED in s["text"], "numbers and Latin inside Arabic lines keep their place")
    ok &= check(s["h1"] == [F.AR_TITLE] and s["h2"] == [F.AR_HEADINGS[0]], "title and heading styles")
    ok &= check([b.strip() for b in s["bullets"]] == F.AR_BULLETS, "3 bulleted list items, exact text")
    ok &= check(rtl_marked(ar_dig), "paragraphs, runs and section are right-to-left")

    print("Arabic shaped (glyph-level text layer, bordered table)")
    ar_shp = Path(tmp) / "as.docx"
    engine.convert(FIX / "arabic_shaped.pdf", ar_shp)
    s = structure(ar_shp)
    table = s["tables"][0] if s["tables"] else []
    ok &= check(s["h1"] == [F.AR_TITLE] and len(s["h2"]) == 2, "title is Heading 1 (not merged with the subtitle), two Heading 2")
    ok &= check(len(s["bullets"]) == 3, "3 bulleted list items")
    ok &= check(len(table) == 4 and all(len(r) == 3 for r in table) and [r[1] for r in table[1:]] == ["7.1", "7.3", "7.6"],
                "table rebuilt cell by cell (4 x 3)")
    ok &= check(bool(table) and table[0][1].endswith("(pH)"), "brackets in reading order: \"(pH)\"")
    ok &= check(any(p.endswith("2026") and "pH" in p for p in s["text"].split("\n")), "numbers and Latin in place")
    ok &= check(rtl_marked(ar_shp), "paragraphs, runs and section are right-to-left")

    print("Arabic scanned + OCR")
    s = structure(ar_ocr)
    got = norm(" ".join(s["text"].split("\n") + [c for t in s["tables"] for r in t for c in r]))
    # paragraphs first, then table cells (the order `got` is built in)
    truth = norm(" ".join([F.AR_TITLE, F.AR_SUBTITLE, F.AR_BODY, F.AR_HEADINGS[0], *F.AR_BULLETS,
                           F.AR_HEADINGS[1], F.AR_MIXED, *[c for r in F.AR_TABLE for c in r]]))
    ratio = difflib.SequenceMatcher(None, truth, got, autojunk=False).ratio()
    ok &= check(ratio >= 0.93, f"OCR text matches the source ({ratio:.1%})")
    numbers = ["2026", "14", "8.2", "7.1", "7.3", "7.6", "pH"]
    ok &= check(all(n in got for n in numbers), "digits and Latin recovered inside Arabic lines")
    ok &= check(s["h1"] == [F.AR_TITLE] and len(s["h2"]) == 2, "title is Heading 1, two Heading 2")
    ok &= check(len(s["bullets"]) == 3, "3 bulleted list items")
    table = s["tables"][0] if s["tables"] else []
    ok &= check(len(table) == 4 and all(len(r) == 3 for r in table)
                and [r[1] for r in table[1:]] == ["7.1", "7.3", "7.6"], "table rebuilt (4 x 3, values in place)")
    ok &= check(rtl_marked(ar_ocr), "paragraphs, runs and section are right-to-left")
    print(f"  info  Arabic OCR stats: {ar_stats}")

    # ------------------------------------------- multi-page Arabic: header/footer, lists
    def band(path):
        d = docx.Document(str(path))
        sec = d.sections[0]
        header = " ".join(p.text.strip() for p in sec.header.paragraphs if p.text.strip())
        fields = [f.get(qn("w:instr")).strip() for f in sec.footer._element.iter(qn("w:fldSimple"))]
        paras = [p for p in d.paragraphs if p.text.strip()]
        numbered = [p for p in paras if p.style.name == "List Number"]
        starts = []
        for p in numbered:
            num = p._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
            starts.append(num.get(qn("w:val")) if num is not None else "style")
        text = "\n".join(p.text for p in paras)
        return header, fields, numbered, starts, text

    for label, name, ocr, min_items in (("Arabic pages (digital)", "arabic_pages.pdf", False, 15),
                                        ("Arabic pages scanned + OCR", "arabic_pages_scanned.pdf", True, 9)):
        print(label)
        out = Path(tmp) / (name + ".docx")
        engine.convert(FIX / name, out, ocr=ocr, lang="ar")
        header, fields, numbered, starts, text = band(out)
        ok &= check(header == F.AR_PAGES_HEADER, f"running header rebuilt in reading order ({header!r})")
        ok &= check(fields[:2] == ["PAGE", "NUMPAGES"], "footer \"صفحة N من M\" keeps live page fields")
        ok &= check(len(numbered) >= min_items, f"numbered items ({len(numbered)} of 15)")
        ok &= check(len(set(starts)) >= 3, "each list counts on its own (\"1.\" and \"١.\" lists restart)")
        ok &= check(not re.search(r"QXR\d|Z{3,}", text + header), "no placeholder text left")
        ok &= check(not re.search(r"(?m)^\.\d", text), "list numbers in reading order (no \".2\")")
        ok &= check(bool(re.search(r"(?m)^Contact: info@example\.com\s*$", text)), "English line is its own paragraph")

    # ------------------------------------------- clean-up rules on an English page
    print("clean-up rules (English)")
    gen = Path(tmp) / "generic.docx"
    engine.convert(FIX / "generic.pdf", gen)
    d = docx.Document(str(gen))
    heads = [p.text.strip() for p in d.paragraphs if p.style.name.startswith("Heading")]
    ok &= check("Ministry of Health" in heads and len(heads) >= 2, "two same-size headings stay separate")
    nums = [p for p in d.paragraphs if p.style.name == "List Number"]
    overrides = [int(s.get(qn("w:val"))) for s in d.part.numbering_part.element.iter(qn("w:startOverride"))]
    ok &= check(len(nums) == 4 and 3 in overrides, "a list continuing after a note keeps counting from 3")
    from docx import Document as _Doc
    from server import docx_fix
    probe = _Doc()
    probe.add_table(rows=1, cols=3)          # plain one-row table (no fill) ...
    probe.add_paragraph("")
    probe.add_table(rows=3, cols=3)          # ... above another table with as many columns
    ok &= check(docx_fix.merge_header_rows(probe) == 0, "a plain one-row table is not taken for a header row")

print("\nALL CHECKS PASSED" if ok else "\nSOME CHECKS FAILED")
sys.exit(0 if ok else 1)
