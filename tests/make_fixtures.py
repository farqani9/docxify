"""Generate synthetic test PDFs (no real documents involved).

- digital.pdf : two pages with a real text layer (headings, styled runs, list, table)
- scanned.pdf : the same pages rasterised at 150 dpi with a slight skew and
                noise, saved as image-only pages (no text layer) -> needs OCR
- arabic_digital.pdf : a right-to-left Arabic page with a text layer, mixed with
                numbers and Latin ("pH")
- arabic_shaped.pdf : a shaped right-to-left Arabic page (presentation-form glyphs,
                headings, list, bordered table with "(pH)"), as an HTML renderer
                writes it. Its text layer maps lam glyphs wrongly (a MuPDF quirk), so
                it is checked for structure, bracket order and numbers only.
- arabic_scanned.pdf : arabic_shaped.pdf scanned the same way -> needs OCR with
                lang="ar"
- arabic_pages.pdf : three Arabic pages with a running header, a "صفحة N من 3"
                footer, two numbered lists per page ("1." and "١.") and English lines;
  arabic_pages_scanned.pdf : the same, scanned
- generic.pdf : English page for clean-up rules: two same-size headings in a row and
                a numbered list that continues at "3." after a note
- light_table_scanned.pdf : a designed page as browsers print it (dark banner with
                white text, a table with very light borders, a coloured header band
                with white labels, zebra rows), scanned -> OCR must rebuild the table
"""
from pathlib import Path

import cv2
import numpy as np
import pymupdf

OUT = Path(__file__).parent / "fixtures"
OUT.mkdir(exist_ok=True)

TITLE = "Quarterly Field Report"
BODY_1 = (
    "This report summarises the results of the spring survey. Water quality "
    "improved at most sampling sites, while two northern stations still show "
    "elevated turbidity after heavy rainfall in April."
)
BODY_2 = (
    "The team recommends repeating the measurements in July and extending the "
    "network with three additional stations along the eastern channel."
)
BULLETS = [
    "Sampling sites visited: 14",
    "Average dissolved oxygen: 8.2 mg/L",
    "Stations needing follow-up: 2",
]
TABLE = [
    ["Station", "pH", "Turbidity (NTU)", "Status"],
    ["North 1", "7.1", "18.4", "Review"],
    ["North 2", "7.3", "15.9", "Review"],
    ["Central", "7.6", "4.2", "Good"],
    ["East", "7.4", "3.8", "Good"],
]
PAGE_2 = [
    ("Methods", 15, True),
    ("Samples were collected between 8:00 and 11:00 using sterile bottles and "
     "analysed within six hours. Each value in this report is the mean of three "
     "replicate measurements.", 11, False),
    ("Next steps", 15, True),
    ("Install the eastern stations, calibrate the turbidity sensors, and share "
     "the July results with the regional office before the end of August.", 11, False),
]


def write_wrapped(page, text, x, y, width, size, font="helv"):
    """Simple word wrap; returns the y after the block."""
    words, line = text.split(), ""
    lh = size * 1.35
    for w in words:
        trial = (line + " " + w).strip()
        if pymupdf.get_text_length(trial, fontname=font, fontsize=size) > width and line:
            page.insert_text((x, y), line, fontname=font, fontsize=size)
            y += lh
            line = w
        else:
            line = trial
    if line:
        page.insert_text((x, y), line, fontname=font, fontsize=size)
        y += lh
    return y


def make_digital(path):
    doc = pymupdf.open()
    left, width = 64, 467

    page = doc.new_page(width=595, height=842)  # A4
    y = 90
    page.insert_text((left, y), TITLE, fontname="hebo", fontsize=22)
    y += 18
    page.insert_text((left, y + 6), "Prepared by the Monitoring Team - May 2026",
                     fontname="heit", fontsize=10, color=(0.35, 0.35, 0.35))
    y += 40
    y = write_wrapped(page, BODY_1, left, y, width, 11)
    y += 8
    y = write_wrapped(page, BODY_2, left, y, width, 11)
    y += 14
    page.insert_text((left, y), "Key figures", fontname="hebo", fontsize=15)
    y += 22
    for b in BULLETS:
        page.insert_text((left + 6, y), "•", fontname="helv", fontsize=11)
        page.insert_text((left + 20, y), b, fontname="helv", fontsize=11)
        y += 16
    y += 18
    page.insert_text((left, y), "Station summary", fontname="hebo", fontsize=15)
    y += 14
    col_w = [120, 70, 140, 100]
    row_h = 22
    for r, row in enumerate(TABLE):
        x = left
        for c, cell in enumerate(row):
            rect = pymupdf.Rect(x, y, x + col_w[c], y + row_h)
            page.draw_rect(rect, color=(0.6, 0.6, 0.6), width=0.7,
                           fill=(0.93, 0.93, 0.93) if r == 0 else None)
            page.insert_text((x + 6, y + 15), cell,
                             fontname="hebo" if r == 0 else "helv", fontsize=10.5)
            x += col_w[c]
        y += row_h

    page = doc.new_page(width=595, height=842)
    y = 90
    for text, size, bold in PAGE_2:
        if bold:
            page.insert_text((left, y), text, fontname="hebo", fontsize=size)
            y += size + 10
        else:
            y = write_wrapped(page, text, left, y, width, size) + 12

    doc.set_metadata({"title": TITLE, "author": "Synthetic fixture"})
    doc.save(path)


def make_scanned(src, path, dpi=150, seed=7):
    rng = np.random.default_rng(seed)
    src_doc = pymupdf.open(src)
    out = pymupdf.open()
    for i, page in enumerate(src_doc):
        pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width).copy()
        # slight skew like a flatbed scan
        angle = 0.6 if i % 2 == 0 else -0.4
        h, w = img.shape
        m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
        img = cv2.warpAffine(img, m, (w, h), borderValue=255)
        # paper tone + sensor noise
        noise = rng.normal(0, 6, img.shape)
        img = np.clip(img.astype(np.float32) * 0.97 + 6 + noise, 0, 255).astype(np.uint8)
        ok, png = cv2.imencode(".png", img)
        p = out.new_page(width=page.rect.width, height=page.rect.height)
        p.insert_image(p.rect, stream=png.tobytes())
    out.save(path, deflate=True)


# ----------------------------------------------------------------- Arabic
AR_TITLE = "تقرير ميداني ربع سنوي"
AR_SUBTITLE = "إعداد فريق الرصد - مايو 2026"
AR_BODY = ("يلخص هذا التقرير نتائج المسح الربيعي. تحسنت جودة المياه في معظم مواقع أخذ العينات، "
           "بينما لا تزال محطتان في الشمال تسجلان عكارة مرتفعة بعد الأمطار الغزيرة في أبريل.")
AR_HEADINGS = ["أرقام رئيسية", "ملخص المحطات"]
AR_BULLETS = ["عدد المواقع التي تمت زيارتها: 14", "متوسط الأكسجين المذاب: 8.2 ملغم/لتر",
              "محطات تحتاج إلى متابعة: 2"]
AR_TABLE = [["المحطة", "الرقم الهيدروجيني (pH)", "الحالة"],
            ["الشمالية 1", "7.1", "مراجعة"],
            ["الشمالية 2", "7.3", "مراجعة"],
            ["الوسطى", "7.6", "جيدة"]]
AR_MIXED = "قيمة pH في المحطة الشمالية 7.1 خلال شهر مايو 2026"

AR_CSS = ("body { font-size: 12pt; } h1 { font-size: 22pt; margin: 0 0 4pt 0; } "
          "h2 { font-size: 15pt; margin: 14pt 0 6pt 0; } p { margin: 0 0 6pt 0; } "
          ".sub { font-size: 10pt; color: #555; } "
          "table { border-collapse: collapse; width: 100%; } "
          "td, th { border: 1px solid #777; padding: 4px 8px; } th { background: #eee; }")


def make_arabic_source(path):
    """Properly shaped right-to-left Arabic page (rendered by MuPDF's HTML engine
    with its built-in Noto Naskh Arabic). Used as the source of the scanned fixture."""
    rows = "".join("<tr>" + "".join(f"<{'th' if i == 0 else 'td'}>{c}</{'th' if i == 0 else 'td'}>"
                                    for c in row) + "</tr>" for i, row in enumerate(AR_TABLE))
    bullets = "".join(f"<p>• {b}</p>" for b in AR_BULLETS)
    html = (f'<div dir="rtl"><h1>{AR_TITLE}</h1><p class="sub">{AR_SUBTITLE}</p><p>{AR_BODY}</p>'
            f"<h2>{AR_HEADINGS[0]}</h2>{bullets}<h2>{AR_HEADINGS[1]}</h2><table>{rows}</table>"
            f"<p>{AR_MIXED}</p></div>")
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    page.insert_htmlbox(pymupdf.Rect(64, 64, 531, 800), html, css=AR_CSS)
    doc.save(path)


def make_arabic_digital(path):
    """Arabic page with a correct text layer. Glyphs are stored in visual order,
    like real PDFs; Latin runs use Helvetica. (Glyphs are unshaped: the text layer,
    not the look, is what this fixture tests.)"""
    import re
    from bidi.algorithm import get_display

    arabic = re.compile(r"[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]")
    arfont = pymupdf.Font(script=6)   # built-in Noto Naskh Arabic
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    page.insert_font(fontname="noar", fontbuffer=arfont.buffer)

    def width(kind, t, size):
        return arfont.text_length(t, fontsize=size) if kind == "ar" else \
            pymupdf.get_text_length(t, fontname="helv", fontsize=size)

    def line(text, y, size, right=531):
        runs = []
        for ch in get_display(text):
            kind = "ar" if arabic.match(ch) or (ch == " " and runs and runs[-1][0] == "ar") else "lat"
            if runs and runs[-1][0] == kind:
                runs[-1][1] += ch
            else:
                runs.append([kind, ch])
        x = right - sum(width(k, t, size) for k, t in runs)
        for k, t in runs:
            page.insert_text((x, y), t, fontname="noar" if k == "ar" else "helv", fontsize=size)
            x += width(k, t, size)

    y = 90
    line(AR_TITLE, y, 22)
    y += 30
    line(AR_SUBTITLE, y, 10)
    y += 34
    for part in (AR_BODY[:AR_BODY.index("،") + 1], AR_BODY[AR_BODY.index("،") + 2:]):
        line(part, y, 12)
        y += 18
    y += 14
    line(AR_HEADINGS[0], y, 15)
    y += 24
    for b in AR_BULLETS:
        line("• " + b, y, 12)
        y += 18
    y += 10
    line(AR_MIXED, y, 12)
    doc.set_metadata({"title": AR_TITLE, "author": "Synthetic fixture"})
    doc.save(path)


AR_PAGES_HEADER = "تقرير ميداني ربع سنوي"


def make_arabic_pages(path):
    css = AR_CSS + " .hdr { font-size: 10pt; color: #333; } .ftr { font-size: 9pt; color: #333; text-align: center; }"
    doc = pymupdf.open()
    for n in range(1, 4):
        page = doc.new_page(width=595, height=842)
        page.insert_htmlbox(pymupdf.Rect(64, 30, 531, 50), f'<div dir="rtl" class="hdr">{AR_PAGES_HEADER}</div>', css=css)
        body = (f'<div dir="rtl"><h2>القسم {n}</h2><p>{AR_BODY}</p>'
                '<p>1. توريد الأجهزة الطبية</p><p>2. تدريب الفريق</p><p>3. صيانة دورية</p>'
                '<p>١. مراجعة أولية</p><p>٢. موافقة نهائية</p>'
                '<p dir="ltr">Contact: info@example.com</p><p dir="ltr">15,000 2024</p></div>')
        page.insert_htmlbox(pymupdf.Rect(64, 170, 531, 760), body, css=css)
        page.insert_htmlbox(pymupdf.Rect(64, 790, 531, 810), f'<div dir="rtl" class="ftr">صفحة {n} من 3</div>', css=css)
    doc.save(path)


def make_generic(path):
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    x, y = 64, 90
    for line in ("Proposal for the Supply of Medical Equipment", "Ministry of Health"):
        page.insert_text((x, y), line, fontname="hebo", fontsize=20)
        y += 40
    y += 10
    page.insert_text((x, y), "The scope covers the items below.", fontname="helv", fontsize=11)
    y += 24
    for line in ("1. Deliver the equipment", "2. Train the staff"):
        page.insert_text((x, y), line, fontname="helv", fontsize=11)
        y += 18
    y += 10
    page.insert_text((x, y), "Maintenance follows.", fontname="helv", fontsize=11)
    y += 24
    for line in ("3. Maintain the equipment", "4. Report every quarter"):
        page.insert_text((x, y), line, fontname="helv", fontsize=11)
        y += 18
    doc.save(path)


LIGHT_TABLE = [["Phase", "Main activities", "Deliverable"],
               ["Analysis", "Workshops and audience study", "Needs report"],
               ["Design", "Scripts and storyboards", "Storyboards"],
               ["Build", "Animation and voice recording", "Beta version"],
               ["Pilot", "Field test in schools", "Final release"]]


def make_light_table(path):
    rows = "".join("<tr>" + "".join((f"<th>{c}</th>" if i == 0 else f"<td>{c}</td>") for c in row) + "</tr>"
                   for i, row in enumerate(LIGHT_TABLE))
    css = ("body { font-family: sans-serif; font-size: 11pt; color: #1d2b33; } "
           ".banner { background: #17616f; color: #ffffff; font-weight: bold; font-size: 14pt; "
           "padding: 6pt 10pt; margin: 0 0 10pt 0; } p { margin: 0 0 10pt 0; } "
           "table { border-collapse: collapse; width: 100%; } "
           "td { border: 1px solid #cfe2e7; padding: 6pt; } "
           "th { background: #53b6ca; color: #ffffff; padding: 6pt; text-align: left; } "
           "tr:nth-child(odd) td { background: #f5fbfc; }")
    html = (f'<div class="banner">Implementation method</div>'
            f"<p>Each phase ends with a deliverable that the ministry approves.</p>"
            f"<table>{rows}</table>")
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    page.insert_htmlbox(pymupdf.Rect(50, 60, 545, 700), html, css=css)
    doc.save(path)


if __name__ == "__main__":
    make_digital(OUT / "digital.pdf")
    make_scanned(OUT / "digital.pdf", OUT / "scanned.pdf")
    make_arabic_source(OUT / "arabic_shaped.pdf")
    make_scanned(OUT / "arabic_shaped.pdf", OUT / "arabic_scanned.pdf", seed=11)
    make_arabic_digital(OUT / "arabic_digital.pdf")
    make_arabic_pages(OUT / "arabic_pages.pdf")
    make_scanned(OUT / "arabic_pages.pdf", OUT / "arabic_pages_scanned.pdf", seed=3)
    make_generic(OUT / "generic.pdf")
    make_light_table(OUT / "_light_table.pdf")
    make_scanned(OUT / "_light_table.pdf", OUT / "light_table_scanned.pdf", seed=5)
    (OUT / "_light_table.pdf").unlink()
    for name in ("digital.pdf", "scanned.pdf", "arabic_shaped.pdf", "arabic_scanned.pdf", "arabic_digital.pdf",
                 "arabic_pages.pdf", "arabic_pages_scanned.pdf", "generic.pdf", "light_table_scanned.pdf"):
        d = pymupdf.open(OUT / name)
        chars = sum(len(p.get_text().strip()) for p in d)
        print(f"{name}: {d.page_count} pages, {chars} extractable chars, "
              f"{(OUT / name).stat().st_size // 1024} KB")
