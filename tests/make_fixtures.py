"""Generate synthetic test PDFs (no real documents involved).

- digital.pdf : two pages with a real text layer (headings, styled runs, list, table)
- scanned.pdf : the same pages rasterised at 150 dpi with a slight skew and
                noise, saved as image-only pages (no text layer) -> needs OCR
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


if __name__ == "__main__":
    make_digital(OUT / "digital.pdf")
    make_scanned(OUT / "digital.pdf", OUT / "scanned.pdf")
    for name in ("digital.pdf", "scanned.pdf"):
        d = pymupdf.open(OUT / name)
        chars = sum(len(p.get_text().strip()) for p in d)
        print(f"{name}: {d.page_count} pages, {chars} extractable chars, "
              f"{(OUT / name).stat().st_size // 1024} KB")
