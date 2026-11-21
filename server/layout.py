"""Running header/footer detection for PDFs.

Documents exported from Word/LibreOffice repeat a header and footer on every
page (logo, document reference, rule lines, "Page N of M"). Converted naively
they become body text on each page, which breaks the flow in Word: content
overflows and every page gains an extra, almost empty page.

``detect_bands`` finds elements that repeat at the same height near the top or
bottom of most pages; ``strip_bands`` removes them from a working copy of the PDF
before layout conversion; ``docx_fix.add_header_footer`` rebuilds them as real
Word header/footer parts (with PAGE / NUMPAGES fields).
"""
from __future__ import annotations

import math
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pymupdf

TOP_ZONE = 0.16      # share of page height searched for a header
BOTTOM_ZONE = 0.16   # ... and for a footer
MIN_SHARE = 0.6      # an element must repeat on this share of pages


@dataclass
class Span:
    text: str
    font: str
    size: float
    color: int
    bold: bool
    italic: bool
    bbox: tuple


@dataclass
class Item:
    kind: str                     # "text" | "image" | "rule"
    bbox: tuple                   # x0, y0, x1, y1 (pt)
    key: tuple = ()
    spans: list[Span] = field(default_factory=list)
    xref: int = 0                 # image
    color: tuple = (0, 0, 0)      # rule stroke colour (0-1 floats)
    width: float = 0.5            # rule stroke width (pt)
    digit_roles: list = field(default_factory=list)   # per number in the text: "PAGE" | "NUMPAGES" | None

    @property
    def cy(self) -> float:
        return (self.bbox[1] + self.bbox[3]) / 2


@dataclass
class Bands:
    pages: int
    page_size: tuple
    ref_page: int
    header: list[Item]
    footer: list[Item]
    header_cut: dict[int, float]      # page -> y below which the header ends
    footer_cut: dict[int, float]      # page -> y above which the footer starts

    @property
    def left(self) -> float:
        xs = [i.bbox[0] for i in self.header + self.footer]
        return min(xs) if xs else 72.0

    @property
    def right(self) -> float:
        xs = [i.bbox[2] for i in self.header + self.footer]
        return max(xs) if xs else self.page_size[0] - 72.0


def _norm(text: str) -> str:
    return re.sub(r"\d+", "#", re.sub(r"\s+", " ", text)).strip()


def _filled_rects(page: pymupdf.Page) -> list[pymupdf.Rect]:
    return [d["rect"] for d in page.get_drawings()
            if d.get("fill") is not None and d["rect"].width > 4 and d["rect"].height > 4]


def _page_items(page: pymupdf.Page) -> tuple[list[Item], list[tuple]]:
    """Header/footer candidates, plus every content box on the page (obstacles)."""
    w, h = page.rect.width, page.rect.height
    drawings = page.get_drawings()
    fills = [d["rect"] for d in drawings
             if d.get("fill") is not None and d["rect"].width > 4 and d["rect"].height > 4]
    verticals = [d["rect"] for d in drawings if d["rect"].width <= 3 and d["rect"].height > 6]
    items: list[Item] = []
    obstacles: list[tuple] = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] == 1:
            x0, y0, x1, y1 = b["bbox"]
            items.append(Item("image", tuple(b["bbox"]),
                              key=("i", round((x1 - x0) / 3), round((y1 - y0) / 3), round(y0 / 3))))
            obstacles.append(tuple(b["bbox"]))
            continue
        for line in b["lines"]:
            spans = [s for s in line["spans"] if s["text"]]
            text = "".join(s["text"] for s in spans)
            if not text.strip():
                continue
            bbox = tuple(line["bbox"])
            obstacles.append(bbox)
            r = pymupdf.Rect(bbox)
            if any(f.contains(r) for f in fills):   # text on a coloured cell = table, not header
                continue
            items.append(Item("text", bbox, key=("t", _norm(text), round(bbox[1] / 4)), spans=[
                Span(s["text"], s["font"], s["size"], s["color"],
                     bool(s["flags"] & 16) or "bold" in s["font"].lower(),
                     bool(s["flags"] & 2) or "italic" in s["font"].lower() or "oblique" in s["font"].lower(),
                     tuple(s["bbox"])) for s in spans]))

    def table_border(r: pymupdf.Rect) -> bool:
        y = (r.y0 + r.y1) / 2
        for f in fills:   # edge of a coloured cell
            if (abs(f.y0 - y) <= 1.5 or abs(f.y1 - y) <= 1.5) and f.x0 < r.x1 and f.x1 > r.x0:
                return True
        return any(v.y0 - 1.5 <= y <= v.y1 + 1.5 and r.x0 - 1.5 <= v.x0 <= r.x1 + 1.5 for v in verticals)

    seen = set()
    for d in drawings:
        r = d["rect"]
        obstacles.append((r.x0, r.y0, r.x1, r.y1))
        stroke = d.get("color") if d.get("color") is not None else (d.get("fill") if r.height <= 3 else None)
        if r.height <= 3 and r.width >= 0.3 * w and stroke is not None and not table_border(r):
            k = ("r", round(r.y0 / 2), round(r.x0 / 4), round(r.x1 / 4))
            if k in seen:
                continue
            seen.add(k)
            items.append(Item("rule", (r.x0, r.y0, r.x1, r.y1), key=k, color=tuple(stroke),
                              width=float(d.get("width") or max(0.5, r.height))))
    # attach image xrefs
    for info in page.get_image_info(xrefs=True):
        bb = tuple(info["bbox"])
        for it in items:
            if it.kind == "image" and all(abs(a - b) < 1.5 for a, b in zip(it.bbox, bb)):
                it.xref = info.get("xref", 0)
    return items, obstacles


def _same_box(a: tuple, b: tuple, tol: float = 1.5) -> bool:
    return all(abs(x - y) <= tol for x, y in zip(a, b))


def detect_bands(path: str | Path, skip_pages: Optional[set] = None) -> Optional[Bands]:
    skip_pages = skip_pages or set()
    with pymupdf.open(path) as doc:
        n = doc.page_count
        if n < 2:
            return None
        per_page: dict[int, list[Item]] = {}
        obstacles: dict[int, list[tuple]] = {}
        sizes = Counter()
        for page in doc:
            if page.number in skip_pages:
                continue
            per_page[page.number], obstacles[page.number] = _page_items(page)
            sizes[(round(page.rect.width), round(page.rect.height))] += 1
        usable = len(per_page)
        if usable < 2:
            return None
        counts = Counter()
        for items in per_page.values():
            counts.update({it.key for it in items})
        need = max(2, math.ceil(MIN_SHARE * usable))
        repeated = {k for k, c in counts.items() if c >= need}
        if not repeated:
            return None

        header_cut, footer_cut = {}, {}
        tops: dict[int, list[Item]] = {}
        bottoms: dict[int, list[Item]] = {}
        for pno, items in per_page.items():
            h = doc[pno].rect.height
            rep_top = [i for i in items if i.key in repeated and i.cy < TOP_ZONE * h]
            rep_bot = [i for i in items if i.key in repeated and i.cy > (1 - BOTTOM_ZONE) * h]
            band_boxes = [i.bbox for i in rep_top + rep_bot]
            others = [o for o in obstacles[pno] if not any(_same_box(o, b) for b in band_boxes)]
            if rep_top and any(i.kind != "rule" for i in rep_top):
                # a repeated element below some body content is not part of the header
                first_body = min((o[1] for o in others if o[1] < TOP_ZONE * h + 40), default=h)
                rep_top = [i for i in rep_top if i.bbox[3] <= first_body + 0.5]
            if rep_top and any(i.kind != "rule" for i in rep_top):
                bottom = max(i.bbox[3] for i in rep_top)
                nxt = min((o[1] for o in others if o[1] >= bottom - 0.5), default=bottom + 4)
                cut = min(bottom + 1.0, (bottom + nxt) / 2)
                if not any(o[1] < cut and o[3] > 0.5 for o in others if o[3] <= bottom + 0.5
                           and not any(_same_box(o, b) for b in band_boxes)):
                    header_cut[pno] = cut
                    tops[pno] = rep_top
            if rep_bot and any(i.kind != "rule" for i in rep_bot):
                last_body = max((o[3] for o in others if o[3] > (1 - BOTTOM_ZONE) * h - 40), default=0)
                rep_bot = [i for i in rep_bot if i.bbox[1] >= last_body - 0.5]
            if rep_bot and any(i.kind != "rule" for i in rep_bot):
                top = min(i.bbox[1] for i in rep_bot)
                prev = max((o[3] for o in others if o[3] <= top + 0.5), default=top - 4)
                footer_cut[pno] = max(top - 1.0, (top + prev) / 2)
                bottoms[pno] = rep_bot
        if not header_cut and not footer_cut:
            return None
        # reference page: the first page carrying the most repeated band items
        ref = max(per_page, key=lambda p: (len(tops.get(p, [])) + len(bottoms.get(p, [])), -p))
        # Numbers that follow the page index are page numbers; a constant number equal
        # to the page count is the total ("Page 3 of 7").
        for it in tops.get(ref, []) + bottoms.get(ref, []):
            if it.kind != "text":
                continue
            seqs = {}
            for pno, items in per_page.items():
                same = next((o for o in items if o.key == it.key), None)
                if same is not None:
                    seqs[pno] = [int(x) for x in re.findall(r"\d+", "".join(s.text for s in same.spans))]
            mine = seqs.get(ref, [])
            roles = []
            for k, val in enumerate(mine):
                vals = {p: s[k] for p, s in seqs.items() if len(s) > k}
                if len(vals) >= 2 and all(v == p + 1 for p, v in vals.items()):
                    roles.append("PAGE")
                elif len(vals) >= 2 and all(v == n for v in vals.values()) and "PAGE" in roles:
                    roles.append("NUMPAGES")
                else:
                    roles.append(None)
            it.digit_roles = roles
        return Bands(pages=n, page_size=sizes.most_common(1)[0][0], ref_page=ref,
                     header=sorted(tops.get(ref, []), key=lambda i: (i.bbox[1], i.bbox[0])),
                     footer=sorted(bottoms.get(ref, []), key=lambda i: (i.bbox[1], i.bbox[0])),
                     header_cut=header_cut, footer_cut=footer_cut)


def strip_bands(src: str | Path, dst: str | Path, bands: Bands) -> None:
    """Remove the detected header/footer regions from a copy of the PDF."""
    with pymupdf.open(src) as doc:
        for page in doc:
            w, h = page.rect.width, page.rect.height
            hc = bands.header_cut.get(page.number)
            fc = bands.footer_cut.get(page.number)
            if hc:
                page.add_redact_annot(pymupdf.Rect(0, 0, w, hc), fill=False)
            if fc:
                page.add_redact_annot(pymupdf.Rect(0, fc, w, h), fill=False)
            if hc or fc:
                page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_REMOVE,
                                      graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                                      text=pymupdf.PDF_REDACT_TEXT_REMOVE)
        doc.save(dst, garbage=3, deflate=True)


def image_png(path: str | Path, xref: int) -> Optional[bytes]:
    """PNG bytes for an image xref (keeps transparency from a soft mask)."""
    if not xref:
        return None
    with pymupdf.open(path) as doc:
        try:
            pix = pymupdf.Pixmap(doc, xref)
            info = doc.extract_image(xref)
            smask = info.get("smask") if info else 0
            if smask:
                pix = pymupdf.Pixmap(pix, pymupdf.Pixmap(doc, smask))
            if pix.colorspace and pix.colorspace.n > 3:
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            return pix.tobytes("png")
        except Exception:
            return None


def body_rules(path: str | Path, bands: Optional[Bands] = None) -> list[dict]:
    """Free-standing horizontal rules in the body (not table borders, not header/footer)."""
    out = []
    with pymupdf.open(path) as doc:
        for page in doc:
            w = page.rect.width
            hc = bands.header_cut.get(page.number, 0) if bands else 0
            fc = bands.footer_cut.get(page.number, page.rect.height) if bands else page.rect.height
            try:
                tables = [pymupdf.Rect(t.bbox) + (-2, -2, 2, 2) for t in page.find_tables().tables]
            except Exception:
                tables = []
            drawings = page.get_drawings()
            fills = [d["rect"] for d in drawings if d.get("fill") is not None and d["rect"].height > 4]
            lines = [l for b in page.get_text("dict")["blocks"] if b["type"] == 0 for l in b["lines"]
                     if "".join(s["text"] for s in l["spans"]).strip()]
            seen = set()
            for d in drawings:
                r = d["rect"]
                stroke = d.get("color") if d.get("color") is not None else (d.get("fill") if r.height <= 3 else None)
                if stroke is None or r.height > 3 or r.width < 0.3 * w or r.y0 <= hc or r.y1 >= fc:
                    continue
                y = (r.y0 + r.y1) / 2
                if any(t.contains(pymupdf.Point(r.x0 + 3, y)) for t in tables):
                    continue
                if any(abs(f.y0 - y) <= 1.5 or abs(f.y1 - y) <= 1.5 for f in fills if f.x0 < r.x1 and f.x1 > r.x0):
                    continue
                key = (page.number, round(y))
                if key in seen:
                    continue
                seen.add(key)
                above = [l for l in lines if l["bbox"][3] <= y + 0.5 and y - l["bbox"][3] <= 24
                         and l["bbox"][0] < r.x1 and l["bbox"][2] > r.x0]
                if not above:
                    continue
                ln = max(above, key=lambda l: l["bbox"][3])
                out.append({"page": page.number, "bbox": (r.x0, r.y0, r.x1, r.y1), "color": tuple(stroke),
                            "width": float(d.get("width") or max(0.5, r.height)),
                            "gap": max(0.0, y - ln["bbox"][3]),
                            "text_above": "".join(s["text"] for s in ln["spans"])})
    return out
