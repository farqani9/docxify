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
    logical: str = ""             # reading-order text of a right-to-left line

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


def _coloured(fill) -> bool:
    """A visible fill: page-coloured (white) boxes, e.g. the backgrounds browsers paint
    behind headers and footers, don't make the text on them a table cell."""
    return fill is not None and min(fill) < 0.985


def _filled_rects(page: pymupdf.Page) -> list[pymupdf.Rect]:
    return [d["rect"] for d in page.get_drawings()
            if _coloured(d.get("fill")) and d["rect"].width > 4 and d["rect"].height > 4]


def _page_items(page: pymupdf.Page, placeholders: Optional[dict] = None) -> tuple[list[Item], list[tuple]]:
    """Header/footer candidates, plus every content box on the page (obstacles)."""
    w, h = page.rect.width, page.rect.height
    drawings = page.get_drawings()
    fills = [d["rect"] for d in drawings
             if _coloured(d.get("fill")) and d["rect"].width > 4 and d["rect"].height > 4]
    verticals = [d["rect"] for d in drawings if d["rect"].width <= 3 and d["rect"].height > 6]
    items: list[Item] = []
    obstacles: list[tuple] = []
    text_lines = []
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
            if text.strip():
                text_lines.append({"bbox": tuple(line["bbox"]), "spans": spans, "text": text})
    glyphs, raw_lines, claimed = None, None, set()
    for ln in _join_rtl_pieces(text_lines):
        bbox, spans, text = ln["bbox"], ln["spans"], ln["text"]
        obstacles.append(bbox)
        r = pymupdf.Rect(bbox)
        if any(f.contains(r) for f in fills):   # text on a coloured cell = table, not header
            continue
        logical = ""
        token = PLACEHOLDER.fullmatch(text.strip()) if placeholders else None
        if token:                            # an OCR'd right-to-left line: use its text
            real = (placeholders.get(placeholder_key(token.group(0))) or {}).get("text", "")
            if real:
                s0 = dict(spans[0], text=real)
                spans, text, logical = [s0], real, real
        if _is_rtl(text) and not logical:
            if glyphs is None:
                glyphs = page_glyphs(page)
                raw_lines = [(pymupdf.Rect(l["bbox"]), [(ch["c"], ch["bbox"]) for s in l["spans"] for ch in s["chars"]])
                             for b in page.get_text("rawdict")["blocks"] if b["type"] == 0 for l in b["lines"]]
            raw = [c for rb, chars in raw_lines if r.contains((rb.tl + rb.br) / 2) for c in chars]
            logical = " ".join(seg["text"] for seg in reversed(rtl_segments(glyphs, bbox, raw, [], claimed)))
        items.append(Item("text", bbox, key=("t", _norm(text), round(bbox[1] / 4)), spans=[
            Span(s["text"], s["font"], s["size"], s["color"],
                 bool(s["flags"] & 16) or "bold" in s["font"].lower(),
                 bool(s["flags"] & 2) or "italic" in s["font"].lower() or "oblique" in s["font"].lower(),
                 tuple(s["bbox"])) for s in spans], logical=logical))

    def table_border(r: pymupdf.Rect) -> bool:
        y = (r.y0 + r.y1) / 2
        for f in fills:   # edge of a coloured cell
            if (abs(f.y0 - y) <= 1.5 or abs(f.y1 - y) <= 1.5) and f.x0 < r.x1 and f.x1 > r.x0:
                return True
        return any(v.y0 - 1.5 <= y <= v.y1 + 1.5 and r.x0 - 1.5 <= v.x0 <= r.x1 + 1.5 for v in verticals)

    seen = set()
    for d in drawings:
        r = d["rect"]
        if d.get("color") is None and d.get("fill") is not None and not _coloured(d["fill"]):
            continue                          # white background box: invisible, not content
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


def _join_rtl_pieces(lines: list[dict]) -> list[dict]:
    """The extractor can cut one right-to-left line in two ("صفحة 1" | "من 3"):
    rejoin touching pieces on the same row. Left-to-right lines are left alone."""
    lines = sorted(lines, key=lambda l: (l["bbox"][1], l["bbox"][0]))
    out: list[dict] = []
    for ln in lines:
        mate = None
        for prev in reversed(out[-4:]):
            a, c = prev["bbox"], ln["bbox"]
            overlap = min(a[3], c[3]) - max(a[1], c[1])
            if overlap < 0.7 * min(a[3] - a[1], c[3] - c[1]):
                continue
            size = max(s["size"] for s in prev["spans"] + ln["spans"])
            gap = max(a[0], c[0]) - min(a[2], c[2])
            if gap < 0.6 * size and (_is_rtl(prev["text"]) or _is_rtl(ln["text"])):
                mate = prev
                break
        if mate is None:
            out.append(dict(ln))
            continue
        a, c = mate["bbox"], ln["bbox"]
        mate["bbox"] = (min(a[0], c[0]), min(a[1], c[1]), max(a[2], c[2]), max(a[3], c[3]))
        mate["spans"] = sorted(mate["spans"] + ln["spans"], key=lambda s: s["bbox"][0])
        mate["text"] = "".join(s["text"] for s in mate["spans"])
    return out


def _same_box(a: tuple, b: tuple, tol: float = 1.5) -> bool:
    return all(abs(x - y) <= tol for x, y in zip(a, b))


def _match_ocr_items(per_page: dict) -> None:
    """On OCR'd pages a running header differs a little from page to page (recognition
    noise, a few points of residual skew). Give such near-identical items one key:
    same kind, about the same place and size, and similar text."""
    import difflib

    clusters: list[dict] = []

    def text_of(it):
        return _norm("".join(s.text for s in it.spans)) if it.kind == "text" else ""

    for pno in sorted(per_page):
        for it in per_page[pno]:
            t = text_of(it)
            for c in clusters:
                a, b = c["bbox"], it.bbox
                if c["kind"] != it.kind or c["page"] == pno:
                    continue
                if abs(a[1] - b[1]) > 8 or abs(a[3] - b[3]) > 8:
                    continue
                if abs(a[0] - b[0]) > 15 and abs(a[2] - b[2]) > 15:      # left or right aligned
                    continue
                wa, wb = a[2] - a[0], b[2] - b[0]
                if it.kind != "rule" and abs(wa - wb) > 0.25 * max(wa, wb, 1):
                    continue
                if it.kind == "text" and difflib.SequenceMatcher(None, c["text"], t).ratio() < 0.5:
                    continue
                it.key = c["key"]
                break
            else:
                clusters.append({"kind": it.kind, "bbox": it.bbox, "text": t, "key": it.key, "page": pno})


def detect_bands(path: str | Path, skip_pages: Optional[set] = None,
                 placeholders: Optional[dict] = None, fuzzy: bool = False) -> Optional[Bands]:
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
            per_page[page.number], obstacles[page.number] = _page_items(page, placeholders)
            sizes[(round(page.rect.width), round(page.rect.height))] += 1
        usable = len(per_page)
        if usable < 2:
            return None
        if fuzzy:
            _match_ocr_items(per_page)
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
            others_top = others
            if fuzzy and rep_top:
                # OCR'd pages: a header line read too differently to match is still part
                # of the band when it sits within it, not body content above the header
                band_bottom = max(i.bbox[3] for i in rep_top)
                others_top = [o for o in others if o[3] > band_bottom + 0.5]
            if rep_top and any(i.kind != "rule" for i in rep_top):
                # a repeated element below some body content is not part of the header
                first_body = min((o[1] for o in others_top if o[1] < TOP_ZONE * h + 40), default=h)
                rep_top = [i for i in rep_top if i.bbox[3] <= first_body + 0.5]
            if rep_top and any(i.kind != "rule" for i in rep_top):
                bottom = max(i.bbox[3] for i in rep_top)
                nxt = min((o[1] for o in others_top if o[1] >= bottom - 0.5), default=bottom + 4)
                cut = min(bottom + 1.0, (bottom + nxt) / 2)
                if fuzzy or not any(o[1] < cut and o[3] > 0.5 for o in others if o[3] <= bottom + 0.5
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
            roles: list = [None] * len(mine)
            columns = [{p: s[k] for p, s in seqs.items() if len(s) > k} for k in range(len(mine))]
            for k, vals in enumerate(columns):
                if len(vals) >= 2 and all(v == p + 1 for p, v in vals.items()):
                    roles[k] = "PAGE"
            if "PAGE" in roles:          # the total may come before the page number (RTL)
                for k, vals in enumerate(columns):
                    if roles[k] is None and len(vals) >= 2 and all(v == n for v in vals.values()):
                        roles[k] = "NUMPAGES"
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
            # generous outer edges: shapes painted past the page box (browser backgrounds,
            # full-bleed rules) only count as covered when the redaction encloses them
            if hc:
                page.add_redact_annot(pymupdf.Rect(-w, -h, 2 * w, hc), fill=False)
            if fc:
                page.add_redact_annot(pymupdf.Rect(-w, fc, 2 * w, 2 * h), fill=False)
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
            glyphs = None
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
                text_above = "".join(s["text"] for s in ln["spans"])
                if _is_rtl(text_above):          # extraction order differs from the DOCX text
                    glyphs = page_glyphs(page) if glyphs is None else glyphs
                    segs = rtl_segments(glyphs, tuple(ln["bbox"]), [], [])
                    text_above = " ".join(seg["text"] for seg in reversed(segs)) or text_above
                out.append({"page": page.number, "bbox": (r.x0, r.y0, r.x1, r.y1), "color": tuple(stroke),
                            "width": float(d.get("width") or max(0.5, r.height)),
                            "gap": max(0.0, y - ln["bbox"][3]),
                            "text_above": text_above})
    return out


# ------------------------------------------------------------------ right-to-left
# Placeholder tokens stand in for right-to-left lines during layout conversion:
# "QXR" + index + "Z", then "Z" filler that stretches the token over the line's width.
# Five digits keep tokens short enough for one-word lines (a token wider than its line
# is drawn smaller, which changes how the layout converter groups lines); a sixth digit
# is only used past 99,999 lines.
PLACEHOLDER = re.compile(r"QXR\d{5,6}Z+")
_KEY = re.compile(r"QXR\d{5,6}Z")


def placeholder(index: int) -> str:
    return f"QXR{index:05d}Z" if index < 100000 else f"QXR{index:06d}Z"


def placeholder_key(token: str) -> str:
    m = _KEY.match(token.strip())
    return m.group(0) if m else token


_RTL = re.compile(r"[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]")
_MARK = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED\u0591-\u05C7]")
_LATIN = re.compile(r"[A-Za-z\u00C0-\u024F]")


def _is_rtl(text: str) -> bool:
    r = len(_RTL.findall(text))
    return r > 0 and r >= len(_LATIN.findall(text))


_SWAP = str.maketrans("()[]{}<>\u00ab\u00bb", ")(][}{><\u00bb\u00ab")


def _nested(text: str) -> bool:
    stack = []
    for c in text:
        if c in "([{":
            stack.append(c)
        elif c in ")]}":
            if not stack or "([{"[")]}".index(c)] != stack.pop():
                return False
    return not stack


def _logical_from_visual(visual: str) -> str:
    """Logical text of a visual-order RTL line. PDF producers disagree on whether a
    mirrored bracket glyph maps to the character drawn or to the logical one; take
    the reading whose brackets nest properly."""
    from bidi.algorithm import get_display
    text = re.sub(r"\s+", " ", get_display(visual, base_dir="R")).strip()
    if not _nested(text):
        alt = text.translate(_SWAP)
        if _nested(alt):
            return alt
    return text


def logical_from_chars(chars: list) -> str:
    """Reading-order text of one RTL line from positioned characters [(c, bbox), ...].

    Characters sorted by x give the true visual order (independent of any reordering
    done by the PDF text extractor); the Unicode bidi algorithm with a right-to-left
    base direction turns that back into logical order. Diacritics are kept with their
    base letter; spaces are added where the PDF positions words apart without one."""
    import unicodedata

    bases, marks = [], []
    for c, bb in chars:
        if not c:
            continue
        cx = (bb[0] + bb[2]) / 2
        (marks if unicodedata.category(c[0]) == "Mn" else bases).append((cx, c, bb))
    if not bases:
        return ""
    bases.sort(key=lambda t: t[0])
    attached: dict[int, list] = {}
    for cx, c, _bb in marks:
        k = min(range(len(bases)), key=lambda i: abs(bases[i][0] - cx))
        attached.setdefault(k, []).append(c)
    visual, prev = [], None
    for i, (cx, c, bb) in enumerate(bases):
        if prev is not None and c != " " and visual and visual[-1] != " ":
            size = max(1.0, bb[3] - bb[1])
            if bb[0] - prev[2] > 0.2 * size:
                visual.append(" ")
        logical = unicodedata.normalize("NFKC", c) if _RTL.search(c) else c
        # a ligature glyph (e.g. lam-alef) decomposes to logical order; store it visually
        visual.extend(attached.get(i, []))
        visual.append(logical[::-1] if len(logical) > 1 and _RTL.search(logical) else logical)
        prev = bb
    return _logical_from_visual("".join(visual))


def page_glyphs(page: pymupdf.Page) -> list[dict]:
    """Every glyph as drawn, in content-stream order (before any text-extraction
    reordering): unicode as mapped by the PDF, pen position, advance, span index."""
    out = []
    for k, sp in enumerate(page.get_texttrace()):
        for u, _gid, org, bb in sp["chars"]:
            out.append({"u": u, "x": org[0], "y": org[1], "x1": max(bb[2], org[0]),
                        "w": bb[2] - bb[0], "span": k, "size": sp["size"], "i": len(out)})
    return out


def _is_mark(g: dict) -> bool:
    import unicodedata

    if g["w"] > 0.05:
        return False
    return g["u"] == 0xFFFD or (g["u"] > 0 and unicodedata.category(chr(g["u"])) == "Mn")


def page_verticals(page: pymupdf.Page) -> list[tuple]:
    """Vertical rules and box edges (x, y0, y1): table cell borders."""
    out = []
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.x - b.x) < 0.8 and abs(a.y - b.y) > 4:
                    out.append(((a.x + b.x) / 2, min(a.y, b.y), max(a.y, b.y)))
            elif it[0] == "re":
                r = it[1]
                if r.height > 4 and r.width > 2:
                    out += [(r.x0, r.y0, r.y1), (r.x1, r.y0, r.y1)]
                elif r.height > 4:                       # a thin filled bar is a rule
                    out.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
    return out


def rtl_segments(glyphs: list[dict], bbox: tuple, raw_chars: list, verticals: list,
                 claimed: Optional[set] = None) -> list[dict]:
    """Reading-order text of the glyphs inside one line box, cut at table borders.

    Glyphs are taken as drawn: sorted by position they give the visual order, with
    presentation forms (unshaped by NFKC) and brackets in their drawn shape, so one
    pass of the Unicode bidi algorithm (right-to-left base) yields the logical text.
    Marks are drawn right before their base letter, in the same span; a mark glyph
    the PDF can't map (stacked harakat) takes the marks MuPDF reports at its position.
    Returns segments left to right: [{"text", "x0", "x1"}]."""
    import unicodedata

    x0, y0, x1, y1 = bbox
    inside = [g for g in glyphs if y0 - 0.5 <= g["y"] <= y1 + 0.5 and x0 - 1.5 <= g["x"] <= x1 + 1.5
              and (claimed is None or g["i"] not in claimed)]
    if not inside:
        return []
    if claimed is not None:
        claimed.update(g["i"] for g in inside)
    bases = [g for g in inside if not _is_mark(g)]
    if not bases:
        return []
    marks: dict[int, list] = {}
    for k, g in enumerate(inside):
        if not _is_mark(g):
            continue
        nxt = next((h for h in inside[k + 1:] if not _is_mark(h)), None)
        prv = next((h for h in reversed(inside[:k]) if not _is_mark(h)), None)
        base = nxt if nxt is not None and nxt["span"] == g["span"] else \
            prv if prv is not None and prv["span"] == g["span"] else nxt or prv
        if g["u"] == 0xFFFD:
            chars = [c for c, bb in raw_chars if abs(bb[0] - g["x"]) < 0.4
                     and unicodedata.category(c) == "Mn"]
        else:
            chars = [chr(g["u"])]
        marks.setdefault(base["i"], []).extend(chars)

    def glyph_text(g: dict) -> str:
        if g["u"] == 0xFFFD:   # unmapped glyph: the letters the extractor reports there
            t = "".join(c for c, bb in raw_chars if g["x"] - 0.4 <= bb[0] < g["x1"] - 0.4
                        and unicodedata.category(c)[0] in "LM")
        else:
            t = chr(g["u"]) if g["u"] > 0 else ""
        if t and (0xFB00 <= ord(t[0]) <= 0xFDFF or 0xFE70 <= ord(t[0]) <= 0xFEFF):
            t = unicodedata.normalize("NFKC", t)        # presentation forms, ligatures
        return t

    bases.sort(key=lambda g: g["x"])
    cuts = set()
    # a cell border reaches past the text's line box (cells have padding); the edges
    # of background boxes painted behind single words don't
    walls = [vx for vx, vy0, vy1 in verticals if vy0 <= y0 - 0.5 and vy1 >= y1 + 0.5]
    for k in range(len(bases) - 1):
        a, b = bases[k], bases[k + 1]
        if any(a["x1"] - 0.5 <= vx <= b["x"] + 0.5 for vx in walls):
            cuts.add(k)
    segments, cur = [], []
    for k, g in enumerate(bases):
        cur.append(g)
        if k in cuts or k == len(bases) - 1:
            segments.append(cur)
            cur = []
    out = []
    for seg in segments:
        visual, prev = [], None
        for g in seg:
            t = glyph_text(g)
            if not t:
                continue
            rtl = bool(_RTL.search(t))
            if prev is not None and t != " " and visual and visual[-1] != " " \
                    and g["x"] - prev["x1"] > 0.2 * max(1.0, g["size"]):
                visual.append(" ")                     # words set apart without a space glyph
            ms = marks.get(g["i"], [])
            if rtl:                                    # stored visually: reversed, marks first
                visual += list(reversed(ms)) + [t[::-1]]
            else:
                visual += [t] + ms
            prev = g
        text = _logical_from_visual("".join(visual))
        if text:
            out.append({"text": text, "x0": seg[0]["x"], "x1": max(g["x1"] for g in seg)})
    return out


def logical_line_text(spans: list) -> str:
    """Span-level variant (header/footer items): RTL spans arrive logically ordered."""
    import unicodedata

    visual = ""
    for s in sorted((s for s in spans if s.get("text")), key=lambda s: s["bbox"][0]):
        t = unicodedata.normalize("NFKC", s["text"])
        visual += t[::-1] if _RTL.search(t) else t
    return _logical_from_visual(visual)


def fit_token(token: str, width: float, size: float, font: str = "helv") -> tuple[str, float]:
    """Placeholder text and font size that cover ``width`` without stretching glyphs
    (a stretched font matrix would be reported as a bigger font size)."""
    natural = pymupdf.get_text_length(token, fontname=font, fontsize=size)
    if natural >= width:
        return token, max(2.0, size * width / natural) if natural else size
    filler = pymupdf.get_text_length("Z", fontname=font, fontsize=size)
    return token + "Z" * min(200, int((width - natural) / filler)), size


def _rtl_rows(page: pymupdf.Page, verticals: list, glyphs: Optional[list] = None) -> list[dict]:
    """Right-to-left text lines, with pieces of one visual line rejoined.

    The extractor sometimes cuts a line in two (around a diacritic, say); pieces on
    the same baseline that touch, with no cell border between them, are one line."""
    lines = []
    for b in page.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for line in b["lines"]:
            spans = [s for s in line["spans"] if s["chars"]]
            if not spans:
                continue
            chars = [(ch["c"], ch["bbox"]) for s in spans for ch in s["chars"]]
            # baseline of the letters (a raised mark can start the line)
            ys = sorted(ch["origin"][1] for s in spans for ch in s["chars"]
                        if ch["c"].strip() and not _MARK.match(ch["c"]))
            base = ys[len(ys) // 2] if ys else spans[0]["origin"][1]
            lines.append({"spans": spans, "chars": chars, "bbox": tuple(line["bbox"]),
                          "base": base, "size": max(s["size"] for s in spans)})
            lines[-1]["first"] = lines[-1]
    # rows: same top edge or same baseline (a diacritic can push one piece's baseline)
    lines.sort(key=lambda l: (l["bbox"][1], l["bbox"][0]))
    rows: list[list] = []
    for ln in lines:
        row = rows[-1] if rows else None
        if row is not None:
            a, c = row[0]["bbox"], ln["bbox"]
            overlap = min(a[3], c[3]) - max(a[1], c[1])
            same = abs(row[0]["base"] - ln["base"]) <= 1.0 or overlap >= 0.7 * min(a[3] - a[1], c[3] - c[1])
        if row is not None and same:
            row.append(ln)
        else:
            rows.append([ln])
    ordered = [ln for row in rows for ln in sorted(row, key=lambda l: l["bbox"][0])]
    merged: list[dict] = []
    row_of = {id(ln): k for k, row in enumerate(rows) for ln in row}
    for ln in ordered:
        prev = merged[-1] if merged else None
        if prev is not None and row_of[id(prev["first"])] == row_of[id(ln)]:
            gap_lo, gap_hi = prev["bbox"][2], ln["bbox"][0]
            ly0, ly1 = min(prev["bbox"][1], ln["bbox"][1]), max(prev["bbox"][3], ln["bbox"][3])
            wall = any(min(gap_lo, gap_hi) - 0.5 <= vx <= max(gap_lo, gap_hi) + 0.5
                       and vy0 <= ly0 - 0.5 and vy1 >= ly1 + 0.5 for vx, vy0, vy1 in verticals)
            # extractor boxes can stop short of a wide glyph (final-form letters):
            # a drawn glyph spanning the gap joins the two pieces
            bridged = any(g["x"] <= gap_lo + 1 and g["x1"] >= gap_hi - 1 and abs(g["y"] - ln["base"]) <= 3
                          for g in glyphs or [])
            if (gap_hi - gap_lo < 0.6 * ln["size"] or bridged) and not wall:
                prev["spans"] += ln["spans"]
                prev["chars"] += ln["chars"]
                a, c = prev["bbox"], ln["bbox"]
                prev["bbox"] = (min(a[0], c[0]), min(a[1], c[1]), max(a[2], c[2]), max(a[3], c[3]))
                continue
        merged.append(dict(ln, spans=list(ln["spans"]), chars=list(ln["chars"])))
    return [ln for ln in merged if _is_rtl("".join(c for c, _ in ln["chars"]))]


def _bullet_dots(page: pymupdf.Page) -> list[pymupdf.Rect]:
    """Small filled round shapes: list bullets that browsers draw as paths."""
    out = []
    for d in page.get_drawings():
        r = d["rect"]
        if d.get("fill") is None or not (1.2 <= r.width <= 7 and 1.2 <= r.height <= 7):
            continue
        if abs(r.width - r.height) > 0.35 * max(r.width, r.height):
            continue
        kinds = {it[0] for it in d["items"]}
        if kinds <= {"c", "l", "re", "qu"} and ("c" in kinds or "re" in kinds or "qu" in kinds):
            out.append(r)
    return out


def rewrite_rtl_lines(src: str | Path, dst: str | Path, start: int = 0,
                      skip_pages: Optional[set] = None) -> dict:
    """Replace every right-to-left text line by a placeholder token over the same box.

    The layout converter mis-orders RTL text (and mixed RTL/LTR lines); with tokens
    it only has to place the lines, and docx_fix restores the logical text.
    Returns {token: {"text", "size", "bold"}}; writes nothing if there is no RTL text."""
    mapping: dict[str, dict] = {}
    skip_pages = skip_pages or set()
    with pymupdf.open(src) as doc:
        for page in doc:
            if page.number in skip_pages or not _RTL.search(page.get_text()):
                continue                     # cheap pre-check: most pages have no RTL text
            jobs = []
            glyphs, verticals, claimed = page_glyphs(page), page_verticals(page), set()
            dots, used_dots = _bullet_dots(page), []
            for line in _rtl_rows(page, verticals, glyphs):
                spans, raw_chars, lb = line["spans"], line["chars"], line["bbox"]
                segs = rtl_segments(glyphs, lb, raw_chars, verticals, claimed)
                if not segs:                         # no glyph trace: position-sorted characters
                    text = logical_from_chars(raw_chars)
                    segs = [{"text": text, "x0": lb[0], "x1": lb[2]}] if text else []
                main = max(spans, key=lambda s: len(s["chars"]))
                n_all = sum(len(s["chars"]) for s in spans)
                n_bold = sum(len(s["chars"]) for s in spans if s["flags"] & 16 or "bold" in s["font"].lower())
                if segs:
                    # a drawn dot just right of the line's start is its list bullet
                    last = segs[-1]
                    size = max(s["size"] for s in spans)
                    dot = next((d for d in dots if lb[1] - 1 <= (d.y0 + d.y1) / 2 <= lb[3] + 1
                                and 0 <= d.x0 - last["x1"] <= 2.5 * size and d not in used_dots), None)
                    if dot is not None and not any(dot.x0 - 0.5 <= vx <= dot.x1 + 0.5 for vx, vy0, vy1 in verticals
                                                   if vy0 <= (dot.y0 + dot.y1) / 2 <= vy1 and vx < dot.x0 - 0.5 and vx > last["x1"]):
                        used_dots.append(dot)
                        last["text"] = "\u2022 " + last["text"]
                        last["x1"] = dot.x1
                for seg in segs:
                    if not _RTL.search(seg["text"]):
                        continue                     # numbers-only cell: leave the original text
                    jobs.append({"bbox": pymupdf.Rect(seg["x0"], lb[1], seg["x1"], lb[3]),
                                 "text": seg["text"], "size": round(max(s["size"] for s in spans) * 2) / 2,
                                 "color": main["color"], "bold": n_bold > n_all / 2,
                                 "baseline": main["origin"][1], "font": main["font"]})
            if not jobs:
                continue
            for j in jobs:
                r = j["bbox"]
                mid = (r.y0 + r.y1) / 2
                band = max(1.0, 0.25 * (r.y1 - r.y0))   # only this line's glyphs
                page.add_redact_annot(pymupdf.Rect(r.x0, mid - band, r.x1, mid + band), fill=False)
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                  graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                                  text=pymupdf.PDF_REDACT_TEXT_REMOVE)
            if used_dots:                            # the bullets now live in the text
                for d in used_dots:
                    page.add_redact_annot(d + (-0.5, -0.5, 0.5, 0.5), fill=False)
                page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                      graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                                      text=pymupdf.PDF_REDACT_TEXT_NONE)
            for j in jobs:
                token = placeholder(start + len(mapping))
                mapping[token] = {"text": j["text"], "size": j["size"], "bold": j["bold"], "font": j["font"]}
                font = "hebo" if j["bold"] else "helv"
                text, size = fit_token(token, j["bbox"].width, j["size"], font)
                c = j["color"]
                page.insert_text((j["bbox"].x0, j["baseline"]), text, fontname=font, fontsize=size,
                                 color=((c >> 16 & 255) / 255, (c >> 8 & 255) / 255, (c & 255) / 255))
        if mapping:
            doc.save(dst, garbage=3, deflate=True)
    return mapping
