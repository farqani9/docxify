"""PDF -> Word (DOCX) conversion engine with OCR for scanned pages.

Pipeline:

1. ``analyze_pdf``  - per-page scan detection: little/no extractable text but
   images covering the page => "scanned". The upload endpoint returns this flag
   and the UI only offers OCR when it is set.
2. ``convert``      - without OCR the PDF goes straight to pdf2docx (scanned
   pages become non-editable images in Word).
   With OCR every scanned page is rebuilt as a clean, text-only PDF page:
   OCR text lines (RapidOCR, PP-OCRv5) are placed at their ink-tight positions
   with fitted, clustered font sizes; bullet dots the recogniser skips are
   recovered; table rule lines found in the scan are redrawn as vectors so the
   table parser can rebuild real Word tables. Pages that already have a text
   layer are copied unchanged.
3. ``enrich_docx``  - adds document semantics on top of the layout conversion:
   wrapped lines are reflowed into paragraphs, Heading 1-3 styles come from
   font-size ranks, bullet/numbered lists get list styles, the title is set.
"""
from __future__ import annotations

import copy
import logging
import math
import re
import tempfile
import threading
import time
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

import cv2
import numpy as np
import pymupdf

from . import docx_fix, layout

log = logging.getLogger("docxify.engine")

# progress(stage, done, total)  stage in {"analyze","ocr","layout","write","enrich"}
Progress = Callable[[str, int, int], None]

SCAN_MAX_CHARS = 25          # a page with fewer extractable characters ...
SCAN_MIN_IMAGE_COVER = 0.35  # ... and images covering at least this share is "scanned"
OCR_DPI = 200
OCR_MIN_SCORE = 0.5
OCR_LANGS = ("en", "latin")


class ConversionError(Exception):
    """User-facing conversion failure."""


# --------------------------------------------------------------------- analysis
@dataclass
class PageInfo:
    index: int
    chars: int
    image_cover: float
    scanned: bool


@dataclass
class PdfInfo:
    pages: int
    encrypted: bool = False
    title: str = ""
    page_info: list[PageInfo] = field(default_factory=list)

    @property
    def scanned_pages(self) -> list[int]:
        return [p.index for p in self.page_info if p.scanned]

    @property
    def scanned(self) -> bool:
        return bool(self.scanned_pages)


def analyze_pdf(path: str | Path) -> PdfInfo:
    """Open a PDF and classify each page as text-based or scanned."""
    try:
        doc = pymupdf.open(path)
    except Exception as exc:  # corrupt / not a PDF
        raise ConversionError("This file is damaged or is not a PDF.") from exc
    with doc:
        if not doc.is_pdf:
            raise ConversionError("This file is not a PDF.")
        if doc.needs_pass:
            return PdfInfo(pages=doc.page_count, encrypted=True)
        info = PdfInfo(pages=doc.page_count, title=(doc.metadata or {}).get("title") or "")
        for page in doc:
            chars = len(page.get_text("text").strip())
            area = max(1.0, page.rect.width * page.rect.height)
            covered = 0.0
            for img in page.get_image_info():
                r = pymupdf.Rect(img["bbox"]) & page.rect
                if not r.is_empty:
                    covered += r.width * r.height
            cover = min(1.0, covered / area)
            info.page_info.append(PageInfo(
                index=page.number, chars=chars, image_cover=round(cover, 3),
                scanned=chars < SCAN_MAX_CHARS and cover >= SCAN_MIN_IMAGE_COVER))
        return info


def render_thumbnail(path: str | Path, width_px: int = 240) -> bytes:
    with pymupdf.open(path) as doc:
        page = doc[0]
        zoom = width_px / max(1.0, page.rect.width)
        return page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False).tobytes("png")


# --------------------------------------------------------------------------- OCR
_engines: dict[str, object] = {}
_engine_lock = threading.Lock()   # guards creation and inference (not thread-safe)


def _ocr_engine(lang: str):
    from rapidocr import LangRec, ModelType, OCRVersion, RapidOCR

    if lang not in OCR_LANGS:
        lang = "en"
    if lang not in _engines:
        _engines[lang] = RapidOCR(params={
            "Global.log_level": "error",
            "Rec.lang_type": LangRec.EN if lang == "en" else LangRec.LATIN,
            "Rec.ocr_version": OCRVersion.PPOCRV5,
            "Rec.model_type": ModelType.MOBILE,
        })
    return _engines[lang]


@dataclass
class OcrLine:
    text: str
    score: float
    x0: float
    y0: float
    x1: float
    y1: float                 # ink-tight box in the deskewed image (px)
    bullet_x: Optional[float] = None
    size: float = 0.0         # assigned font size (pt)


def _page_gray(page: pymupdf.Page, dpi: int) -> np.ndarray:
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, alpha=False)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width).copy()


def _skew_from_boxes(boxes: list[np.ndarray]) -> float:
    """Length-weighted median text-line angle (degrees, image y-down)."""
    angles, weights = [], []
    for b in boxes:
        (x0, y0), (x1, y1) = b[0], b[1]           # top-left -> top-right edge
        w = math.hypot(x1 - x0, y1 - y0)
        if w < 60:
            continue
        a = math.degrees(math.atan2(y1 - y0, x1 - x0))
        if abs(a) <= 10:
            angles.append(a)
            weights.append(w)
    if not angles:
        return 0.0
    order = np.argsort(angles)
    cum = np.cumsum(np.array(weights)[order])
    median = float(np.array(angles)[order][np.searchsorted(cum, cum[-1] / 2)])
    return median if abs(median) >= 0.15 else 0.0


def _tighten(binary: np.ndarray, x0: float, y0: float, x1: float, y1: float):
    """Shrink a padded OCR box to the ink it contains."""
    h, w = binary.shape
    ix0, iy0 = max(0, int(x0)), max(0, int(y0))
    ix1, iy1 = min(w, int(math.ceil(x1)) + 1), min(h, int(math.ceil(y1)) + 1)
    crop = binary[iy0:iy1, ix0:ix1]
    if crop.size == 0:
        return x0, y0, x1, y1
    cols = np.where((crop > 0).sum(axis=0) >= 2)[0]
    rows = np.where((crop > 0).sum(axis=1) >= 2)[0]
    if len(cols) == 0 or len(rows) == 0:
        return x0, y0, x1, y1
    return ix0 + cols[0], iy0 + rows[0], ix0 + cols[-1] + 1, iy0 + rows[-1] + 1


def _find_bullet(binary: np.ndarray, ln: OcrLine, others: list[OcrLine]) -> Optional[float]:
    """A small round blob just left of a text line is a bullet the recogniser skipped."""
    th = ln.y1 - ln.y0
    if th < 6:
        return None
    rx0 = int(max(0, ln.x0 - 3.5 * th))
    rx1 = int(ln.x0 - 2)
    ry0, ry1 = int(ln.y0), int(ln.y1) + 1
    if rx1 - rx0 < 4:
        return None
    region = binary[ry0:ry1, rx0:rx1]
    n, _, stats, _ = cv2.connectedComponentsWithStats(region, connectivity=8)
    best = None
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if area < 6 or w > 0.7 * th or h > 0.7 * th or not (0.5 <= w / max(1, h) <= 2.0):
            continue
        if area / float(w * h) < 0.5:          # dots/squares are compact
            continue
        bx = rx0 + x
        if any(o is not ln and o.x0 - 2 <= bx <= o.x1 + 2 and o.y0 - 2 <= ry0 + y <= o.y1 + 2
               for o in others):
            continue
        if best is None or bx > best:
            best = bx
    return float(best) if best is not None else None


def _detect_rules(binary: np.ndarray) -> list[tuple[str, float, float, float, float, float]]:
    """Horizontal/vertical ruling lines (table borders) in a deskewed scan."""
    h, w = binary.shape
    horiz = cv2.morphologyEx(binary, cv2.MORPH_OPEN,
                             cv2.getStructuringElement(cv2.MORPH_RECT, (max(40, w // 25), 1)))
    vert = cv2.morphologyEx(binary, cv2.MORPH_OPEN,
                            cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(30, h // 45))))
    segs: list[tuple[str, float, float, float, float, float]] = []
    for mask, kind in ((horiz, "h"), (vert, "v")):
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for c in contours:
            x, y, cw, ch = cv2.boundingRect(c)
            if kind == "h" and cw >= w * 0.04 and ch <= 10:
                segs.append(("h", x, y + ch / 2, x + cw, y + ch / 2, ch))
            elif kind == "v" and ch >= h * 0.012 and cw <= 10:
                segs.append(("v", x + cw / 2, y, x + cw / 2, y + ch, cw))

    def touches(a, b, tol=6):  # keep only lines that are part of a grid
        (_, ax0, ay0, ax1, ay1, _), (_, bx0, by0, bx1, by1, _) = a, b
        if a[0] == "h":
            return ax0 - tol <= bx0 <= ax1 + tol and by0 - tol <= ay0 <= by1 + tol
        return bx0 - tol <= ax0 <= bx1 + tol and ay0 - tol <= by0 <= ay1 + tol

    hs = [s for s in segs if s[0] == "h"]
    vs = [s for s in segs if s[0] == "v"]
    return [s for s in hs if any(touches(s, v) for v in vs)] + \
           [s for s in vs if any(touches(s, hh) for hh in hs)]


def ocr_page(page: pymupdf.Page, lang: str = "en") -> tuple[list[OcrLine], list, tuple[int, int]]:
    """OCR one page. Returns text lines, table rules (both in deskewed px) and image size."""
    gray = _page_gray(page, OCR_DPI)
    h, w = gray.shape
    with _engine_lock:
        result = _ocr_engine(lang)(cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR))
    boxes = [np.asarray(b, dtype=np.float32) for b in (result.boxes if result.boxes is not None else [])]
    texts = list(result.txts or [])
    scores = [float(s) for s in (result.scores or [])]

    # Deskew: rotate the image and the OCR boxes with the same matrix.
    angle = _skew_from_boxes(boxes)
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    if angle:
        gray = cv2.warpAffine(gray, m, (w, h), flags=cv2.INTER_LINEAR, borderValue=255)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                   cv2.THRESH_BINARY_INV, 25, 15)

    lines: list[OcrLine] = []
    for box, text, score in zip(boxes, texts, scores):
        text = (text or "").strip()
        if not text or score < OCR_MIN_SCORE:
            continue
        pts = cv2.transform(box.reshape(-1, 1, 2), m).reshape(-1, 2) if angle else box
        x0, y0 = pts.min(axis=0)
        x1, y1 = pts.max(axis=0)
        lines.append(OcrLine(text, score, *_tighten(binary, x0, y0, x1, y1)))
    for ln in lines:
        ln.bullet_x = _find_bullet(binary, ln, lines)

    return lines, _detect_rules(binary), (w, h)


_DESC = set("gjpqy,;()[]{}|/@$")


def _assign_sizes(lines: list[OcrLine], pt_per_px: float) -> None:
    """Font size per line from ink-tight metrics, then snapped to size clusters."""
    for ln in lines:
        w_pt, h_pt = (ln.x1 - ln.x0) * pt_per_px, (ln.y1 - ln.y0) * pt_per_px
        unit = pymupdf.get_text_length(ln.text, fontname="helv", fontsize=1)
        by_w = w_pt / unit if unit > 0 else 0
        # glyph extent: cap height 0.718 (+ descender 0.207 when present)
        by_h = h_pt / (0.925 if _DESC & set(ln.text) else 0.718)
        n = len(ln.text)
        if n >= 10 and by_w:
            ln.size = by_w
        elif by_w:
            ln.size = (by_w + by_h) / 2
        else:
            ln.size = by_h
    if not lines:
        return
    # 1-D clustering: consecutive sizes within 9% belong together; use the
    # character-weighted median so long lines (most reliable) dominate.
    order = sorted(lines, key=lambda l: l.size)
    groups, cur = [], [order[0]]
    for ln in order[1:]:
        if ln.size <= cur[0].size * 1.09:
            cur.append(ln)
        else:
            groups.append(cur)
            cur = [ln]
    groups.append(cur)
    for g in groups:
        expanded = sorted(l.size for l in g for _ in range(max(1, len(l.text))))
        value = expanded[len(expanded) // 2]
        for ln in g:
            ln.size = max(4.0, round(value * 2) / 2)
    # Very short strings ("pH", "7.1") give unreliable estimates: take the size of
    # longer text on the same row (typically the other cells of a table row).
    for ln in lines:
        if len(ln.text) > 4:
            continue
        cy, half = (ln.y0 + ln.y1) / 2, (ln.y1 - ln.y0)
        mates = [o.size for o in lines
                 if o is not ln and len(o.text) > 4 and abs((o.y0 + o.y1) / 2 - cy) <= half]
        if mates:
            ln.size = sorted(mates)[len(mates) // 2]


def build_ocr_pdf(src: str | Path, dst: str | Path, scanned_pages: list[int], lang: str,
                  progress: Optional[Progress] = None) -> dict:
    """Write a copy of ``src`` where every scanned page is replaced by a text page."""
    stats = {"ocr_pages": 0, "ocr_lines": 0, "ocr_bullets": 0, "table_rules": 0}
    todo = set(scanned_pages)
    with pymupdf.open(src) as doc, pymupdf.open() as out:
        for idx in range(doc.page_count):
            page = doc[idx]
            if idx not in todo:
                out.insert_pdf(doc, from_page=idx, to_page=idx)
                continue
            if progress:
                progress("ocr", stats["ocr_pages"], len(todo))
            lines, rules, (w, h) = ocr_page(page, lang)
            pw, ph = page.rect.width, page.rect.height
            sx, sy = pw / w, ph / h
            _assign_sizes(lines, sx)
            newp = out.new_page(width=pw, height=ph)
            ext = 3.0  # px: extend rules so neighbouring borders actually intersect
            for kind, x0, y0, x1, y1, thick in rules:
                if kind == "h":
                    x0, x1 = x0 - ext, x1 + ext
                else:
                    y0, y1 = y0 - ext, y1 + ext
                newp.draw_line((x0 * sx, y0 * sy), (x1 * sx, y1 * sy),
                               color=(0.45, 0.45, 0.45), width=max(0.5, min(2.0, thick * sx)))
            for ln in sorted(lines, key=lambda l: (round(l.y0 / 12), l.x0)):
                bottom = ln.y1 * sy
                baseline = bottom - (0.207 * ln.size if _DESC & set(ln.text) else 0)
                if ln.bullet_x is not None:
                    newp.insert_text((ln.bullet_x * sx, baseline), "•", fontname="helv",
                                     fontsize=ln.size)
                    stats["ocr_bullets"] += 1
                newp.insert_text((ln.x0 * sx, baseline), ln.text, fontname="helv", fontsize=ln.size)
            stats["ocr_pages"] += 1
            stats["ocr_lines"] += len(lines)
            stats["table_rules"] += len(rules)
        if progress:
            progress("ocr", stats["ocr_pages"], len(todo))
        out.save(dst, garbage=3, deflate=True)
    return stats


# ------------------------------------------------------------------- pdf2docx
class _PageProgress(logging.Handler):
    """Turns pdf2docx's '(i/n) Page k' log lines into progress callbacks."""

    PAT = re.compile(r"\((\d+)/(\d+)\)\s*Page")

    def __init__(self, progress: Progress, thread_id: int):
        super().__init__(logging.INFO)
        self.progress, self.thread_id, self.phase = progress, thread_id, "layout"

    def emit(self, record: logging.LogRecord) -> None:
        if record.thread != self.thread_id:
            return
        msg = record.getMessage()
        if "Creating pages" in msg:
            self.phase = "write"
        m = self.PAT.search(msg)
        if m:
            self.progress(self.phase, int(m.group(1)), int(m.group(2)))


def _pdf_to_docx(src: str | Path, dst: str | Path, progress: Optional[Progress]) -> None:
    from pdf2docx import Converter

    root = logging.getLogger()
    handler = None
    if progress:
        handler = _PageProgress(progress, threading.get_ident())
        root.addHandler(handler)
    old_level = root.level
    if root.level > logging.INFO:
        root.setLevel(logging.INFO)   # pdf2docx reports page progress at INFO
    try:
        cv = Converter(str(src))
        try:
            # Real page margins (pdf2docx halves them by default and pads the first paragraph
            # of every page, which turns into stray gaps once pages flow into one section).
            cv.convert(str(dst), page_margin_factor_top=1.0, page_margin_factor_bottom=1.0)
        finally:
            cv.close()
    finally:
        root.setLevel(old_level)
        if handler:
            root.removeHandler(handler)


# ------------------------------------------------------------------ enrichment
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
_BULLET = re.compile(r"^\s*([•·▪●◦‣⁃∙\-–—*]|o)[ \t ]+")
_NUMBER = re.compile(r"^\s*(\d{1,2}|[a-zA-Z]|[ivxIVX]{1,4})[.)][ \t ]+")
_MARKER = re.compile(_BULLET.pattern + "|" + _NUMBER.pattern)


@dataclass
class _Line:
    nodes: list            # run / hyperlink elements (in order)
    ppr: object            # w:pPr element of the source paragraph (may be None)
    first_in_para: bool
    gap_before: float      # pt of space before the source paragraph
    indent: float          # pt left indent of the source paragraph
    text: str = ""
    size: float = 0.0
    bold: bool = False


def _rpr_size(rpr, default: float) -> float:
    if rpr is not None:
        sz = rpr.find(W + "sz")
        if sz is not None and sz.get(W + "val"):
            return int(sz.get(W + "val")) / 2
    return default


def _rpr_bold(rpr) -> bool:
    if rpr is None:
        return False
    b = rpr.find(W + "b")
    return b is not None and b.get(W + "val", "true") not in ("0", "false")


def _node_text(node) -> str:
    return "".join((t.text or "") if t.tag == W + "t" else "\t" if t.tag == W + "tab" else ""
                   for t in node.iter() if t.tag in (W + "t", W + "tab"))


def _split_lines(p, default_size: float) -> list[_Line]:
    """Split a w:p at its soft line breaks into line units (copies of the runs)."""
    ppr = p.find(W + "pPr")
    gap, indent = 0.0, 0.0
    if ppr is not None:
        sp = ppr.find(W + "spacing")
        if sp is not None and sp.get(W + "before"):
            gap = int(sp.get(W + "before")) / 20
        ind = ppr.find(W + "ind")
        if ind is not None and ind.get(W + "left"):
            indent = int(ind.get(W + "left")) / 20
    lines, cur = [], []
    for child in p:
        if child.tag == W + "pPr":
            continue
        if child.tag != W + "r":
            cur.append(copy.deepcopy(child))
            continue
        rpr = child.find(W + "rPr")
        piece = []
        for el in child:
            if el.tag == W + "rPr":
                continue
            if el.tag == W + "br" and el.get(W + "type") in (None, "textWrapping"):
                if piece:
                    cur.append(_make_run(rpr, piece))
                    piece = []
                lines.append(cur)
                cur = []
            else:
                piece.append(copy.deepcopy(el))
        if piece:
            cur.append(_make_run(rpr, piece))
    lines.append(cur)
    out = []
    for i, nodes in enumerate(lines):
        if not nodes:
            continue
        ln = _Line(nodes, ppr, i == 0, gap if i == 0 else 0.0, indent)
        ln.text = "".join(_node_text(n) for n in nodes)
        weights: Counter = Counter()
        bold_chars = 0
        for n in nodes:
            t = _node_text(n).strip()
            if not t:
                continue
            rpr = n.find(W + "rPr")
            weights[_rpr_size(rpr, default_size)] += len(t)
            bold_chars += len(t) if _rpr_bold(rpr) else 0
        total = sum(weights.values())
        ln.size = weights.most_common(1)[0][0] if weights else default_size
        ln.bold = total > 0 and bold_chars / total > 0.5
        out.append(ln)
    return out


def _make_run(rpr, children):
    from docx.oxml import OxmlElement

    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    for c in children:
        r.append(c)
    return r


def _reflow(doc, default_size: float) -> int:
    """Join visually wrapped lines back into flowing paragraphs."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    body = doc.element.body
    merged_count = 0
    groups, cur = [], []
    for child in list(body):
        is_plain_p = (child.tag == W + "p" and child.find(f"{W}pPr/{W}sectPr") is None
                      and child.find(f".//{W}drawing") is None)
        is_empty = is_plain_p and not "".join(t.text or "" for t in child.iter(W + "t")).strip()
        if is_plain_p and not is_empty:
            cur.append(child)
        else:
            # tables, section breaks, drawings and empty spacer paragraphs (Word needs a
            # paragraph between two tables) separate groups and are left untouched
            groups.append(cur)
            cur = []
    groups.append(cur)

    for group in groups:
        plain = [p for p in group if p.find(f"{W}pPr/{W}sectPr") is None]
        if len(plain) < 1:
            continue
        lines: list[_Line] = []
        per_para = []
        for p in plain:
            ls = _split_lines(p, default_size)
            per_para.append(len(ls))
            lines.extend(ls)
        if not lines:
            continue
        full: dict[float, int] = {}
        for ln in lines:
            n = len(ln.text.strip())
            full[ln.size] = max(full.get(ln.size, 0), n)

        def joins(prev: _Line, cur_: _Line) -> bool:
            pt, ct = prev.text.strip(), cur_.text.strip()
            if not pt or not ct or _MARKER.match(cur_.text):
                return False
            if abs(prev.size - cur_.size) > 0.6 or prev.bold != cur_.bold:
                return False
            if cur_.first_in_para and (cur_.gap_before > max(3.0, 0.35 * cur_.size)
                                       or abs(cur_.indent - prev.indent) > 6):
                return False
            longest = max(1, full.get(prev.size, len(pt)))
            prev_full = len(pt) >= 0.72 * longest
            if pt[-1] in ".!?:" and not prev_full:
                return False
            return prev_full

        paragraphs: list[list[_Line]] = [[lines[0]]]
        for ln in lines[1:]:
            if joins(paragraphs[-1][-1], ln):
                paragraphs[-1].append(ln)
            else:
                paragraphs.append([ln])

        if all(c <= 1 for c in per_para) and all(len(p) == 1 for p in paragraphs):
            continue  # one line per paragraph before and after: nothing to change
        anchor = plain[0]
        parent = anchor.getparent()
        insert_at = parent.index(anchor)
        for p in plain:
            parent.remove(p)
        for k, para in enumerate(paragraphs):
            new_p = OxmlElement("w:p")
            if para[0].ppr is not None:
                new_p.append(copy.deepcopy(para[0].ppr))
            for j, ln in enumerate(para):
                nodes = ln.nodes
                if j > 0:
                    prev_text = para[j - 1].text.rstrip()
                    if not prev_text.endswith("-"):
                        sp = OxmlElement("w:r")
                        src_rpr = nodes[0].find(W + "rPr") if nodes else None
                        if src_rpr is not None:
                            sp.append(copy.deepcopy(src_rpr))
                        t = OxmlElement("w:t")
                        t.set(qn("xml:space"), "preserve")
                        t.text = " "
                        sp.append(t)
                        new_p.append(sp)
                    merged_count += 1
                for n in nodes:
                    new_p.append(n)
            parent.insert(insert_at + k, new_p)
    return merged_count


def _run_size(run, default: float) -> float:
    return run.font.size.pt if run.font.size else default


def enrich_docx(path: str | Path, title: str = "") -> dict:
    """Reflow paragraphs and add heading/list styles without changing the look."""
    from docx import Document
    from docx.shared import Pt, RGBColor

    doc = Document(str(path))
    normal = doc.styles["Normal"].font.size
    default_size = normal.pt if normal else 11.0

    joined = _reflow(doc, default_size)
    paras = [p for p in doc.paragraphs if p.text.strip()]

    sizes: Counter = Counter()
    for p in paras:
        for r in p.runs:
            if r.text.strip():
                sizes[round(_run_size(r, default_size) * 2) / 2] += len(r.text.strip())
    if not sizes:
        return {"lines_joined": joined, "headings": 0, "list_items": 0}
    body = sizes.most_common(1)[0][0]

    def freeze_look(p, keep_style_bold=False):
        """Pin run formatting so a style change can't alter colour/weight/size."""
        for r in p.runs:
            f = r.font
            if f.color is None or f.color.rgb is None:
                f.color.rgb = RGBColor(0, 0, 0)
            if f.bold is None and not keep_style_bold:
                f.bold = False
            if f.italic is None:
                f.italic = False
            if f.size is None:
                f.size = Pt(default_size)
        pf = p.paragraph_format
        if pf.space_before is None:
            pf.space_before = Pt(0)
        if pf.space_after is None:
            pf.space_after = Pt(0)

    # Headings: short lines set clearly larger (or bold and a bit larger) than body text.
    candidates = []
    for p in paras:
        text = p.text.strip()
        runs = [r for r in p.runs if r.text.strip()]
        if not runs or len(text) > 120 or len(text.split()) > 14:
            continue
        if text.endswith((".", ",", ";")) and len(text.split()) > 6:
            continue
        size = max(_run_size(r, default_size) for r in runs)
        bold = all(r.bold for r in runs)
        if size >= body * 1.25 or (bold and size >= body * 1.08):
            candidates.append((p, round(size * 2) / 2))
    levels = {s: i + 1 for i, s in enumerate(sorted({s for _, s in candidates}, reverse=True)[:3])}
    headings = 0
    for p, size in candidates:
        if size in levels:
            freeze_look(p, keep_style_bold=True)
            p.style = doc.styles[f"Heading {levels[size]}"]
            nxt = p._p.getnext()
            if nxt is not None and nxt.tag == W + "tbl":
                # With keep-with-next/keep-lines on a heading, Pages moves the whole following
                # table to the next page (leaving a gap): let the table start under its heading.
                p.paragraph_format.keep_with_next = False
                p.paragraph_format.keep_together = False
            headings += 1

    # Lists: bullet glyph or "1." / "a)" markers at the start of a paragraph.
    def strip_marker(p, pattern):
        consumed = len(pattern.match(p.text).group(0))
        for r in p.runs:
            if consumed <= 0:
                break
            take = min(len(r.text), consumed)
            r.text = r.text[take:]
            consumed -= take

    list_items = 0
    numbered_run: list = []

    def flush_numbered():
        nonlocal list_items
        if len(numbered_run) >= 2:
            for q in numbered_run:
                freeze_look(q)
                strip_marker(q, _NUMBER)
                q.style = doc.styles["List Number"]
                list_items += 1
        numbered_run.clear()

    for p in paras:
        if p.style.name.startswith("Heading"):
            flush_numbered()
            continue
        if _BULLET.match(p.text) and len(p.text.strip()) > 2:
            flush_numbered()
            freeze_look(p)
            strip_marker(p, _BULLET)
            p.style = doc.styles["List Bullet"]
            fli = p.paragraph_format.first_line_indent
            if fli is not None and fli > 0:      # keep hanging indents, drop first-line indents
                p.paragraph_format.first_line_indent = None
            list_items += 1
        elif _NUMBER.match(p.text):
            numbered_run.append(p)
        else:
            flush_numbered()
    flush_numbered()

    if not title:
        first = next((p for p in doc.paragraphs if p.style.name == "Heading 1"), None)
        title = first.text.strip() if first is not None else ""
    if title:
        doc.core_properties.title = title[:250]
    doc.core_properties.comments = "Converted from PDF"
    doc.save(str(path))
    return {"lines_joined": joined, "headings": headings, "list_items": list_items, "body_size": body}


# ---------------------------------------------------------------------- facade
def convert(src: str | Path, dst: str | Path, ocr: bool = False, lang: str = "en",
            progress: Optional[Progress] = None) -> dict:
    """Convert ``src`` (PDF) to ``dst`` (DOCX). Returns conversion statistics."""
    t0 = time.time()
    if progress:
        progress("analyze", 0, 1)
    info = analyze_pdf(src)
    if info.encrypted:
        raise ConversionError("This PDF is password-protected. Unlock it first, then convert.")
    if info.pages == 0:
        raise ConversionError("This PDF has no pages.")
    stats: dict = {"pages": info.pages, "scanned_pages": info.scanned_pages, "ocr": False}

    with tempfile.TemporaryDirectory(prefix="docxify-") as tmp:
        work = Path(src)
        if ocr and info.scanned_pages:
            work = Path(tmp) / "ocr.pdf"
            stats.update(build_ocr_pdf(src, work, info.scanned_pages, lang, progress))
            stats["ocr"] = True

        # Running header/footer -> real Word header/footer (removed from the body first).
        bands = None
        try:
            bands = layout.detect_bands(work, skip_pages=set() if stats["ocr"] else set(info.scanned_pages))
        except Exception:
            log.exception("header/footer detection failed; converting without it")
        body_pdf = work
        if bands is not None:
            body_pdf = Path(tmp) / "body.pdf"
            layout.strip_bands(work, body_pdf, bands)
            stats["header_footer"] = True

        rules = []
        try:
            rules = layout.body_rules(work, bands)
        except Exception:
            log.exception("rule detection failed")

        _pdf_to_docx(body_pdf, dst, progress)
        try:
            stats.update(docx_fix.postprocess(dst, bands, work, rules))
        except Exception:
            log.exception("docx post-processing failed; keeping raw layout")

    if progress:
        progress("enrich", 0, 1)
    try:
        stats.update(enrich_docx(dst, info.title))
    except Exception:  # semantics are a bonus; never fail the conversion for them
        log.exception("enrichment failed; keeping layout-only document")
    stats["seconds"] = round(time.time() - t0, 2)
    return stats
