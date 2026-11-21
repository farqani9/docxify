"""Post-processing that turns pdf2docx output into a well-behaved Word document.

pdf2docx reproduces each PDF page faithfully but rigidly: fonts named after the
PDF's embedded fonts, exact row heights and line spacing, one section (page
break) per PDF page, tables split at page boundaries, header/footer as body
text, and some invalid attribute values. This module fixes those, so the result
reflows and prints like a document authored in Word.
"""
from __future__ import annotations

import copy
import io
import re
import unicodedata
from typing import Optional

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from . import layout

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Free/embedded PDF fonts -> metric-compatible fonts that Word and Pages have.
_FONT_PREFIXES = [
    ("liberationsansnarrow", "Arial Narrow"), ("liberationsans", "Arial"),
    ("liberationserif", "Times New Roman"), ("liberationmono", "Courier New"),
    ("arimo", "Arial"), ("tinos", "Times New Roman"), ("cousine", "Courier New"),
    ("carlito", "Calibri"), ("caladea", "Cambria"),
    ("freesans", "Arial"), ("freeserif", "Times New Roman"), ("freemono", "Courier New"),
    ("nimbussansnarrow", "Arial Narrow"), ("nimbussans", "Arial"),
    ("nimbusroman", "Times New Roman"), ("nimbusmono", "Courier New"),
    ("texgyreheros", "Arial"), ("texgyretermes", "Times New Roman"), ("texgyrecursor", "Courier New"),
    ("dejavusansmono", "Courier New"), ("dejavusans", "Verdana"), ("dejavuserif", "Georgia"),
    ("helveticaneue", "Arial"), ("helvetica", "Arial"),
    ("arialnarrow", "Arial Narrow"), ("arial", "Arial"),
    ("timesnewroman", "Times New Roman"), ("times", "Times New Roman"),
    ("couriernew", "Courier New"), ("courier", "Courier New"),
    # Arabic: Naskh styles -> Times New Roman (Naskh Arabic glyphs), sans -> Arial
    ("notonaskharabic", "Times New Roman"), ("amiri", "Times New Roman"),
    ("scheherazade", "Times New Roman"), ("notosansarabic", "Arial"), ("notokufiarabic", "Arial"),
]

# Placeholder tokens written by the OCR step for right-to-left lines.
PLACEHOLDER = layout.PLACEHOLDER
_RTL_CHARS = re.compile(r"[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]")
_LTR_LETTERS = re.compile(r"[A-Za-z\u00C0-\u024F]")
_ARABIC_FORMS = re.compile(r"[\uFB50-\uFDFF\uFE70-\uFEFC]+")


def is_rtl(text: str) -> bool:
    """True when right-to-left letters (Arabic, Hebrew) dominate the text."""
    r = len(_RTL_CHARS.findall(text))
    return r > 0 and r >= len(_LTR_LETTERS.findall(text))


# Child order required by the OOXML schema (Word rejects out-of-order elements).
_SEQ = {
    "pPr": ("pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr "
            "suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct "
            "topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind "
            "contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment "
            "textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange").split(),
    "rPr": ("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint "
            "noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u "
            "effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath").split(),
    "tblPr": ("tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc "
              "tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption "
              "tblDescription tblPrChange").split(),
    "sectPr": ("headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc "
               "pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection "
               "bidi rtlGutter docGrid printerSettings sectPrChange").split(),
}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _ordered_child(parent, name: str):
    """Return parent's child ``w:name``, creating it at its schema position."""
    existing = parent.find(W + name)
    if existing is not None:
        return existing
    seq = _SEQ[_local(parent.tag)]
    rank = seq.index(name)
    el = OxmlElement(f"w:{name}")
    for i, child in enumerate(parent):
        ln = _local(child.tag)
        if ln in seq and seq.index(ln) > rank:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el
_BULLETS = "•·▪●◦‣⁃∙–-*"


def map_font(name: Optional[str]) -> Optional[str]:
    if not name:
        return name
    base = name.split("+", 1)[-1]
    key = re.sub(r"[^a-z]", "", base.lower())
    for prefix, target in _FONT_PREFIXES:
        if key.startswith(prefix):
            return target
    return base


def _hex(color) -> str:
    if isinstance(color, int):
        return f"{color:06X}"
    r, g, b = (max(0, min(255, round(c * 255))) for c in color[:3])
    return f"{r:02X}{g:02X}{b:02X}"


# ------------------------------------------------------------------ fixes
def normalize_fonts(doc) -> int:
    changed = 0
    roots = [doc.element, doc.styles.element]
    for root in roots:
        for rf in root.iter(W + "rFonts"):
            for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                v = rf.get(W + attr)
                if v:
                    m = map_font(v)
                    if m != v:
                        rf.set(W + attr, m)
                        changed += 1
    return changed


_FLOAT = re.compile(r"^-?\d+\.\d+$")


def _sort_children(parent) -> None:
    """Put a property container's children in schema order (stable for unknowns)."""
    seq = _SEQ[_local(parent.tag)]
    kids = list(parent)
    ranked = sorted(kids, key=lambda c: seq.index(_local(c.tag)) if _local(c.tag) in seq else len(seq))
    if ranked != kids:
        for c in kids:
            parent.remove(c)
        for c in ranked:
            parent.append(c)


def split_size_changes(doc) -> int:
    """pdf2docx can put a title, its subtitle and the body in one paragraph, lines
    separated by breaks. Where the text size jumps across a break (by a quarter or
    more), start a new paragraph so each part gets its own style."""
    made = 0

    def size(r):
        sz = r.find(f"{W}rPr/{W}sz")
        return int(sz.get(W + "val")) / 2 if sz is not None and sz.get(W + "val") else None

    for p in list(doc.element.body.iter(W + "p")):
        if p.getparent() is not doc.element.body:
            continue
        kids = [c for c in p if c.tag != W + "pPr"]
        rtl = is_rtl("".join(t.text or "" for t in p.iter(W + "t")))
        cut_after = []                         # break run ending a part, or ("before", k)
        last, pending = None, None
        for k, c in enumerate(kids):
            if c.tag != W + "r":
                continue
            if c.find(W + "br") is not None and not _run_text(c).strip():
                pending = k
                continue
            if _run_text(c).strip():
                s = size(c)
                jump = bool(s and last and max(s, last) / min(s, last) >= (1.15 if rtl else 1.25))
                if jump and pending is not None:
                    cut_after.append(pending)
                elif jump and rtl and k > 0:
                    # right-to-left lines arrive one run per line, without breaks
                    cut_after.append(("before", k))
                pending = None
                last = s or last
        if not cut_after:
            continue
        ppr = p.find(W + "pPr")
        groups, start = [], 0
        for k in cut_after:
            if isinstance(k, tuple):           # cut right before run k, nothing to drop
                groups.append(kids[start:k[1]])
                start = k[1]
            else:
                groups.append(kids[start:k])
                start = k + 1
        groups.append(kids[start:])
        for k in cut_after:
            if not isinstance(k, tuple):
                p.remove(kids[k])
        anchor = p
        for g in groups[1:]:
            np_ = OxmlElement("w:p")
            if ppr is not None:
                np_.append(copy.deepcopy(ppr))
            for c in g:
                np_.append(c)
            anchor.addnext(np_)
            anchor = np_
            made += 1
    return made


def readable_light_text(doc) -> int:
    """White (or near-white) text with no shading behind it, in its cell, paragraph or
    run, would be invisible on the page: make it black."""
    fixed = 0

    def shaded(el) -> bool:
        shd = el.find(W + "shd") if el is not None else None
        return shd is not None and (shd.get(W + "fill") or "auto").lower() not in ("auto", "ffffff")

    for r in doc.element.body.iter(W + "r"):
        rpr = r.find(W + "rPr")
        col = rpr.find(W + "color") if rpr is not None else None
        val = (col.get(W + "val") or "") if col is not None else ""
        if not re.fullmatch(r"[0-9A-Fa-f]{6}", val) or min(int(val[i:i + 2], 16) for i in (0, 2, 4)) < 230:
            continue
        p = r.getparent()
        while p is not None and p.tag != W + "p":
            p = p.getparent()
        tc = p.getparent() if p is not None else None
        while tc is not None and tc.tag != W + "tc":
            tc = tc.getparent() if tc.tag != W + "body" else None
        if shaded(rpr) or (p is not None and shaded(p.find(W + "pPr"))) or \
                (tc is not None and shaded(tc.find(W + "tcPr"))):
            continue
        col.set(W + "val", "000000")
        fixed += 1
    return fixed


def trim_edge_breaks(doc) -> int:
    """A plain line break at the very start or end of a paragraph only adds an empty
    line (left over where items or lines were split apart). Page breaks stay."""
    trimmed = 0

    def plain_br(el) -> bool:
        return el.tag == W + "br" and el.get(W + "type") in (None, "textWrapping")

    for p in doc.element.body.iter(W + "p"):
        for edge in ("end", "start"):
            while True:
                runs = [r for r in p if r.tag == W + "r"]
                if not runs:
                    break
                r = runs[-1] if edge == "end" else runs[0]
                kids = [k for k in r if k.tag != W + "rPr"]
                if not kids:                      # empty run: look past it
                    p.remove(r)
                    continue
                k = kids[-1] if edge == "end" else kids[0]
                if plain_br(k):
                    r.remove(k)
                    trimmed += 1
                    continue
                break
    return trimmed


def sort_properties(doc) -> None:
    """Schema order for every pPr/rPr/tblPr/sectPr (pdf2docx and later edits append
    children out of order; Pages then ignores some of them, e.g. pStyle or bidi)."""
    names = {W + n for n in _SEQ}
    for root in _parts(doc):
        for el in root.iter():
            if isinstance(el.tag, str) and el.tag in names:
                _sort_children(el)


def sanitize(doc) -> None:
    """Integer measurements, bare hex colours, complete shading (valid OOXML)."""
    for el in doc.element.iter():
        if not isinstance(el.tag, str):
            continue
        if el.tag == W + "tblPr":
            _sort_children(el)
        for k, v in list(el.attrib.items()):
            if _FLOAT.match(v):
                el.set(k, str(int(round(float(v)))))
            elif v.startswith("#") and k.endswith(("color", "fill")):
                el.set(k, v[1:].upper())
        if el.tag == W + "shd" and el.get(W + "val") is None:
            el.set(W + "val", "clear")
            el.set(W + "color", el.get(W + "color") or "auto")


def relax_heights(doc) -> dict:
    """Rows and lines may grow instead of clipping text."""
    rows = lines = 0
    for tr in doc.element.iter(W + "trHeight"):
        if tr.get(W + "hRule") == "exact":
            tr.set(W + "hRule", "atLeast")
            rows += 1
    for sp in doc.element.iter(W + "spacing"):
        if sp.get(W + "lineRule") == "exact":
            sp.set(W + "lineRule", "atLeast")
            lines += 1
    return {"rows_relaxed": rows, "lines_relaxed": lines}


def _cell_span(tc) -> int:
    gs = tc.find(f"{W}tcPr/{W}gridSpan")
    return int(gs.get(W + "val")) if gs is not None and gs.get(W + "val") else 1


def fix_table_grids(doc) -> int:
    """Make tblGrid agree with the per-cell widths (Word/Pages lay out from the grid)."""
    fixed = 0
    for tbl in doc.element.body.iter(W + "tbl"):
        grid = tbl.find(W + "tblGrid")
        if grid is None:
            continue
        cols = grid.findall(W + "gridCol")
        n = len(cols)
        widths = None
        rows = tbl.findall(W + "tr")
        for tr in rows:
            cells = tr.findall(W + "tc")
            ws = []
            for tc in cells:
                w = tc.find(f"{W}tcPr/{W}tcW")
                if w is None or _cell_span(tc) != 1 or not (w.get(W + "w") or "").isdigit():
                    ws = None
                    break
                ws.append(int(w.get(W + "w")))
            if ws and len(ws) == n:
                widths = ws
                break
        if widths is None and rows:   # spread spanned widths over their columns
            ws = []
            for tc in rows[0].findall(W + "tc"):
                w = tc.find(f"{W}tcPr/{W}tcW")
                span = _cell_span(tc)
                total = int(w.get(W + "w")) if w is not None and (w.get(W + "w") or "").isdigit() else 0
                ws.extend([total // span] * span)
            if len(ws) == n and all(ws):
                widths = ws
        if not widths:
            continue
        for col, w in zip(cols, widths):
            col.set(W + "w", str(w))
        tblpr = tbl.find(W + "tblPr")
        if tblpr is not None:
            tblw = tblpr.find(W + "tblW")
            if tblw is None:
                tblw = OxmlElement("w:tblW")
                tblpr.insert(0, tblw)
            tblw.set(W + "type", "dxa")
            tblw.set(W + "w", str(sum(widths)))
        fixed += 1
    return fixed


def _is_empty_p(el) -> bool:
    return (el is not None and el.tag == W + "p"
            and not "".join(t.text or "" for t in el.iter(W + "t")).strip()
            and el.find(f".//{W}drawing") is None and el.find(f".//{W}pict") is None)


def merge_sections(doc) -> list:
    """One continuous flow instead of a hard page break per PDF page.

    Returns the positions (element after which the page changed) so split tables
    can be rejoined."""
    body = doc.element.body
    sect_ps = [p for p in body.iterchildren(W + "p") if p.find(f"{W}pPr/{W}sectPr") is not None]
    final = body.find(W + "sectPr")
    if not sect_ps or final is None:
        return []
    sizes = {(s.find(W + "pgSz").get(W + "w"), s.find(W + "pgSz").get(W + "h"))
             for s in [p.find(f"{W}pPr/{W}sectPr") for p in sect_ps] + [final]
             if s.find(W + "pgSz") is not None}
    if len(sizes) > 1:
        return []      # mixed page sizes/orientations: keep the sections
    first = sect_ps[0].find(f"{W}pPr/{W}sectPr")
    fm, gm = first.find(W + "pgMar"), final.find(W + "pgMar")
    if fm is not None and gm is not None:
        for k, v in fm.attrib.items():
            gm.set(k, v)
    boundaries = []
    for p in sect_ps:
        sp = p.find(f"{W}pPr/{W}sectPr")
        sp.getparent().remove(sp)
        boundaries.append(p)
    return boundaries


def _row_text(tr) -> str:
    return re.sub(r"\s+", " ", "".join(t.text or "" for t in tr.iter(W + "t"))).strip()


def join_split_tables(doc, boundaries: list) -> int:
    """Rejoin a table that pdf2docx split at a page boundary; repeat its header row."""
    joined = 0
    for marker in boundaries:
        if marker.getparent() is None:
            continue
        prev = marker if marker.tag == W + "tbl" else marker.getprevious()
        while prev is not None and _is_empty_p(prev):
            prev = prev.getprevious()
        nxt = marker.getnext()
        while nxt is not None and _is_empty_p(nxt):
            nxt = nxt.getnext()
        if prev is None or nxt is None or prev.tag != W + "tbl" or nxt.tag != W + "tbl":
            continue
        ga = [int(c.get(W + "w") or 0) for c in prev.findall(f"{W}tblGrid/{W}gridCol")]
        gb = [int(c.get(W + "w") or 0) for c in nxt.findall(f"{W}tblGrid/{W}gridCol")]
        if len(ga) != len(gb) or not ga or any(abs(a - b) > max(60, 0.04 * a) for a, b in zip(ga, gb)):
            continue
        rows_a, rows_b = prev.findall(W + "tr"), nxt.findall(W + "tr")
        if not rows_a or not rows_b:
            continue
        if _row_text(rows_a[0]) and _row_text(rows_a[0]) == _row_text(rows_b[0]):
            nxt.remove(rows_b[0])
            rows_b = rows_b[1:]
            trpr = rows_a[0].find(W + "trPr")
            if trpr is None:
                trpr = OxmlElement("w:trPr")
                rows_a[0].insert(0, trpr)
            if trpr.find(W + "tblHeader") is None:
                trpr.append(OxmlElement("w:tblHeader"))
        for tr in rows_b:
            prev.append(tr)
        # drop the second table and the empty paragraphs between the two parts
        el = prev.getnext()
        while el is not None and el is not nxt:
            following = el.getnext()
            if _is_empty_p(el):
                el.getparent().remove(el)
            el = following
        nxt.getparent().remove(nxt)
        joined += 1
    # page-boundary paragraphs that are now empty and sit between blocks
    for marker in boundaries:
        if marker.getparent() is not None and marker.tag == W + "p" and _is_empty_p(marker):
            marker.getparent().remove(marker)
    return joined


def merge_header_rows(doc) -> int:
    """pdf2docx turns a filled, borderless header row into a separate one-row table,
    with column widths guessed from its text. When the next block is a table with as
    many columns over the same span, the row is that table's header: move it there,
    take the table's column widths, and repeat it on each page."""
    merged = 0
    body = doc.element.body
    for t1 in list(body.iter(W + "tbl")):
        if t1.getparent() is not body:
            continue
        rows = t1.findall(W + "tr")
        if not 1 <= len(rows) <= 3:
            continue
        # a coloured header band (a header whose cells wrap is read as one row per
        # text line); a plain one-row table is left alone
        def filled(tc):
            shd = tc.find(f"{W}tcPr/{W}shd")
            return shd is not None and (shd.get(W + "fill") or "auto").lower() not in ("auto", "ffffff")

        def has_text(tc):
            return bool("".join(t.text or "" for t in tc.iter(W + "t")).strip())

        # first row fully coloured; later rows coloured wherever they hold text
        shaded = all(filled(tc) for tc in rows[0].findall(W + "tc")) and \
            all(filled(tc) for tr in rows[1:] for tc in tr.findall(W + "tc") if has_text(tc))
        if not shaded or len({len(tr.findall(W + "tc")) for tr in rows}) != 1:
            continue
        nxt = t1.getnext()
        while nxt is not None and _is_empty_p(nxt):
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != W + "tbl":
            continue
        cells1 = rows[0].findall(W + "tc")
        g2 = [int(c.get(W + "w") or 0) for c in nxt.findall(f"{W}tblGrid/{W}gridCol")]
        if len(cells1) < 2 or len(cells1) != len(g2) or any(_cell_span(tc) != 1 for tc in cells1):
            continue

        def span(t):
            ind = t.find(f"{W}tblPr/{W}tblInd")
            x0 = int(ind.get(W + "w") or 0) if ind is not None else 0
            return x0, x0 + sum(int(c.get(W + "w") or 0) for c in t.findall(f"{W}tblGrid/{W}gridCol"))

        (a0, a1), (b0, b1) = span(t1), span(nxt)
        if abs(a0 - b0) > 200 or abs(a1 - b1) > 200:
            continue
        for tr in rows[1:]:              # fold the extra text lines into the first row
            for tc0, tc in zip(cells1, tr.findall(W + "tc")):
                for p in tc.findall(W + "p"):
                    if "".join(t.text or "" for t in p.iter(W + "t")).strip():
                        tc0.append(p)
            t1.remove(tr)
        for tc0 in cells1:               # drop the empty paragraphs the fold leaves first
            ps = tc0.findall(W + "p")
            filled = [p for p in ps if "".join(t.text or "" for t in p.iter(W + "t")).strip()]
            for p in ps:
                if filled and p not in filled:
                    tc0.remove(p)
        for tc, w in zip(cells1, g2):
            tcw = tc.find(f"{W}tcPr/{W}tcW")
            if tcw is not None:
                tcw.set(W + "w", str(w))
                tcw.set(W + "type", "dxa")
        trpr = rows[0].find(W + "trPr")
        if trpr is None:
            trpr = OxmlElement("w:trPr")
            rows[0].insert(0, trpr)
        if trpr.find(W + "tblHeader") is None:
            trpr.append(OxmlElement("w:tblHeader"))
        first = nxt.find(W + "tr")
        first.addprevious(rows[0])
        el = t1.getnext()
        while el is not None and el is not nxt:
            following = el.getnext()
            if _is_empty_p(el):
                body.remove(el)
            el = following
        body.remove(t1)
        merged += 1
    return merged


_INLINE_BULLETS = "\u2022\u25aa\u25cf\u25e6\u2023\u2043\u2219"   # not "\u00b7": it is also a separator


def _run_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(W + "t"))


def split_inline_bullets(doc) -> int:
    """pdf2docx sometimes packs a whole bulleted list into one paragraph
    ("item one\u2022item two<br/>\u2022item three"): give every item its own paragraph."""
    made = 0
    number = re.compile(r"\s*\d{1,3}[.)]\s")
    for p in list(doc.element.body.iter(W + "p")):
        children = [c for c in p if c.tag != W + "pPr"]
        cuts, seen = [], False
        whole = "".join(t.text or "" for t in p.iter(W + "t"))
        # a numbered item (paragraph or piece already cut off) holding the next item(s) too
        numbered = bool(number.match(whole))
        rtl = is_rtl(whole)
        in_body = p.getparent() is doc.element.body
        before, prev_rtl = "", False
        for idx, ch in enumerate(children):
            txt = _run_text(ch) if ch.tag in (W + "r", W + "hyperlink") else ""
            head = txt.strip()[:1]
            is_number = bool(number.match(txt))
            marker = head in _INLINE_BULLETS or is_number
            # right-to-left lines arrive one run per line: an item after a finished
            # sentence, or a left-to-right line after an Arabic one, starts a new paragraph
            after_sentence = rtl and before.rstrip().endswith((".", ":", "!", "?", "\u061f", "\u061b"))
            ltr_line = rtl and in_body and prev_rtl and bool(txt.strip()) and not _RTL_CHARS.search(txt) \
                and bool(re.search(r"[A-Za-z0-9]", txt))
            if seen and head and (head in _INLINE_BULLETS or (numbered and is_number)
                                  or (marker and after_sentence) or ltr_line):
                cuts.append(idx)
                numbered = is_number
            if txt.strip():
                seen = True
                before += txt
                prev_rtl = bool(_RTL_CHARS.search(txt))
        if not cuts:
            continue
        bounds = [0] + cuts + [len(children)]
        segments = [children[a:b] for a, b in zip(bounds, bounds[1:])]

        def trim(seg):  # drop line breaks at the edges of an item (whole runs or trailing <w:br/>)
            is_br = lambda c: c.tag == W + "r" and c.find(W + "br") is not None and not _run_text(c).strip()
            while seg and is_br(seg[-1]):
                seg.pop()
            while seg and is_br(seg[0]):
                seg.pop(0)
            for edge in (seg[-1:], seg[:1]):
                for c in edge:
                    if c.tag == W + "r":
                        kids = [k for k in c if k.tag != W + "rPr"]
                        while kids and kids[-1].tag == W + "br":
                            c.remove(kids.pop())
                        while kids and kids[0].tag == W + "br":
                            c.remove(kids.pop(0))
            return seg

        ppr = p.find(W + "pPr")
        after = None
        if ppr is not None and ppr.find(W + "spacing") is not None:
            after = ppr.find(W + "spacing").get(W + "after")
            if after:
                ppr.find(W + "spacing").set(W + "after", "0")   # the gap follows the last item
        for ch in children:
            p.remove(ch)
        for ch in trim(segments[0]):
            p.append(ch)
        anchor = p
        for k, seg in enumerate(segments[1:]):
            np_ = OxmlElement("w:p")
            if ppr is not None:
                ppr2 = copy.deepcopy(ppr)
                sp = ppr2.find(W + "spacing")
                if sp is not None and sp.get(W + "before"):
                    sp.set(W + "before", "20")   # the gap belonged before the first item only
                if sp is not None and after and k == len(segments) - 2:
                    sp.set(W + "after", after)
                np_.append(ppr2)
            for ch in trim(seg):
                np_.append(ch)
            anchor.addnext(np_)
            anchor = np_
            made += 1
    return made


def fix_bullets(doc) -> int:
    """Glyph + tab + hanging indent; drop wrap-continuation tabs inside paragraphs."""
    fixed = 0
    for p in list(doc.element.body.iter(W + "p")):
        if is_rtl("".join(t.text or "" for t in p.iter(W + "t"))):
            continue   # its glyph/tab layout is left-to-right; RTL items keep "• text"
        runs = [r for r in p.iterchildren(W + "r")]
        texts = [(r, r.find(W + "t")) for r in runs]
        first = next(((r, t) for r, t in texts if t is not None and (t.text or "").strip()), None)
        bullet_tab = None
        if first is not None:
            r0, t0 = first
            s = (t0.text or "").lstrip()
            is_bullet = s[:1] in _BULLETS and (len(s) == 1 or s[1] in " \t ") and \
                len("".join(t.text or "" for _, t in texts if t is not None).strip()) > 2
            if s[:1] in "-*–" and len(s) > 1 and s[1] not in " \t ":
                is_bullet = False
            if is_bullet:
                glyph, rest = s[0], s[1:].lstrip("  ")
                t0.text = glyph
                insert_at = list(p).index(r0) + 1
                nxt_run = r0.getnext()
                has_tab = nxt_run is not None and nxt_run.tag == W + "r" and nxt_run.find(W + "tab") is not None \
                    and nxt_run.find(W + "t") is None
                if rest:
                    rr = OxmlElement("w:r")
                    if r0.find(W + "rPr") is not None:
                        rr.append(copy.deepcopy(r0.find(W + "rPr")))
                    tt = OxmlElement("w:t")
                    tt.set(qn("xml:space"), "preserve")
                    tt.text = rest
                    rr.append(tt)
                    p.insert(insert_at, rr)
                if not has_tab:
                    tr = OxmlElement("w:r")
                    tr.append(OxmlElement("w:tab"))
                    p.insert(insert_at, tr)
                    bullet_tab = tr
                else:
                    bullet_tab = nxt_run
                ppr = p.find(W + "pPr")
                if ppr is None:
                    ppr = OxmlElement("w:pPr")
                    p.insert(0, ppr)
                ind = _ordered_child(ppr, "ind")
                left = int(float(ind.get(W + "left") or ind.get(W + "start") or 0))
                stops = [int(float(t.get(W + "pos"))) for t in ppr.iter(W + "tab") if t.get(W + "pos")]
                text_pos = min([s for s in stops if s > left + 60], default=left + 230)
                for a in ("start", "firstLine", "hanging"):
                    if ind.get(W + a) is not None:
                        del ind.attrib[W + a]
                ind.set(W + "left", str(text_pos))
                ind.set(W + "hanging", str(text_pos - left))
                fixed += 1
        # soft line breaks: in a bullet item they are wraps; elsewhere only when the next
        # line continues the sentence (starts lower-case and the previous line has no stop)
        for br in list(p.iter(W + "br")):
            if br.get(W + "type") not in (None, "textWrapping"):
                continue
            run = br.getparent()
            before = "".join(t.text or "" for t in _preceding_text(p, run))
            after = "".join(t.text or "" for t in _following_text(p, run))
            wrap = bullet_tab is not None or (
                after[:1].islower() and before.rstrip()[-1:] not in ".:;!?")
            if wrap and before.strip() and after.strip():
                sp = OxmlElement("w:t")
                sp.set(qn("xml:space"), "preserve")
                sp.text = "" if before.endswith((" ", "-")) else " "
                run.replace(br, sp)
        # continuation tabs: a tab right after text that already ends with a space
        acc = ""
        for r in list(p.iterchildren(W + "r")):
            t = r.find(W + "t")
            if t is not None:
                acc += t.text or ""
            elif r.find(W + "tab") is not None and r is not bullet_tab and acc and acc[-1] in "  ":
                p.remove(r)
    return fixed


def _preceding_text(p, run):
    out = []
    for r in p.iter(W + "r"):
        if r is run:
            for c in r:
                if c.tag == W + "br":
                    break
                if c.tag == W + "t":
                    out.append(c)
            break
        out.extend(r.findall(W + "t"))
    return out


def _following_text(p, run):
    out, found = [], False
    for r in p.iter(W + "r"):
        if r is run:
            found, after_br = True, False
            for c in r:
                if c.tag == W + "br":
                    after_br = True
                elif after_br and c.tag == W + "t":
                    out.append(c)
            continue
        if found:
            out.extend(r.findall(W + "t"))
            if out and "".join(t.text or "" for t in out).strip():
                break
    return out


def zero_cell_vertical_margins(doc) -> int:
    """Word's default top/bottom cell margin is 0, Pages adds its own unless told."""
    n = 0
    for tblpr in doc.element.body.iter(W + "tblPr"):
        mar = _ordered_child(tblpr, "tblCellMar")
        for side in ("top", "bottom"):
            if mar.find(W + side) is None:
                el = OxmlElement(f"w:{side}")
                el.set(W + "w", "0")
                el.set(W + "type", "dxa")
                mar.append(el)
        n += 1
    return n


def fix_cell_indents(doc) -> int:
    """pdf2docx gives cell paragraphs a right indent equal to where each PDF line
    happened to end; in Word that wraps the last word early. Cells already pad."""
    n = 0
    for tc in doc.element.body.iter(W + "tc"):
        for ind in tc.iter(W + "ind"):
            for a in ("right", "end"):
                if ind.get(W + a) not in (None, "0"):
                    ind.set(W + a, "0")
                    n += 1
    return n


def restore_letter_spacing(doc) -> int:
    """Letter-spaced text ("E D U C A T I O N") arrives as one run per letter with
    the spacing dropped; give those runs character spacing again."""
    fixed = 0
    for p in doc.element.body.iter(W + "p"):
        runs = [r for r in p.iterchildren(W + "r") if r.find(W + "t") is not None]
        texts = [r.find(W + "t").text or "" for r in runs]
        singles = [r for r, t in zip(runs, texts) if len(t.strip()) == 1]
        if len(runs) < 8 or len(singles) / len(runs) < 0.7 or is_rtl("".join(texts)):
            continue
        for r in runs:
            rpr = r.find(W + "rPr")
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            sz = rpr.find(W + "sz")
            size_pt = int(sz.get(W + "val")) / 2 if sz is not None and (sz.get(W + "val") or "").isdigit() else 10
            sp = _ordered_child(rpr, "spacing")
            sp.set(W + "val", str(round(size_pt * 0.3 * 20)))   # twentieths of a point
        fixed += 1
    return fixed


# ------------------------------------------------------------------ header / footer
def _set_border(paragraph, side: str, rule: layout.Item, gap_pt: float) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    bdr = _ordered_child(ppr, "pBdr")
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(max(2, min(96, round(rule.width * 8)))))
    el.set(qn("w:space"), str(max(0, min(31, round(gap_pt)))))
    el.set(qn("w:color"), _hex(rule.color))
    bdr.append(el)


def _field(paragraph, instr: str, text: str, span: layout.Span) -> None:
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f" {instr} ")
    r = OxmlElement("w:r")
    r.append(_rpr(span))
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)


def _rpr(span: layout.Span):
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    f = map_font(span.font)
    for a in ("ascii", "hAnsi", "cs"):
        fonts.set(qn(f"w:{a}"), f)
    rpr.append(fonts)
    if span.bold:
        rpr.append(OxmlElement("w:b"))
    if span.italic:
        rpr.append(OxmlElement("w:i"))
    col = OxmlElement("w:color")
    col.set(qn("w:val"), _hex(span.color))
    rpr.append(col)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(round(span.size * 2)))
    rpr.append(sz)
    return rpr


def _text_run(paragraph, text: str, span: layout.Span) -> None:
    r = OxmlElement("w:r")
    r.append(_rpr(span))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    paragraph._p.append(r)


def _build_band(container, items: list, page_w: float, left: float, right: float,
                pdf_path, is_header: bool) -> None:
    container.is_linked_to_previous = False
    content = [i for i in items if i.kind != "rule"]
    rules = [i for i in items if i.kind == "rule"]
    if not content:
        return
    text_w = page_w - left - right
    rows: list[list] = []
    for it in sorted(content, key=lambda i: i.cy):
        if rows and abs(rows[-1][0].cy - it.cy) <= max(6.0, 0.4 * (it.bbox[3] - it.bbox[1])):
            rows[-1].append(it)
        else:
            rows.append([it])
    paragraphs = []
    for ri, row in enumerate(rows):
        p = container.paragraphs[0] if ri == 0 else container.add_paragraph()
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        pf = p.paragraph_format
        pf.space_before = pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        row.sort(key=lambda i: i.bbox[0])
        segs: list[list] = []
        for it in row:
            if segs and it.bbox[0] - segs[-1][-1].bbox[2] < 24:
                segs[-1].append(it)
            else:
                segs.append([it])
        aligns = []
        for seg in segs:
            x0, x1 = seg[0].bbox[0], seg[-1].bbox[2]
            if x0 - left < 0.25 * text_w:
                aligns.append("left")
            elif (page_w - right) - x1 < 0.25 * text_w:
                aligns.append("right")
            else:
                aligns.append("center")
        ours = []
        if "center" in aligns:
            ours.append(round(text_w / 2))
        if "right" in aligns:
            ours.append(round(text_w))
        for ts in p.style.paragraph_format.tab_stops:   # e.g. Header style: 3.25" centre, 6.5" right
            if round(ts.position.pt) not in ours:
                pf.tab_stops.add_tab_stop(ts.position, WD_TAB_ALIGNMENT.CLEAR)
        if "center" in aligns:
            pf.tab_stops.add_tab_stop(Pt(text_w / 2), WD_TAB_ALIGNMENT.CENTER)
        if "right" in aligns:
            pf.tab_stops.add_tab_stop(Pt(text_w), WD_TAB_ALIGNMENT.RIGHT)
        for seg, align in zip(segs, aligns):
            if align != "left":
                r = OxmlElement("w:r")
                r.append(OxmlElement("w:tab"))
                p._p.append(r)
            for it in seg:
                if it.kind == "image":
                    png = layout.image_png(pdf_path, it.xref)
                    if png:
                        p.add_run().add_picture(io.BytesIO(png), width=Pt(it.bbox[2] - it.bbox[0]))
                    continue
                if layout._is_rtl("".join(s.text for s in it.spans)):
                    main = max(it.spans, key=lambda s: len(s.text.strip()))
                    text = it.logical or layout.logical_line_text(
                        [{"bbox": s.bbox, "text": s.text} for s in it.spans])
                    # numbers are matched by value: reading order differs from extraction
                    values = [int(x) for x in re.findall(r"\d+", "".join(s.text for s in it.spans))]
                    pool = [(v, r) for v, r in zip(values, it.digit_roles) if r]
                    if not pool:                 # one run keeps the line's bidi order intact
                        _text_run(p, text, main)
                        continue
                    plain = ""
                    for tok in re.split(r"(\d+)", text):
                        hit = next((pr for pr in pool if tok.isdigit() and pr[0] == int(tok)), None)
                        if not hit:
                            plain += tok
                            continue
                        if plain:
                            _text_run(p, plain, main)
                            plain = ""
                        pool.remove(hit)
                        _field(p, hit[1], tok, main)
                        for rpr in p._p[-1].iter(W + "rPr"):    # page number inside RTL text
                            _ordered_child(rpr, "rtl")
                    if plain:
                        _text_run(p, plain, main)
                    continue
                roles = list(it.digit_roles)
                k = 0
                for span in it.spans:
                    for tok in re.split(r"(\d+)", span.text):
                        if not tok:
                            continue
                        if tok.isdigit():
                            role = roles[k] if k < len(roles) else None
                            k += 1
                            if role:
                                _field(p, role, tok, span)
                                continue
                        _text_run(p, tok, span)
        paragraphs.append((p, row))
    # rule lines -> paragraph borders
    top_p, top_row = paragraphs[0]
    bot_p, bot_row = paragraphs[-1]
    for rule in rules:
        ry = (rule.bbox[1] + rule.bbox[3]) / 2
        if ry <= min(i.bbox[1] for i in top_row):
            _set_border(top_p, "top", rule, min(i.bbox[1] for i in top_row) - ry)
        elif ry >= max(i.bbox[3] for i in bot_row):
            _set_border(bot_p, "bottom", rule, ry - max(i.bbox[3] for i in bot_row))


def add_header_footer(doc, bands: layout.Bands, pdf_path) -> dict:
    page_w, page_h = bands.page_size
    added = {"header": False, "footer": False}
    for sec in doc.sections:
        left, right = sec.left_margin.pt, sec.right_margin.pt
        head = [i for i in bands.header if i.kind != "rule"]
        foot = [i for i in bands.footer if i.kind != "rule"]
        if head:
            sec.header_distance = Pt(max(8.0, min(i.bbox[1] for i in head)))
            _build_band(sec.header, bands.header, page_w, left, right, pdf_path, True)
            added["header"] = True
        if foot:
            sec.footer_distance = Pt(max(8.0, page_h - max(i.bbox[3] for i in foot)))
            _build_band(sec.footer, bands.footer, page_w, left, right, pdf_path, False)
            added["footer"] = True
        break   # one merged section; later sections link to it
    return added


def apply_body_rules(doc, rules: list, placeholders: Optional[dict] = None) -> int:
    """Rules drawn under a paragraph in the PDF become that paragraph's bottom border."""
    def norm(t: str) -> str:   # same form on both sides: forms unshaped, no harakat
        t = unicodedata.normalize("NFKC", t)
        t = re.sub(r"[\u064B-\u065F\u0670\u0640]", "", t)
        return re.sub(r"\s+", " ", t).strip()

    def untoken(t: str) -> str:
        if not placeholders:
            return t
        return PLACEHOLDER.sub(lambda m: (placeholders.get(layout.placeholder_key(m.group(0))) or {}).get("text", ""), t)

    paras = [p for p in doc.paragraphs if p.text.strip()]
    applied, start = 0, 0
    for rule in rules:
        target = norm(untoken(rule["text_above"]))
        if len(target) < 3:
            continue
        for k in range(start, len(paras)):
            if norm(paras[k].text).endswith(target) or target in norm(paras[k].text):
                item = layout.Item("rule", rule["bbox"], color=rule["color"], width=rule["width"])
                _set_border(paras[k], "bottom", item, rule["gap"])
                # the PDF's distance to the next block already includes the rule's area
                pf = paras[k].paragraph_format
                if pf.space_after is not None:
                    pf.space_after = Pt(max(0.0, pf.space_after.pt - min(31, rule["gap"]) - rule["width"]))
                start = k + 1
                applied += 1
                break
    return applied


# ------------------------------------------------------------------ right-to-left
def _parts(doc) -> list:
    """Body plus every header/footer part that exists."""
    roots = [doc.element.body]
    for sec in doc.sections:
        for hf in (sec.header, sec.footer):
            if not hf.is_linked_to_previous:
                roots.append(hf._element)
    return roots


def break_rtl_lines(doc, mapping: dict) -> int:
    """pdf2docx decides where a line ends by whether it reaches the right edge (a
    left-to-right assumption); right-to-left lines are flush right, so every line looks
    wrapped and a whole page can land in one paragraph without breaks. Each placeholder
    is one source line: in body paragraphs, a line that starts a list item or changes
    the text size gets a line break before it, so the reflow keeps it apart. Plain
    continuation lines stay joined, as pdf2docx joined them."""
    number = re.compile(r"\s*\d{1,3}[.)]\s")
    made = 0

    def info(txt: str) -> dict:
        e = mapping.get(layout.placeholder_key(txt)) or {}
        return e if isinstance(e, dict) else {"text": e}

    for p in doc.element.body.iterchildren(W + "p"):
        runs = [r for r in p if r.tag == W + "r"]
        prev_token, between, prev_size = False, [], None
        for r in runs:
            txt = "".join(t.text or "" for t in r.findall(W + "t"))
            only_space = not txt.strip() and r.find(W + "br") is None and r.find(W + "drawing") is None
            if PLACEHOLDER.fullmatch(txt.strip() or "-"):
                e = info(txt)
                line, size = e.get("text", ""), e.get("size")
                starts_item = line.strip()[:1] in _INLINE_BULLETS or bool(number.match(line))
                size_jump = bool(size and prev_size and max(size, prev_size) / min(size, prev_size) >= 1.15)
                if prev_token and (starts_item or size_jump) \
                        and not any(b.find(W + "br") is not None for b in between):
                    for b in between:             # tabs/spaces pdf2docx put between the lines
                        if b.getparent() is not None:
                            p.remove(b)
                    br = OxmlElement("w:r")
                    br.append(OxmlElement("w:br"))
                    r.addprevious(br)
                    made += 1
                prev_token, between, prev_size = True, [], size or prev_size
            elif only_space or r.find(W + "tab") is not None and not txt.strip():
                between.append(r)
            else:
                prev_token, between = False, []
    return made


_REVIEW_INK = re.compile(r"^(0100|FEFF)([0-9A-Fa-f]{2})$")


def _highlight(r) -> None:
    rpr = r.find(W + "rPr")
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r.insert(0, rpr)
    _ordered_child(rpr, "highlight").set(W + "val", "yellow")


def _highlight_words(r, words: list) -> int:
    """Split run ``r`` so that each of ``words`` (first occurrence) is its own highlighted run."""
    t = r.find(W + "t")
    text = (t.text or "") if t is not None else ""
    spans, pos = [], 0
    for wd in words:
        k = text.find(wd, pos)
        if k < 0:
            k = text.find(wd)
        if k >= 0 and wd:
            spans.append((k, k + len(wd)))
            pos = k + len(wd)
    if not spans:
        return 0
    spans = sorted(set(spans))
    pieces, cur = [], 0
    for a, b in spans:
        if a < cur:
            continue
        if a > cur:
            pieces.append((text[cur:a], False))
        pieces.append((text[a:b], True))
        cur = b
    if cur < len(text):
        pieces.append((text[cur:], False))
    rpr = r.find(W + "rPr")
    for piece, marked in pieces:
        nr = OxmlElement("w:r")
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        nt = OxmlElement("w:t")
        nt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        nt.text = piece
        nr.append(nt)
        if marked:
            _highlight(nr)
        r.addprevious(nr)
    r.getparent().remove(r)
    return sum(1 for _p, m in pieces if m)


def mark_review(doc, mapping: dict) -> int:
    """Highlight text the OCR read with low confidence. Left-to-right lines come
    with a marker colour (the confidence in the blue channel), which is replaced by
    the real colour; right-to-left lines were flagged in their placeholder record."""
    marks = 0
    for r in list(doc.element.body.iter(W + "r")):
        col = r.find(f"{W}rPr/{W}color")
        m = _REVIEW_INK.match((col.get(W + "val") or "") if col is not None else "")
        if not m:
            continue
        col.set(W + "val", "000000" if m.group(1) == "0100" else "FFFFFF")
        if _run_text(r).strip():
            _highlight(r)
            marks += 1
    for r in list(doc.element.body.iter(W + "r")):
        info = r.attrib.pop("{urn:docxify}review", None)
        if info is None:
            continue
        e = mapping.get(info) or {}
        words = (e.get("review") or {}).get("words") or []
        if words:
            marks += _highlight_words(r, words)
        else:
            _highlight(r)
            marks += 1
    return marks


def substitute_placeholders(doc, mapping: dict) -> int:
    """Swap placeholder tokens for the logical-order text, with the intended font
    size and weight (the token itself was sized only to cover the right width)."""
    if not mapping:
        return 0

    def entry(tok: str) -> dict:
        e = mapping.get(layout.placeholder_key(tok), {"text": ""})
        return e if isinstance(e, dict) else {"text": e}

    def style_run(r, e: dict) -> None:
        rpr = r.find(W + "rPr")
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        if e.get("size"):
            v = str(round(float(e["size"]) * 2))
            _ordered_child(rpr, "sz").set(W + "val", v)
            _ordered_child(rpr, "szCs").set(W + "val", v)
        if "bold" in e:
            for tag in ("b", "bCs"):
                el = _ordered_child(rpr, tag)
                if not e["bold"]:
                    el.set(W + "val", "0")
        if map_font(e.get("font") or "") == "Times New Roman":
            # a serif (Naskh-style) source: the token itself was set in Helvetica
            fonts = _ordered_child(rpr, "rFonts")
            for a in ("ascii", "hAnsi", "cs"):
                fonts.set(W + a, "Times New Roman")

    n = 0
    for root in _parts(doc):
        for t in list(root.iter(W + "t")):
            if t.text and PLACEHOLDER.search(t.text):
                toks = PLACEHOLDER.findall(t.text)
                t.text = PLACEHOLDER.sub(lambda m: entry(m.group(0))["text"], t.text)
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                run = t.getparent()
                if run is not None and run.tag == W + "r":
                    style_run(run, entry(toks[0]))
                    if len(toks) == 1 and entry(toks[0]).get("review"):
                        run.set("{urn:docxify}review", layout.placeholder_key(toks[0]))
                n += 1
        for p in root.iter(W + "p"):   # a token split across runs
            ts = list(p.iter(W + "t"))
            full = "".join(x.text or "" for x in ts)
            if PLACEHOLDER.search(full):
                tok = PLACEHOLDER.search(full).group(0)
                ts[0].text = PLACEHOLDER.sub(lambda m: entry(m.group(0))["text"], full)
                for x in ts[1:]:
                    x.text = ""
                if ts[0].getparent() is not None and ts[0].getparent().tag == W + "r":
                    style_run(ts[0].getparent(), entry(tok))
                n += 1
        # safety net: pieces of a token split across cells or paragraphs
        for t in root.iter(W + "t"):
            if t.text and (re.search(r"QXR\d+Z*", t.text) or re.fullmatch(r"\s*Z{2,}\s*", t.text)):
                t.text = re.sub(r"QXR\d+Z*", "", t.text)
                if re.fullmatch(r"\s*Z{2,}\s*", t.text):
                    t.text = ""
    return n


def normalize_arabic_forms(doc) -> int:
    """Presentation forms (shaped glyph code points from PDF text layers) -> letters."""
    n = 0
    for root in _parts(doc):
        for t in root.iter(W + "t"):
            if t.text and _ARABIC_FORMS.search(t.text):
                t.text = _ARABIC_FORMS.sub(lambda m: unicodedata.normalize("NFKC", m.group(0)), t.text)
                n += 1
    return n


def fix_rtl(doc) -> dict:
    """Mark right-to-left paragraphs and runs the way Word expects for Arabic."""
    paras = rtl = 0
    for root in _parts(doc):
        in_band = root is not doc.element.body
        for p in root.iter(W + "p"):
            text = "".join(t.text or "" for t in p.iter(W + "t"))
            if not text.strip():
                continue
            paras += 1
            if not is_rtl(text):
                continue
            rtl += 1
            ppr = p.find(W + "pPr")
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                p.insert(0, ppr)
            if in_band and (p.find(f".//{W}tab") is not None or p.find(f".//{W}ptab") is not None):
                # header/footer items placed with tab stops measured from the left:
                # a right-to-left paragraph would mirror them. Only the runs are RTL.
                _mark_rtl_runs(p)
                continue
            _ordered_child(ppr, "bidi")
            jc = ppr.find(W + "jc")
            if jc is not None and jc.get(W + "val") in ("left", "right", "start", "end"):
                ppr.remove(jc)            # bidi paragraphs default to right alignment
            ind = ppr.find(W + "ind")
            if ind is not None:           # PDF left offsets are meaningless once mirrored
                ppr.remove(ind)
            if not in_band:
                _drop_layout_tabs(p, ppr)
            _mark_rtl_runs(p)
    if paras and rtl / paras >= 0.5:      # mostly RTL: right-to-left page layout
        for sect in doc.element.body.iter(W + "sectPr"):
            _ordered_child(sect, "bidi")
    return {"rtl_paragraphs": rtl}


def _drop_layout_tabs(p, ppr) -> None:
    """pdf2docx places the short last line of a right-to-left paragraph with a tab
    (it measures from the left); in flowing RTL text that tab is just a word gap."""
    tabs = ppr.find(W + "tabs")
    if tabs is not None:
        ppr.remove(tabs)
    runs = list(p.iter(W + "r"))
    for k, r in enumerate(runs):
        tab = r.find(W + "tab")
        if tab is None or _run_text(r).strip():
            continue
        before = "".join(_run_text(x) for x in runs[:k])
        after = "".join(_run_text(x) for x in runs[k + 1:])
        if not before.strip() or not after.strip() or before.endswith((" ", "\u00a0")) or after.startswith(" "):
            r.remove(tab)                # leading/trailing tab, or a space is already there
        else:
            t = OxmlElement("w:t")
            t.text = " "
            t.set(qn("xml:space"), "preserve")
            tab.addprevious(t)
            r.remove(tab)


def _mark_rtl_runs(p) -> None:
    """Right-to-left run marks: complex-script font, size, weight, language."""
    for r in p.iter(W + "r"):
        rt = "".join(t.text or "" for t in r.findall(W + "t"))
        if not rt:
            continue
        rpr = r.find(W + "rPr")
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        fonts = _ordered_child(rpr, "rFonts")
        base = map_font(fonts.get(W + "ascii") or "") or ""
        fonts.set(W + "cs", base if base in ("Arial", "Times New Roman") else "Arial")
        sz = rpr.find(W + "sz")
        if sz is not None and sz.get(W + "val"):
            _ordered_child(rpr, "szCs").set(W + "val", sz.get(W + "val"))
        b = rpr.find(W + "b")
        if b is not None and b.get(W + "val", "true") not in ("0", "false"):
            _ordered_child(rpr, "bCs")
        if _RTL_CHARS.search(rt):
            _ordered_child(rpr, "rtl")
            _ordered_child(rpr, "lang").set(W + "bidi", "ar-SA")


# ------------------------------------------------------------------ facade
def postprocess(path, bands: Optional[layout.Bands] = None, source_pdf=None, body_rules=None,
                placeholders: Optional[dict] = None, ocr: bool = False, review: bool = False) -> dict:
    doc = Document(str(path))
    stats = {"fonts_mapped": normalize_fonts(doc)}
    sanitize(doc)
    # real text first: later steps look at it (bullets, repeated table header rows)
    if placeholders:
        stats["rtl_line_breaks"] = break_rtl_lines(doc, placeholders)
    stats["rtl_lines_restored"] = substitute_placeholders(doc, placeholders or {})
    normalize_arabic_forms(doc)
    if review:
        stats["ocr_review_marks"] = mark_review(doc, placeholders or {})
    stats.update(relax_heights(doc))
    stats["table_grids_fixed"] = fix_table_grids(doc)
    boundaries = merge_sections(doc)
    stats["sections_merged"] = len(boundaries)
    stats["header_rows_merged"] = merge_header_rows(doc)
    stats["tables_joined"] = join_split_tables(doc, boundaries)
    stats["cell_indents_fixed"] = fix_cell_indents(doc)
    zero_cell_vertical_margins(doc)
    stats["letter_spaced"] = restore_letter_spacing(doc)
    stats["size_breaks_split"] = split_size_changes(doc)
    stats["bullet_items_split"] = split_inline_bullets(doc)
    stats["bullets_fixed"] = fix_bullets(doc)
    if body_rules:
        stats["rules_as_borders"] = apply_body_rules(doc, body_rules, placeholders or {})
    if bands is not None and source_pdf is not None:
        stats.update({f"{k}_rebuilt": v for k, v in add_header_footer(doc, bands, source_pdf).items()})
        # header/footer text comes from the source PDF: on OCR pages that holds tokens too
        stats["rtl_lines_restored"] += substitute_placeholders(doc, placeholders or {})
        normalize_arabic_forms(doc)
    # header/footer runs were rebuilt from the source text: no placeholders there
    stats.update(fix_rtl(doc))
    stats["edge_breaks_trimmed"] = trim_edge_breaks(doc)
    if ocr:   # rebuilt pages: white text belongs on a fill; digital PDFs can put it on pictures
        stats["light_text_fixed"] = readable_light_text(doc)
    doc.save(str(path))
    return stats
