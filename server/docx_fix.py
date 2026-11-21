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
]
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


def sanitize(doc) -> None:
    """Integer measurements, bare hex colours, complete shading (valid OOXML)."""
    for el in doc.element.iter():
        if not isinstance(el.tag, str):
            continue
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


_INLINE_BULLETS = "\u2022\u25aa\u25cf\u25e6\u2023\u2043\u2219"   # not "\u00b7": it is also a separator


def _run_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(W + "t"))


def split_inline_bullets(doc) -> int:
    """pdf2docx sometimes packs a whole bulleted list into one paragraph
    ("item one\u2022item two<br/>\u2022item three"): give every item its own paragraph."""
    made = 0
    for p in list(doc.element.body.iter(W + "p")):
        children = [c for c in p if c.tag != W + "pPr"]
        cuts, seen = [], False
        for idx, ch in enumerate(children):
            txt = _run_text(ch) if ch.tag in (W + "r", W + "hyperlink") else ""
            head = txt.strip()[:1]
            if seen and head and head in _INLINE_BULLETS:
                cuts.append(idx)
            if txt.strip():
                seen = True
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
        for ch in children:
            p.remove(ch)
        for ch in trim(segments[0]):
            p.append(ch)
        anchor = p
        for seg in segments[1:]:
            np_ = OxmlElement("w:p")
            if ppr is not None:
                ppr2 = copy.deepcopy(ppr)
                sp = ppr2.find(W + "spacing")
                if sp is not None and sp.get(W + "before"):
                    sp.set(W + "before", "20")   # the gap belonged before the first item only
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
                ind = ppr.find(W + "ind")
                if ind is None:
                    ind = OxmlElement("w:ind")
                    ppr.append(ind)
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
        mar = tblpr.find(W + "tblCellMar")
        if mar is None:
            mar = OxmlElement("w:tblCellMar")
            tblpr.append(mar)
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
        if len(runs) < 8 or len(singles) / len(runs) < 0.7:
            continue
        for r in runs:
            rpr = r.find(W + "rPr")
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            sz = rpr.find(W + "sz")
            size_pt = int(sz.get(W + "val")) / 2 if sz is not None and (sz.get(W + "val") or "").isdigit() else 10
            sp = rpr.find(W + "spacing")
            if sp is None:
                sp = OxmlElement("w:spacing")
                rpr.append(sp)
            sp.set(W + "val", str(round(size_pt * 0.3 * 20)))   # twentieths of a point
        fixed += 1
    return fixed


# ------------------------------------------------------------------ header / footer
def _set_border(paragraph, side: str, rule: layout.Item, gap_pt: float) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    bdr = ppr.find(W + "pBdr")
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        ppr.append(bdr)
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


def apply_body_rules(doc, rules: list) -> int:
    """Rules drawn under a paragraph in the PDF become that paragraph's bottom border."""
    norm = lambda t: re.sub(r"\s+", " ", t).strip()
    paras = [p for p in doc.paragraphs if p.text.strip()]
    applied, start = 0, 0
    for rule in rules:
        target = norm(rule["text_above"])
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


# ------------------------------------------------------------------ facade
def postprocess(path, bands: Optional[layout.Bands] = None, source_pdf=None, body_rules=None) -> dict:
    doc = Document(str(path))
    stats = {"fonts_mapped": normalize_fonts(doc)}
    sanitize(doc)
    stats.update(relax_heights(doc))
    stats["table_grids_fixed"] = fix_table_grids(doc)
    boundaries = merge_sections(doc)
    stats["sections_merged"] = len(boundaries)
    stats["tables_joined"] = join_split_tables(doc, boundaries)
    stats["cell_indents_fixed"] = fix_cell_indents(doc)
    zero_cell_vertical_margins(doc)
    stats["letter_spaced"] = restore_letter_spacing(doc)
    stats["bullet_items_split"] = split_inline_bullets(doc)
    stats["bullets_fixed"] = fix_bullets(doc)
    if body_rules:
        stats["rules_as_borders"] = apply_body_rules(doc, body_rules)
    if bands is not None and source_pdf is not None:
        stats.update({f"{k}_rebuilt": v for k, v in add_header_footer(doc, bands, source_pdf).items()})
    doc.save(str(path))
    return stats
