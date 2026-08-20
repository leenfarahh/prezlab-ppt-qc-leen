"""Report rendering: PDF and CSV projections of an audit manifest.

Both functions consume the manifest dict produced by
AuditResult.to_manifest(). The PDF is pure-python via reportlab (no slide
rendering); the CSV is the full record set in Appendix A.2 column order.

Arabic note: we register Segoe UI so Arabic codepoints render as glyphs
instead of tofu, but reportlab does no bidi/shaping, so connected Arabic
script in PDF cells is best-effort at pilot stage. The CSV is lossless.
"""

import csv
import html
import io
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

# Appendix A.2 order; also the CSV header. Do not reorder.
CSV_FIELDS = (
    "record_id", "job_id", "slide_index", "shape_id", "shape_path", "module",
    "issue_type", "property", "old_value", "new_value", "severity", "action",
    "confidence", "arabic_flag", "profile_rule_id", "message", "locator",
    "created_at",
)

_SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}
_SEVERITY_COLOR = {"error": "#C0392B", "warning": "#B9770E", "info": "#5D6D7E"}

_SEGOE_PATH = Path(r"C:\Windows\Fonts\segoeui.ttf")
_SEGOE_BOLD_PATH = Path(r"C:\Windows\Fonts\segoeuib.ttf")

_fonts = None


def _register_fonts() -> dict:
    """Return {'base': name, 'bold': name}, preferring Segoe UI so Arabic
    codepoints in messages do not tofu. Silent Helvetica fallback."""
    global _fonts
    if _fonts is not None:
        return _fonts
    base, bold = "Helvetica", "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("SegoeUI", str(_SEGOE_PATH)))
        base = "SegoeUI"
        bold = "SegoeUI"
        try:
            pdfmetrics.registerFont(TTFont("SegoeUI-Bold", str(_SEGOE_BOLD_PATH)))
            bold = "SegoeUI-Bold"
        except Exception:
            pass
    except Exception:
        pass
    _fonts = {"base": base, "bold": bold}
    return _fonts


def _sorted_records(records: list[dict]) -> list[dict]:
    return sorted(records, key=lambda r: (r.get("slide_index") or 0,
                                          _SEVERITY_ORDER.get(r.get("severity"), 3)))


def _generated_at(records: list[dict]) -> str | None:
    stamps = [r["created_at"] for r in records if r.get("created_at")]
    return max(stamps) if stamps else None


def _footer(canvas, doc):
    fonts = _register_fonts()
    canvas.saveState()
    canvas.setFont(fonts["base"], 8)
    canvas.setFillColor(colors.HexColor("#5D6D7E"))
    canvas.drawCentredString(A4[0] / 2, 10 * mm, f"Page {doc.page}")
    canvas.restoreState()


def render_pdf(manifest: dict) -> bytes:
    fonts = _register_fonts()
    records = _sorted_records(manifest.get("records") or [])
    summary = manifest.get("summary") or {}
    by_sev = summary.get("by_severity") or {}

    body = ParagraphStyle("body", fontName=fonts["base"], fontSize=8, leading=10)
    head = ParagraphStyle("head", parent=body, fontName=fonts["bold"],
                          textColor=colors.white)
    title = ParagraphStyle("title", parent=body, fontName=fonts["bold"],
                           fontSize=14, leading=18)
    meta = ParagraphStyle("meta", parent=body, fontSize=9, leading=12)

    def esc(value) -> str:
        return html.escape("" if value is None else str(value))

    story = [Paragraph(f"QC Audit Report: {esc(manifest.get('deck'))}", title),
             Spacer(1, 2 * mm)]
    meta_lines = [
        f"Profile: {esc(manifest.get('profile_id'))} "
        f"(v{esc(manifest.get('profile_version'))})",
        f"Slides: {esc(manifest.get('slides'))}",
    ]
    generated = _generated_at(records)
    if generated:
        meta_lines.append(f"Generated: {esc(generated)}")
    meta_lines.append(
        f"Findings: {by_sev.get('error', 0)} errors / "
        f"{by_sev.get('warning', 0)} warnings / {by_sev.get('info', 0)} info / "
        f"{summary.get('arabic_flagged', 0)} Arabic-flagged")
    for line in meta_lines:
        story.append(Paragraph(line, meta))
    story.append(Spacer(1, 6 * mm))

    if not records:
        story.append(Paragraph("No findings.", meta))
    else:
        rows = [[Paragraph(h, head)
                 for h in ("Slide", "Severity", "Issue", "Detail")]]
        style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDC3C7")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F4F6F7")]),
        ])
        for rec in records:
            sev = rec.get("severity") or ""
            sev_text = sev + (" [AR]" if rec.get("arabic_flag") else "")
            sev_style = ParagraphStyle(
                f"sev_{sev}", parent=body,
                textColor=colors.HexColor(_SEVERITY_COLOR.get(sev, "#000000")))
            rows.append([
                Paragraph(str((rec.get("slide_index") or 0) + 1), body),
                Paragraph(esc(sev_text), sev_style),
                Paragraph(esc(rec.get("issue_type")), body),
                Paragraph(esc(rec.get("message")), body),
            ])
        table = Table(rows, colWidths=[14 * mm, 24 * mm, 42 * mm, 90 * mm],
                      repeatRows=1)
        table.setStyle(style)
        story.append(table)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=18 * mm, bottomMargin=20 * mm,
                            title="QC Audit Report")
    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


def render_csv(manifest: dict) -> str:
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(CSV_FIELDS)
    for rec in manifest.get("records") or []:
        writer.writerow(["" if rec.get(f) is None else rec.get(f)
                         for f in CSV_FIELDS])
    return buf.getvalue()


def render_diff_pdf(deck_name: str, diff: dict) -> bytes:
    """Landscape PDF of the before/after review: one changed slide per page,
    PowerPoint-rendered images side by side, with the same highlight
    conventions as the web view (dashed orange = element before the fix,
    solid teal = the same element after). Consumes the cached diff dict
    from qc.render.build_diff, so no slide rendering happens here."""
    from reportlab.lib.colors import Color, HexColor
    from reportlab.lib.pagesizes import landscape
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as pdfcanvas

    fonts = _register_fonts()
    page_w, page_h = landscape(A4)
    margin = 36.0
    gap = 24.0
    img_w = (page_w - 2 * margin - gap) / 2

    teal = HexColor("#002528")
    slate = HexColor("#4a666e")
    orange = HexColor("#ff7c4a")
    green = HexColor("#0e7c66")

    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=(page_w, page_h))
    slides = diff.get("slides", [])
    images = diff.get("images", {})

    def _overlays(rects, x0, y0, w, h, stroke, fill, dashed):
        c.saveState()
        c.setLineWidth(1.4)
        c.setStrokeColor(stroke)
        c.setFillColor(fill)
        if dashed:
            c.setDash(4, 3)
        for r in rects:
            rx = x0 + r["x"] * w
            rh = r["h"] * h
            ry = y0 + (1 - r["y"] - r["h"]) * h  # PDF origin is bottom-left
            c.rect(rx, ry, r["w"] * w, rh, stroke=1, fill=1)
        c.restoreState()

    for page_idx, sl in enumerate(slides):
        idx = sl["index"]
        y = page_h - margin

        if page_idx == 0:
            c.setFont(fonts["bold"], 16)
            c.setFillColor(teal)
            c.drawString(margin, y - 14, "Before / after review")
            c.setFont(fonts["base"], 10)
            c.setFillColor(slate)
            c.drawString(margin, y - 30,
                         f"{deck_name}  |  {len(slides)} slide"
                         f"{'s' if len(slides) != 1 else ''} changed  |  "
                         "rendered by PowerPoint, original file untouched")
            y -= 52

        n = sl["changes"]
        c.setFont(fonts["bold"], 13)
        c.setFillColor(teal)
        c.drawString(margin, y - 12, f"Slide {idx + 1}")
        c.setFont(fonts["base"], 9)
        c.setFillColor(slate)
        c.drawString(margin + 60, y - 12,
                     f"{n} change{'s' if n != 1 else ''}  |  "
                     + "  ".join(sl.get("labels", [])))
        y -= 30

        before_png = images.get(f"before:{idx}")
        after_png = images.get(f"after:{idx}")
        if not before_png or not after_png:
            continue
        reader_b = ImageReader(io.BytesIO(before_png))
        iw, ih = reader_b.getSize()
        img_h = img_w * ih / iw

        for tag, png, rects, x0 in (
            ("BEFORE", before_png, sl["before_rects"], margin),
            ("AFTER", after_png, sl["after_rects"], margin + img_w + gap),
        ):
            c.setFont(fonts["bold"], 8)
            c.setFillColor(slate)
            c.drawString(x0, y - 8, tag)
            iy = y - 14 - img_h
            c.drawImage(ImageReader(io.BytesIO(png)), x0, iy,
                        width=img_w, height=img_h)
            c.setStrokeColor(HexColor("#c9d2d4"))
            c.setLineWidth(0.5)
            c.rect(x0, iy, img_w, img_h, stroke=1, fill=0)
            if tag == "BEFORE":
                _overlays(rects, x0, iy, img_w, img_h,
                          orange, Color(1, 0.49, 0.29, alpha=0.12), dashed=True)
            else:
                _overlays(rects, x0, iy, img_w, img_h,
                          green, Color(0.05, 0.49, 0.4, alpha=0.10), dashed=False)

        # legend + page footer
        c.setFont(fonts["base"], 8)
        c.setFillColor(slate)
        c.drawString(margin, margin - 10,
                     "Dashed orange: changed element before the fix.  "
                     "Solid teal: the same element after.")
        c.drawRightString(page_w - margin, margin - 10,
                          f"Page {page_idx + 1} of {len(slides)}")
        c.showPage()

    if not slides:
        c.setFont(fonts["bold"], 14)
        c.setFillColor(teal)
        c.drawString(margin, page_h / 2, "No applied fixes to compare.")
        c.showPage()
    c.save()
    return buf.getvalue()


def render_visual_audit_pdf(manifest: dict, images: dict[int, bytes],
                            rects: dict[int, list[dict]]) -> bytes:
    """Landscape visual audit report: page 1 is an executive summary, then
    one page per flagged slide (any slide with a non-preflight record) with
    the PowerPoint-rendered image and finding highlights overlaid. Consumes
    pre-rendered PNGs and qc.render.audit_rects output, so pure reportlab."""
    from reportlab.lib.colors import Color, HexColor
    from reportlab.lib.pagesizes import landscape
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as pdfcanvas

    fonts = _register_fonts()
    page_w, page_h = landscape(A4)
    margin = 36.0

    teal = HexColor("#002528")
    slate = HexColor("#4a666e")
    severity_hex = {"error": "#40182d", "warning": "#ff7c4a", "info": "#62848c"}
    severity_name = {0: "error", 1: "warning", 2: "info"}

    records = manifest.get("records") or []
    summary = manifest.get("summary") or {}
    by_sev = summary.get("by_severity") or {}

    flagged: dict[int, list[dict]] = {}
    for rec in records:
        if rec.get("module") == "preflight":
            continue
        flagged.setdefault(int(rec.get("slide_index") or 0), []).append(rec)
    slide_order = sorted(flagged)
    total_pages = 1 + len(slide_order)

    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=(page_w, page_h))

    def _trim(value, limit: int = 110) -> str:
        text = "" if value is None else str(value)
        return text if len(text) <= limit else text[:limit - 3] + "..."

    def _page_footer(page_num: int):
        c.setFont(fonts["base"], 8)
        c.setFillColor(slate)
        c.drawString(margin, margin - 10,
                     "Numbered boxes match #N in the findings below; "
                     "dashed = Arabic content (manual review)")
        c.drawRightString(page_w - margin, margin - 10,
                          f"Page {page_num} of {total_pages}")

    # -- page 1: executive summary
    y = page_h - margin
    c.setFont(fonts["bold"], 16)
    c.setFillColor(teal)
    c.drawString(margin, y - 14, "Audit report")
    deck_name = Path(str(manifest.get("deck") or "")).name or "(unnamed deck)"
    c.setFont(fonts["base"], 10)
    c.setFillColor(slate)
    y -= 34
    for line in (
        f"Deck: {deck_name}",
        f"Profile: {manifest.get('profile_id')} "
        f"(v{manifest.get('profile_version')})",
        f"Slides: {manifest.get('slides')}",
        f"Findings: {by_sev.get('error', 0)} errors / "
        f"{by_sev.get('warning', 0)} warnings / {by_sev.get('info', 0)} info / "
        f"{summary.get('arabic_flagged', 0)} Arabic-flagged",
    ):
        c.drawString(margin, y, line)
        y -= 14

    counts: dict[str, dict] = {}
    for rec in records:
        slot = counts.setdefault(str(rec.get("issue_type") or "(unknown)"),
                                 {"count": 0, "worst": 3})
        slot["count"] += 1
        slot["worst"] = min(slot["worst"],
                            _SEVERITY_ORDER.get(rec.get("severity"), 3))
    issue_rows = sorted(counts.items(), key=lambda kv: (-kv[1]["count"], kv[0]))

    y -= 8
    if issue_rows:
        c.setFont(fonts["bold"], 9)
        c.setFillColor(teal)
        c.drawString(margin, y, "Issue type")
        c.drawRightString(margin + 350, y, "Count")
        c.drawString(margin + 380, y, "Worst severity")
        y -= 12
        c.setFont(fonts["base"], 9)
        c.setFillColor(slate)
        for issue_type, slot in issue_rows[:18]:
            c.drawString(margin, y, _trim(issue_type, 70))
            c.drawRightString(margin + 350, y, str(slot["count"]))
            c.drawString(margin + 380, y,
                         severity_name.get(slot["worst"], "info"))
            y -= 11
        if len(issue_rows) > 18:
            c.drawString(margin, y, f"and {len(issue_rows) - 18} more")
            y -= 11
    else:
        c.setFont(fonts["base"], 9)
        c.setFillColor(slate)
        c.drawString(margin, y, "No findings.")
        y -= 11

    top = sorted(flagged.items(), key=lambda kv: (-len(kv[1]), kv[0]))[:5]
    if top:
        y -= 8
        c.setFont(fonts["base"], 9)
        c.setFillColor(slate)
        c.drawString(margin, y, "Top slides: " + ", ".join(
            f"slide {idx + 1} ({len(recs)})" for idx, recs in top))
    _page_footer(1)
    c.showPage()

    # -- one page per flagged slide
    from .render import pin_numbers

    _pins = pin_numbers(records)
    for page_i, idx in enumerate(slide_order):
        recs = sorted(flagged[idx],
                      key=lambda r: _SEVERITY_ORDER.get(r.get("severity"), 3))
        n = len(recs)
        y_top = page_h - margin
        c.setFont(fonts["bold"], 13)
        c.setFillColor(teal)
        c.drawString(margin, y_top - 12, f"Slide {idx + 1}")
        c.setFont(fonts["base"], 9)
        c.setFillColor(slate)
        c.drawString(margin + 60, y_top - 12,
                     f"{n} finding{'s' if n != 1 else ''}")

        def _line(r):
            p = _pins.get(r.get("record_id"))
            tag = (f"#{p} " if p is not None
                   else "(whole slide) " if str(r.get("shape_id") or "-") == "-"
                   else "")
            return (f"{tag}[{r.get('severity')}] {r.get('issue_type')}: "
                    f"{_trim(r.get('message'))}")

        lines = [_line(r) for r in recs[:8]]
        if n > 8:
            lines.append(f"and {n - 8} more findings")

        text_h = len(lines) * 11 + 10
        box_top = y_top - 28
        box_bottom = margin + text_h
        box_w = page_w - 2 * margin
        box_h = box_top - box_bottom

        png = images.get(idx)
        if png:
            reader = ImageReader(io.BytesIO(png))
            iw, ih = reader.getSize()
            scale = min(box_w / iw, box_h / ih)
            img_w, img_h = iw * scale, ih * scale
            img_x = margin + (box_w - img_w) / 2
            img_y = box_bottom + (box_h - img_h) / 2
            c.drawImage(reader, img_x, img_y, width=img_w, height=img_h)
            c.setStrokeColor(HexColor("#c9d2d4"))
            c.setLineWidth(0.5)
            c.rect(img_x, img_y, img_w, img_h, stroke=1, fill=0)
            for r in rects.get(idx) or []:
                stroke = HexColor(severity_hex.get(r.get("severity"),
                                                   "#62848c"))
                c.saveState()
                c.setLineWidth(1.4)
                c.setStrokeColor(stroke)
                c.setFillColor(Color(stroke.red, stroke.green, stroke.blue,
                                     alpha=0.10))
                if r.get("arabic"):
                    c.setDash(4, 3)
                ry = img_y + (1 - r["y"] - r["h"]) * img_h  # PDF origin is bottom-left
                c.rect(img_x + r["x"] * img_w, ry,
                       r["w"] * img_w, r["h"] * img_h, stroke=1, fill=1)
                c.restoreState()
                if r.get("pin") is not None:
                    # numbered badge at the box's top-left, matching the list
                    cx = img_x + r["x"] * img_w
                    cy = ry + r["h"] * img_h
                    c.saveState()
                    c.setFillColor(stroke)
                    c.setStrokeColor(HexColor("#ffffff"))
                    c.setLineWidth(1)
                    c.circle(cx, cy, 6.5, stroke=1, fill=1)
                    c.setFillColor(HexColor("#ffffff"))
                    c.setFont(fonts["bold"], 7)
                    c.drawCentredString(cx, cy - 2.4, str(r["pin"]))
                    c.restoreState()
        else:
            c.setFont(fonts["base"], 10)
            c.setFillColor(slate)
            c.drawCentredString(page_w / 2, box_bottom + box_h / 2,
                                "(slide image unavailable)")

        ty = box_bottom - 12
        c.setFont(fonts["base"], 8)
        c.setFillColor(slate)
        for line in lines:
            c.drawString(margin, ty, line)
            ty -= 11
        _page_footer(page_i + 2)
        c.showPage()

    c.save()
    return buf.getvalue()
