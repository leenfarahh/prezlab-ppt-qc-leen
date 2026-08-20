"""Tests for render_visual_audit_pdf: pure reportlab, no COM.

Slide images are synthetic 1x1 PNGs; the realistic case runs the real audit
on the fixture corpus and builds rects through qc.render.audit_rects.
"""

import base64

from qc.report import render_visual_audit_pdf

# Smallest valid PNG (1x1 opaque pixel), enough for ImageReader.getSize().
_PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8"
    "z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")


def _record(slide_index=0, issue_type="text.font_family", severity="warning",
            message="Body run uses Arial; profile expects Poppins.",
            module="text", arabic_flag=False, action="flagged", rid=1):
    return {
        "record_id": f"r{rid:04d}", "slide_index": slide_index,
        "shape_id": "3", "shape_path": "3", "module": module,
        "issue_type": issue_type, "severity": severity, "action": action,
        "arabic_flag": arabic_flag, "message": message,
        "created_at": "2026-07-12T10:00:00",
    }


def _manifest(records, slides=3):
    sev = {}
    for rec in records:
        sev[rec["severity"]] = sev.get(rec["severity"], 0) + 1
    return {
        "deck": r"C:\decks\sample.pptx",
        "profile_id": "prezlab_en", "profile_version": 1, "slides": slides,
        "summary": {
            "by_severity": sev,
            "arabic_flagged": sum(1 for r in records if r["arabic_flag"]),
            "total": len(records),
        },
        "records": records,
    }


def _page_count(pdf: bytes) -> int:
    return pdf.count(b"/Type /Page") - pdf.count(b"/Type /Pages")


def _rect(severity="error", arabic=False):
    return {"x": 0.1, "y": 0.2, "w": 0.5, "h": 0.3,
            "label": "text.font_family", "severity": severity,
            "arabic": arabic, "record_ids": ["r0001"]}


def test_full_manifest_with_images_and_rects():
    records = [
        _record(slide_index=0, severity="error", rid=1),
        _record(slide_index=0, severity="info", arabic_flag=True, rid=2),
        _record(slide_index=2, severity="warning", rid=3),
        _record(slide_index=0, module="preflight",
                issue_type="preflight.unmodifiable_content",
                severity="info", rid=4),
    ]
    images = {i: _PNG_1X1 for i in range(3)}
    rects = {0: [_rect("error"), _rect("info", arabic=True)],
             2: [_rect("warning")]}
    pdf = render_visual_audit_pdf(_manifest(records), images, rects)
    assert pdf.startswith(b"%PDF-")
    # summary page + slides 1 and 3; the preflight-only record adds no page
    assert _page_count(pdf) == 3


def test_missing_image_for_flagged_slide():
    records = [_record(slide_index=1, rid=1)]
    pdf = render_visual_audit_pdf(_manifest(records), {}, {})
    assert pdf.startswith(b"%PDF-")
    assert _page_count(pdf) == 2


def test_empty_rects_and_over_eight_findings():
    records = [_record(slide_index=0, severity="warning", rid=i,
                       message="m" * 200) for i in range(12)]
    images = {0: _PNG_1X1}
    pdf = render_visual_audit_pdf(_manifest(records), images, {0: []})
    assert pdf.startswith(b"%PDF-")
    assert _page_count(pdf) == 2


def test_summary_table_with_many_issue_types():
    records = [_record(slide_index=i % 4, issue_type=f"synthetic.type_{i:02d}",
                       severity=("error", "warning", "info")[i % 3], rid=i)
               for i in range(25)]
    images = {i: _PNG_1X1 for i in range(4)}
    pdf = render_visual_audit_pdf(_manifest(records, slides=4), images, {})
    assert pdf.startswith(b"%PDF-")
    assert _page_count(pdf) == 5


def test_empty_manifest_summary_page_only():
    pdf = render_visual_audit_pdf(_manifest([], slides=5), {}, {})
    assert pdf.startswith(b"%PDF-")
    assert _page_count(pdf) == 1


def test_realistic_audit_manifest(fixtures_dir):
    from qc.engine import run_audit
    from qc.render import audit_rects

    deck = fixtures_dir / "mixed_layouts.pptx"
    manifest = run_audit(deck, "prezlab_en").to_manifest()
    images = {i: _PNG_1X1 for i in range(manifest["slides"])}
    rects = audit_rects(deck.read_bytes(), manifest["records"])
    pdf = render_visual_audit_pdf(manifest, images, rects)
    assert pdf.startswith(b"%PDF-")
    flagged = {r["slide_index"] for r in manifest["records"]
               if r["module"] != "preflight"}
    assert _page_count(pdf) == 1 + len(flagged)
