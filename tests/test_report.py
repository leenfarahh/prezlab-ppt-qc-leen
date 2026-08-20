"""Tests for qc.report: PDF and CSV projections of an audit manifest."""

import csv
import io

from qc.engine import run_audit
from qc.report import CSV_FIELDS, render_csv, render_pdf


def _base_record(**overrides) -> dict:
    rec = {
        "record_id": "abc123",
        "job_id": None,
        "slide_index": 0,
        "shape_id": "2",
        "shape_path": None,
        "module": "font",
        "issue_type": "font.family_out_of_set",
        "property": "rPr.latin.typeface",
        "old_value": "Calibri",
        "new_value": "Georgia",
        "severity": "error",
        "action": "flagged",
        "confidence": "high",
        "arabic_flag": False,
        "profile_rule_id": "font.roles.title.latin",
        "message": "Latin family 'Calibri' not allowed.",
        "locator": "p0/r0",
        "created_at": "2026-07-05T12:00:00+00:00",
    }
    rec.update(overrides)
    return rec


def _manifest(records: list[dict]) -> dict:
    return {
        "deck": "hand_built.pptx",
        "profile_id": "prezlab_bilingual",
        "profile_version": 1,
        "slides": 3,
        "summary": {"by_severity": {}, "by_issue_type": {}, "by_module": {},
                    "arabic_flagged": 0, "total": len(records)},
        "records": records,
    }


def test_pdf_real_manifest(fixtures_dir):
    manifest = run_audit(fixtures_dir / "bilingual_ar.pptx",
                         "prezlab_bilingual").to_manifest()
    assert manifest["records"], "fixture should produce findings"
    pdf = render_pdf(manifest)
    assert isinstance(pdf, bytes)
    assert pdf.startswith(b"%PDF-")
    assert len(pdf) > 1500


def test_pdf_empty_manifest():
    pdf = render_pdf(_manifest([]))
    assert pdf.startswith(b"%PDF-")


def test_csv_round_trip(fixtures_dir):
    manifest = run_audit(fixtures_dir / "bilingual_ar.pptx",
                         "prezlab_bilingual").to_manifest()
    rows = list(csv.reader(io.StringIO(render_csv(manifest))))
    assert len(rows) == len(manifest["records"]) + 1
    assert rows[0] == list(CSV_FIELDS)
    first = manifest["records"][0]
    issue_col = CSV_FIELDS.index("issue_type")
    message_col = CSV_FIELDS.index("message")
    assert rows[1][issue_col] == first["issue_type"]
    assert rows[1][message_col] == first["message"]


def test_csv_quoting_survives_commas_quotes_arabic():
    message = 'Font "Dubai", size 12, نص عربي\nsecond line'
    manifest = _manifest([_base_record(message=message, arabic_flag=True)])
    rows = list(csv.reader(io.StringIO(render_csv(manifest))))
    assert len(rows) == 2
    assert rows[1][CSV_FIELDS.index("message")] == message


def test_pdf_paginates_many_records():
    records = [_base_record(record_id=f"rec{i:04d}",
                            slide_index=i % 7,
                            severity=("error", "warning", "info")[i % 3],
                            arabic_flag=(i % 5 == 0),
                            message=f"Finding {i}: " + "long detail text " * 8)
               for i in range(120)]
    pdf = render_pdf(_manifest(records))
    assert pdf.startswith(b"%PDF-")
    assert len(pdf) > 5000
