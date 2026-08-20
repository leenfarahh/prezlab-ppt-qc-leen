"""Engine integration: full pipeline over the fixture corpus."""

import json

import pytest

from qc.engine import run_audit
from qc.records import MODULES


def test_full_audit_over_bilingual_fixture(fixtures_dir):
    result = run_audit(fixtures_dir / "bilingual_ar.pptx", "prezlab_bilingual")
    assert result.slides == 1
    assert result.summary["total"] == len(result.records)
    # The Arabic guard must surface: at least one record carries arabic_flag
    # (the AR run has no cs typeface in one case and Inter latin in another).
    assert result.summary["arabic_flagged"] >= 1
    issue_types = {r.issue_type for r in result.records}
    assert "font.cs_typeface_missing" in issue_types


def test_preflight_flags_chart_in_heavy_fixture(fixtures_dir):
    result = run_audit(fixtures_dir / "heavy.pptx", "prezlab_en", modules=["header_footer"])
    pre = [r for r in result.records if r.issue_type == "preflight.unmodifiable_content"]
    assert any("chart" in r.message for r in pre)
    assert all(r.severity == "info" and r.action == "flagged" for r in pre)


def test_module_selection_runs_subset(fixtures_dir):
    result = run_audit(fixtures_dir / "clean.pptx", "prezlab_en", modules=["font"])
    assert set(result.summary["by_module"]) <= {"font", "preflight"}


def test_unknown_module_rejected(fixtures_dir):
    with pytest.raises(ValueError):
        run_audit(fixtures_dir / "clean.pptx", "prezlab_en", modules=["not_a_module"])


def test_manifest_round_trips(fixtures_dir, tmp_path):
    result = run_audit(fixtures_dir / "theme_colors.pptx", "prezlab_en")
    out = tmp_path / "manifest.json"
    result.save_manifest(out)
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["profile_id"] == "prezlab_en"
    assert len(data["records"]) == data["summary"]["total"]
    required = {"record_id", "slide_index", "shape_id", "module", "issue_type",
                "severity", "action", "confidence", "arabic_flag", "message"}
    for rec in data["records"]:
        assert required <= set(rec)
        assert rec["module"] in MODULES + ("preflight",)


def test_all_modules_run_on_all_fixtures_without_crashing(fixtures_dir):
    for deck in ("clean.pptx", "bilingual_ar.pptx", "theme_colors.pptx",
                 "mixed_layouts.pptx", "heavy.pptx"):
        result = run_audit(fixtures_dir / deck, "prezlab_en")
        assert result.summary["total"] >= 0
