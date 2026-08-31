"""One destination with tabs, and a report that leads with the interesting half.

Two complaints drove this, and they are the same complaint at two scales.

AT THE PAGE LEVEL: six views of one prepared deck, none of them linking to the
others. A designer on Design QC could reach the upload form and nothing else,
so the way between them was the back button and the way to the profile they
were auditing against was to remember the URL.

AT THE ROW LEVEL: the four findings worth reading sat wherever severity put
them, in a list of two thousand. Measured findings are facts and mostly boring;
what the visual model noticed is the answer to the question someone opened the
tool to ask, and it now sorts above them and is marked so it stays findable
after a filter.
"""

from qc.records import make_record
from qc.ui import JOB_TABS, job_tabs, render_report


# ------------------------------------------------------------------- the tabs


def test_the_strip_reaches_every_view():
    html = job_tabs("j1", "findings")
    for _key, label, href in JOB_TABS:
        assert label in html
        assert href.format(id="j1") in html


def test_the_current_tab_is_marked_for_a_screen_reader_too():
    html = job_tabs("j1", "design")
    assert '<a href="/design/j1" aria-current="page">Design QC</a>' in html
    assert 'href="/audit/j1" aria-current' not in html


def test_a_view_this_job_does_not_have_is_left_out():
    """A tab that 404s reads as the tool losing the page, which is worse than
    an absent one. A deck that was audited and never rebuilt has no
    before/after."""
    html = job_tabs("j1", "findings", available={"findings", "design"})
    assert "/format/j1/review" not in html
    assert "/diff/j1" not in html
    assert "/audit/j1" in html


def test_a_count_rides_along_when_there_is_one_worth_seeing():
    html = job_tabs("j1", "findings", counts={"design": 5})
    assert "Design QC <b>5</b>" in html
    # Zero is not a badge. "Findings 0" invites a click that shows nothing.
    assert job_tabs("j1", "findings", counts={"design": 0}).count("<b>") == 0


# ------------------------------------------------ which tabs a job really has


def _job(**over):
    job = {"prep": None, "plans": [], "deck": None, "manifest": None,
           "applied_records": None, "prev_deck": None}
    job.update(over)
    return job


def test_an_audit_only_job_offers_no_before_and_after():
    from qc.web import _tabs_for

    html = _tabs_for("j1", _job(deck=b"x", manifest={"records": []}), "findings")
    assert "/format/j1/review" not in html, "nothing was rebuilt to compare"
    assert "/audit/j1" in html
    assert "/checklist/j1" in html


def test_a_run_whose_audit_failed_offers_no_findings_or_design():
    from qc.web import _tabs_for

    html = _tabs_for("j1", _job(prep=object(), plans=[1], deck=b"x"),
                     "overview")
    assert "/audit/j1" not in html
    assert "/design/j1" not in html
    assert "/prep/j1" in html


def test_a_job_with_one_view_gets_no_strip_at_all():
    """A strip with one tab on it is a label, not navigation."""
    from qc.web import _tabs_for

    assert _tabs_for("j1", _job(deck=b"x"), "checklist") == ""
    assert _tabs_for("j1", None, "findings") == ""


def test_the_findings_tab_counts_real_findings_only():
    """Preflight rows are notes about the run, not defects. Counting them puts
    a number on the tab that does not match the list behind it."""
    from qc.web import _tabs_for

    manifest = {"records": [{"module": "font"}, {"module": "font"},
                            {"module": "preflight"}]}
    html = _tabs_for("j1", _job(deck=b"x", manifest=manifest), "findings")
    assert "Findings <b>2</b>" in html


# ------------------------------------------------------------- vision leads


def _rec(source, severity="warning", issue="font.family_out_of_set", slide=0):
    return make_record(slide_index=slide, shape_id="7", module="font",
                       issue_type=issue, message=f"a {source} finding",
                       severity=severity, source=source).to_dict()


def _manifest(records):
    return {"deck": "client.pptx", "slides": 1, "profile_id": "prezlab_en",
            "profile_version": 1, "records": records,
            "summary": {"total": len(records), "by_severity": {},
                        "arabic_flagged": 0}}


def test_a_vision_finding_sorts_above_a_worse_measured_one():
    """Severity orders WITHIN each half, not across them. An error the rules
    measured is still a fact about a font; a warning the model raised is what a
    designer opened the tool to see."""
    measured = _rec("measured", severity="error")
    seen = _rec("vision", severity="warning")
    html = render_report(_manifest([measured, seen]), "j1")

    assert html.index("a vision finding") < html.index("a measured finding")


def test_a_vision_finding_stays_marked_once_the_list_is_rearranged():
    """Sorted first is not enough: a filter or a search reorders the table, and
    the mark is what keeps it identifiable afterwards."""
    html = render_report(_manifest([_rec("vision")]), "j1")
    assert 'class="pill vision"' in html
    assert 'data-seen="1"' in html


def test_a_measured_finding_carries_no_mark():
    html = render_report(_manifest([_rec("measured")]), "j1")
    assert 'class="pill vision"' not in html
    assert 'data-seen="0"' in html


def test_the_filter_chip_counts_what_the_model_saw():
    html = render_report(
        _manifest([_rec("vision"), _rec("vision"), _rec("measured")]), "j1")
    assert "Seen by the model 2" in html
    assert 'data-f="seen"' in html


def test_a_record_defaults_to_measured():
    """Every existing emitter says nothing about provenance and means the rules
    measured it. A default of "vision" would relabel the whole audit."""
    assert _rec("measured")["source"] == "measured"
    plain = make_record(slide_index=0, shape_id="1", module="font",
                        issue_type="font.family_out_of_set", message="x")
    assert plain.source == "measured"


def test_the_report_survives_a_manifest_from_before_provenance_existed():
    """An archived manifest has no `source` key at all. It must read as
    measured rather than crashing the page it is on."""
    old = {"record_id": "r1", "slide_index": 0, "shape_id": "1",
           "module": "font", "issue_type": "font.family_out_of_set",
           "severity": "warning", "action": "flagged", "confidence": "high",
           "arabic_flag": False, "message": "an archived finding",
           "locator": None, "shape_path": None, "property": None,
           "old_value": None, "new_value": None, "profile_rule_id": None,
           "job_id": None, "created_at": ""}
    html = render_report(_manifest([old]), "j1")
    assert "an archived finding" in html
    assert 'data-seen="0"' in html
