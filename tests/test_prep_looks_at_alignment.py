"""The prepare run asks the visual model about alignment, not just about layouts.

The alignment judgments used to live behind a button on the audit report, which
is two clicks down from the prepared deck and announces itself nowhere. A
feature a designer has to know to press is a feature that does not run, and the
complaint that started this was "it did not detect any of it" (design lead,
31/08/2026).

So the same "Look at the slides" tick that decides whether the unplaced slides
get looked at now also decides whether the REBUILT deck gets an alignment pass.
"""

import pytest

from qc import web


@pytest.fixture()
def job():
    """A registered prepare job, as _register_prep leaves one."""
    j = {
        "deck": b"PK-deck", "filename": "d.pptx", "profile": "prezlab_en",
        "manifest": {"slides": 2, "summary": {"total": 0}, "records": []},
        "prep": object(),
    }
    return j


def _stub_pass(monkeypatch, records, reviewed=2):
    seen = {}

    def _thumbs(job_id, j):
        j["thumbs"] = {0: b"png", 1: b"png"}

    def _run(deck, thumbs, manifest):
        seen["deck"] = deck
        seen["thumbs"] = thumbs
        return list(records), reviewed

    monkeypatch.setattr(web, "_ensure_thumbs", _thumbs)
    monkeypatch.setattr("qc.copilot.run_copilot", _run)
    monkeypatch.setattr(web, "_can_look", lambda: (True, ""))
    return seen


def _record(shape_id="7"):
    return {"slide_index": 0, "shape_id": shape_id, "severity": "warning",
            "issue_type": "margin_alignment.edge_misaligned",
            "module": "margin_alignment", "arabic_flag": False,
            "record_id": f"r{shape_id}", "confidence": "medium",
            "action": "flagged", "message": "Design copilot: snap it."}


def test_the_findings_land_in_the_manifest(job, monkeypatch):
    _stub_pass(monkeypatch, [_record()])

    web._prep_layout_review("j1", job)

    assert len(job["manifest"]["records"]) == 1
    assert job["manifest"]["summary"]["total"] == 1
    assert job["layout_ok"] is True
    assert "1 alignment thing" in job["layout_note"]


def test_it_reads_the_rebuilt_deck_not_the_upload(job, monkeypatch):
    """Judging the raw file means judging one nobody will send: the master is
    about to move half of it."""
    seen = _stub_pass(monkeypatch, [])
    job["source"] = b"PK-the-messy-upload"

    web._prep_layout_review("j1", job)

    assert seen["deck"] == b"PK-deck"


def test_finding_nothing_is_said_out_loud(job, monkeypatch):
    """Silence and "nothing is wrong" must not look the same on the page."""
    _stub_pass(monkeypatch, [])

    web._prep_layout_review("j1", job)

    assert job["layout_ok"] is True
    assert "0 alignment things" in job["layout_note"]
    assert "reads as out of line" in job["layout_note"]


def test_a_host_that_cannot_look_says_which_weaker_answer_this_is(job,
                                                                 monkeypatch):
    monkeypatch.setattr(web, "_can_look",
                        lambda: (False, "No model key is configured."))

    web._prep_layout_review("j1", job)

    assert job["layout_ok"] is False
    assert "No model key" in job["layout_note"]
    assert job["manifest"]["records"] == []


def test_a_pass_that_blows_up_costs_the_judgments_and_nothing_else(job,
                                                                   monkeypatch):
    """The deliverable is the rebuilt deck. Losing the alignment judgments must
    not lose it, or the audit that came with it."""
    _stub_pass(monkeypatch, [])

    def _boom(deck, thumbs, manifest):
        raise RuntimeError("the renderer fell over")

    monkeypatch.setattr("qc.copilot.run_copilot", _boom)

    web._prep_layout_review("j1", job)

    assert job["layout_ok"] is False
    assert "the renderer fell over" in job["layout_note"]
    assert "still stands" in job["layout_note"]
    assert job["deck"] == b"PK-deck"


def test_a_run_whose_audit_failed_is_left_alone(job, monkeypatch):
    """There is no manifest to merge into, and the page already says the audit
    did not run. Saying it twice is not more honest."""
    _stub_pass(monkeypatch, [_record()])
    job["manifest"] = None

    web._prep_layout_review("j1", job)

    assert job["layout_note"] == ""
    assert job["layout_ok"] is True


def test_the_thumbnails_are_kept_for_the_design_page(job, monkeypatch):
    """Design QC renders every slide too. Rendering them twice through
    PowerPoint is a minute a designer waits for nothing."""
    _stub_pass(monkeypatch, [])

    web._prep_layout_review("j1", job)

    assert job["thumbs"], "the renders were thrown away"


def test_the_tick_is_what_decides(monkeypatch):
    """Not asked for means not sent. Slide images leave the machine on this
    pass, so it must never happen without the box ticked.

    The tick is now collected one press earlier than it is acted on: applying a
    master stops for the layout decisions, and the alignment pass still has to
    run on the REBUILT slides, so the box on step 1 is remembered on the plan
    and read on step 2. Both halves are pinned, because a tick that is read
    without being stored, or stored without being read, sends the images either
    always or never."""
    import inspect

    upload = inspect.getsource(web.prep_deck)
    assert '_plans[plan_id]["look"] = bool(look)' in upload, (
        "step 1 must remember the tick")
    assert "_prep_layout_review" not in upload, (
        "nothing is looked at before the rebuild; that would judge the upload")

    apply = inspect.getsource(web.prep_apply_layouts)
    assert 'if held.get("look"):' in apply
    assert "_prep_layout_review" in apply
