"""The pause between an upload and a rewritten file.

Applying a master was one press that guessed at the slides it could not place,
and the only way to see what it had guessed was to open the result - by which
point the guess was in the file. The step in between is where a designer says
which layout each of those slides belongs on.

What the route has to get right:

  - the first press REBUILDS NOTHING, so a wrong deck or a wrong profile costs a
    moment rather than a rebuild;
  - the picks posted are the layouts applied, or the page is decoration;
  - a host with no renderer still gets a usable page, in words;
  - and an expired plan says so rather than 500ing, because the layout page is
    the one page in the tool a designer sits on for a while.

apply_master is stubbed throughout: everything asserted here has to hold on a
machine with no PowerPoint.
"""

import io

import pytest
from fastapi.testclient import TestClient
from pptx import Presentation
from pptx.util import Emu, Inches

from qc import web
from qc.applymaster import ApplyResult


# ------------------------------------------------------------------ fixtures


def _bytes(prs) -> bytes:
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _master() -> bytes:
    return _bytes(Presentation())


def _deck(slides=2) -> bytes:
    """A deck of blank slides carrying two columns of content. Blank matches the
    master's own 'Blank' layout by name, so nothing is uncertain until a test
    makes it so."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    for _ in range(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for left in (0.8, 7.0):
            box = slide.shapes.add_textbox(Inches(left), Inches(2.0),
                                           Inches(5.0), Inches(3.0))
            box.text_frame.text = "Point"
    return _bytes(prs)


@pytest.fixture()
def client(monkeypatch):
    monkeypatch.setattr(web, "AUTH_REQUIRED", False)
    monkeypatch.setattr(web.app.state, "auth_required", False)
    web._plans.clear()
    return TestClient(web.app)


@pytest.fixture()
def no_powerpoint(monkeypatch):
    seen = {}

    def _fake(deck_bytes, master_bytes, plans):
        seen["plans"] = plans
        out = Presentation(io.BytesIO(deck_bytes))
        buf = io.BytesIO()
        out.save(buf)
        return ApplyResult(deck=buf.getvalue(), plans=plans, masters=1)

    monkeypatch.setattr("qc.applymaster.apply_master", _fake)
    monkeypatch.setattr("qc.engine.run_audit", lambda *a, **k: type(
        "R", (), {"to_manifest": lambda self: {
            "slides": 2, "summary": {"total": 0, "by_severity": {},
                                     "arabic_flagged": 0},
            "records": [], "profile_id": "p", "profile_version": 1}})())
    return seen


@pytest.fixture()
def no_renderer(monkeypatch):
    def _boom(*a, **k):
        raise RuntimeError("PowerPoint is not on this machine")

    monkeypatch.setattr("qc.render.export_decks_png", _boom)


@pytest.fixture()
def profile_with_master(monkeypatch):
    """A profile carrying a real master, saved into the test's own store."""
    from qc.profile import PROFILES_DIR
    from qc.templates import save_master
    import json

    pid = "layout_step_client"
    (PROFILES_DIR / f"{pid}.json").write_text(json.dumps({
        "id": pid, "name": "Layout Step Client", "version": 1,
        "config": {}}), encoding="utf-8")
    save_master(pid, _master())
    return pid


def _uncertain(monkeypatch):
    """Force every slide to be a fallback, so the page has something to ask."""
    import qc.applymaster as AM

    real = AM.plan_assignments

    def _all_fallback(deck_prs, layouts):
        plans = real(deck_prs, layouts)
        for plan in plans:
            plan.match_rule = "fallback"
            plan.note = "no layout named 'Blank'"
        return plans

    monkeypatch.setattr(AM, "plan_assignments", _all_fallback)


def _start(client, pid, slides=2):
    return client.post("/prep", data={"profile": pid},
                       files={"deck": ("rough.pptx", _deck(slides), "app/x")})


# --------------------------------------------------------- the first press


def test_the_first_press_rebuilds_nothing(client, profile_with_master,
                                          no_renderer, monkeypatch):
    """The whole point of the pause. A wrong deck or a wrong profile has to
    cost a moment, not a rebuild."""
    def _never(*a, **k):
        raise AssertionError("the master was applied before it was approved")

    monkeypatch.setattr("qc.applymaster.apply_master", _never)
    _uncertain(monkeypatch)

    page = _start(client, profile_with_master)
    assert page.status_code == 200
    assert "Step 2 of 3" in page.text
    assert "slides to place" in page.text
    assert len(web._plans) == 1


def test_the_page_offers_only_the_masters_own_layouts(client,
                                                      profile_with_master,
                                                      no_renderer, monkeypatch):
    _uncertain(monkeypatch)
    page = _start(client, profile_with_master, slides=1).text

    assert 'value="Two Content"' in page
    assert 'name="pick_0"' in page
    # and the answer that is always available
    assert 'value="__leave__"' in page


def test_a_host_with_no_renderer_still_gets_a_usable_page(client,
                                                          profile_with_master,
                                                          no_renderer,
                                                          monkeypatch):
    """The choice is readable from the files alone - what the slide holds and
    what each layout offers are both arithmetic - so a missing renderer costs
    the pictures and nothing else."""
    _uncertain(monkeypatch)
    page = _start(client, profile_with_master, slides=1).text

    assert "could not be rendered" in page
    assert "content box" in page, "what each layout offers is still stated"
    assert 'name="pick_0"' in page, "and the choice is still makeable"


def test_every_slide_matching_skips_the_questions_but_not_the_press(
        client, profile_with_master, no_renderer):
    """Certainty is not a reason to rebuild without being asked. The page states
    what it is about to do and waits."""
    page = _start(client, profile_with_master)
    assert page.status_code == 200
    assert "Every slide matched a layout in this master" in page.text
    # The button wraps in the source, so compare on collapsed whitespace.
    assert "Apply master to 2 slides" in " ".join(page.text.split())


# -------------------------------------------------------- the second press


def test_the_picks_posted_are_the_layouts_applied(client, profile_with_master,
                                                  no_powerpoint, no_renderer,
                                                  monkeypatch):
    _uncertain(monkeypatch)
    _start(client, profile_with_master)
    plan_id = next(reversed(web._plans))

    done = client.post(f"/prep/{plan_id}/layouts",
                       data={"pick_0": "Two Content",
                             "pick_1": "Section Header"})
    assert done.status_code == 200

    applied = {p.slide_index: p.target_layout for p in no_powerpoint["plans"]}
    assert applied == {0: "Two Content", 1: "Section Header"}


def test_the_select_beats_the_radio(client, profile_with_master,
                                    no_powerpoint, no_renderer, monkeypatch):
    """The radio carries a default nobody had to touch; the select is only
    reachable by opening it and choosing."""
    _uncertain(monkeypatch)
    _start(client, profile_with_master, slides=1)
    plan_id = next(reversed(web._plans))

    client.post(f"/prep/{plan_id}/layouts",
                data={"pick_0": "Two Content", "other_0": "Picture with Caption"})
    assert no_powerpoint["plans"][0].target_layout == "Picture with Caption"


def test_leaving_a_slide_rebuilds_it_and_reports_the_gap(client,
                                                         profile_with_master,
                                                         no_powerpoint,
                                                         no_renderer,
                                                         monkeypatch):
    """"None of these fit" is a real answer and the page has to carry it all the
    way through: the slide still rebuilds on whatever the file gave it, and the
    refusal is recorded so the coverage report can count it as a gap in the
    master rather than a slide nobody opened."""
    from qc.layoutpick import LEAVE

    _uncertain(monkeypatch)
    _start(client, profile_with_master, slides=1)
    plan_id = next(reversed(web._plans))

    done = client.post(f"/prep/{plan_id}/layouts", data={"pick_0": LEAVE})
    plan = no_powerpoint["plans"][0]
    assert plan.match_rule == "fallback"
    assert plan.review == "no fit"
    assert "no layout in this master that fits" in done.text


def test_when_nothing_fits_the_default_is_to_say_so(client,
                                                    profile_with_master,
                                                    no_renderer, monkeypatch):
    """A master with no layout that can hold the slide should not have one
    pre-selected anyway. The gap is the answer, and pre-picking a near-miss
    hides it behind a choice nobody made."""
    import qc.layoutpick as LP

    _uncertain(monkeypatch)
    # Every candidate reports that it does not fit.
    real_rank = LP.rank

    def _never_fits(*a, **kw):
        cands, sig = real_rank(*a, **kw)
        for c in cands:
            c.fits = False
        return cands, sig

    monkeypatch.setattr(LP, "rank", _never_fits)

    page = _start(client, profile_with_master, slides=1).text
    leave = page.split(f'value="{LP.LEAVE}"')[1][:40]
    assert "checked" in leave, "the refusal is not the default when it should be"


def test_pressing_apply_with_no_picks_uses_the_suggestions(
        client, profile_with_master, no_powerpoint, no_renderer, monkeypatch):
    _uncertain(monkeypatch)
    _start(client, profile_with_master, slides=1)
    plan_id = next(reversed(web._plans))

    done = client.post(f"/prep/{plan_id}/layouts", data={})
    assert done.status_code == 200
    assert no_powerpoint["plans"][0].target_layout


def test_the_run_says_what_was_decided(client, profile_with_master,
                                       no_powerpoint, no_renderer, monkeypatch):
    """The layout step's own record, on the page that comes out of it. A
    decision nobody can see afterwards is a decision nobody will trust."""
    _uncertain(monkeypatch)
    _start(client, profile_with_master, slides=1)
    plan_id = next(reversed(web._plans))

    done = client.post(f"/prep/{plan_id}/layouts",
                       data={"pick_0": "Section Header"})
    assert "needed a layout decision" in done.text


def test_the_plan_is_dropped_once_it_has_been_spent(client,
                                                    profile_with_master,
                                                    no_powerpoint, no_renderer,
                                                    monkeypatch):
    """Its bytes live on the job afterwards. Keeping both is a second copy of a
    client deck in memory for no reason."""
    _uncertain(monkeypatch)
    _start(client, profile_with_master, slides=1)
    plan_id = next(reversed(web._plans))

    client.post(f"/prep/{plan_id}/layouts", data={})
    assert plan_id not in web._plans


# ------------------------------------------------------------ what goes wrong


def test_an_expired_plan_says_start_again(client):
    page = client.get("/prep/deadbeef/layouts")
    assert page.status_code == 404
    assert "expired" in page.text

    posted = client.post("/prep/deadbeef/layouts", data={})
    assert posted.status_code == 404
    assert "expired" in posted.text


def test_an_image_for_a_plan_that_is_gone_is_a_404_not_a_crash(client):
    assert client.get("/plan-img/deadbeef/slide-0.png").status_code == 404


def test_a_junk_image_key_is_refused(client, profile_with_master, no_renderer,
                                     monkeypatch):
    _uncertain(monkeypatch)
    _start(client, profile_with_master, slides=1)
    plan_id = next(reversed(web._plans))

    for key in ("slide-nope", "layout-99", "whatever"):
        assert client.get(f"/plan-img/{plan_id}/{key}.png").status_code == 404


def test_the_layout_page_can_be_reloaded(client, profile_with_master,
                                         no_renderer, monkeypatch):
    """A designer sits on this page. It has to survive a refresh, and it must
    not re-render the slides to do it."""
    _uncertain(monkeypatch)
    _start(client, profile_with_master, slides=1)
    plan_id = next(reversed(web._plans))

    again = client.get(f"/prep/{plan_id}/layouts")
    assert again.status_code == 200
    assert "Step 2 of 3" in again.text
