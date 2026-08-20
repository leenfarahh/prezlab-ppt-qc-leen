"""Stage 1 in the web tier: submit a master, review the spec, save it as a
profile the audit engine can read."""

import io
import json

from fastapi.testclient import TestClient
from pptx import Presentation

from qc import web
from qc.profile import Profile
from qc.stylespec import extract_style_spec, spec_to_profile


def _master_bytes(**kw) -> bytes:
    """A master-only file: the Stage 1 submission shape, no slides."""
    prs = Presentation(**kw)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _client(tmp_path, monkeypatch):
    monkeypatch.setattr(web, "AUTH_REQUIRED", False)
    web.app.state.auth_required = False
    return TestClient(web.app)


def _as_lead(client):
    from qc.store import add_user

    add_user("Lead Person", "lead")
    client.post("/whoami", json={"name": "Lead Person"})


# ------------------------------------------------------------------- routes


def test_master_page_renders(tmp_path, monkeypatch):
    r = _client(tmp_path, monkeypatch).get("/master")
    assert r.status_code == 200
    assert "Read a master" in r.text


def test_reading_a_master_shows_the_spec(tmp_path, monkeypatch):
    client = _client(tmp_path, monkeypatch)
    r = client.post("/master", files={"master": ("brand.pptx", _master_bytes(),
                                                 "application/vnd.ms-powerpoint")})
    assert r.status_code == 200
    assert "brand.pptx" in r.text
    # The theme is shown by role, with real swatches.
    assert "accent1" in r.text
    # No slides is the normal case and the page says so rather than warning.
    assert "carries no slides" in r.text
    # Layout archetypes are surfaced, since they drive Stage 2 matching.
    assert "secHead" in r.text


def test_non_pptx_is_refused(tmp_path, monkeypatch):
    client = _client(tmp_path, monkeypatch)
    r = client.post("/master", files={"master": ("notes.txt", b"nope", "text/plain")})
    assert r.status_code == 400
    assert "Only .pptx" in r.text


def test_unreadable_file_fails_with_a_message_not_a_traceback(tmp_path, monkeypatch):
    client = _client(tmp_path, monkeypatch)
    r = client.post("/master", files={"master": ("broken.pptx", b"PK\x03\x04junk",
                                                 "application/vnd.ms-powerpoint")})
    assert r.status_code == 422
    assert "Could not read that master" in r.text


def test_spec_json_downloads_and_round_trips(tmp_path, monkeypatch):
    client = _client(tmp_path, monkeypatch)
    client.post("/master", files={"master": ("brand.pptx", _master_bytes(), "app/x")})
    spec_id = next(iter(web._specs))

    r = client.get(f"/spec/{spec_id}.json")
    assert r.status_code == 200
    assert "brand-stylespec.json" in r.headers["content-disposition"]
    spec = r.json()
    assert spec["spec_version"] == 1
    assert spec["theme"]["colors"]["accent1"]


def test_unknown_spec_id_is_a_clean_404(tmp_path, monkeypatch):
    r = _client(tmp_path, monkeypatch).get("/spec/deadbeef.json")
    assert r.status_code == 404


# ------------------------------------------------------------ saving a spec


def test_saving_a_profile_needs_a_lead(tmp_path, monkeypatch):
    client = _client(tmp_path, monkeypatch)
    client.post("/master", files={"master": ("brand.pptx", _master_bytes(), "app/x")})
    spec_id = next(iter(web._specs))

    r = client.post(f"/spec/{spec_id}/profile", data={"name": "Client X"},
                    follow_redirects=False)
    assert r.status_code == 403
    assert "lead or admin" in r.text


def test_lead_saves_a_spec_as_a_usable_profile(tmp_path, monkeypatch):
    from qc.profile import PROFILES_DIR

    client = _client(tmp_path, monkeypatch)
    _as_lead(client)
    client.post("/master", files={"master": ("brand.pptx", _master_bytes(), "app/x")})
    spec_id = next(iter(web._specs))

    r = client.post(f"/spec/{spec_id}/profile", data={"name": "Client Zed master"},
                    follow_redirects=False)
    assert r.status_code == 303
    pid = r.headers["location"].split("/")[2]
    written = PROFILES_DIR / f"{pid}.json"
    try:
        assert written.exists()
        # It has to load through the ordinary profile path, or the audit
        # engine cannot use it.
        profile = Profile.load(pid)
        assert profile.get("font.roles.title.latin")
        assert profile.get("color_palette.named_colors")
        assert profile.get("theme.colors.accent1")
    finally:
        written.unlink(missing_ok=True)


# --------------------------------------------- replacing a stored master
#
# A profile's master is stored once, when the profile is created, and the format
# step hands PowerPoint that stored copy. Everything the designer adds to their
# master afterwards - a presentation-space rectangle, a moved guide, a new
# layout - reached no deck at all before this existed (design lead, 21/08/2026:
# the presentation-space box was missing from every formatted deck).


def _master_with_space(left_in=1.2, top_in=0.4, w_in=10.5, h_in=6.4) -> bytes:
    from pptx.oxml.shapes.autoshape import CT_Shape

    IN = 914400
    prs = Presentation()
    prs.slide_width, prs.slide_height = 12192000, 6858000
    master = prs.slide_masters[0]
    sp = CT_Shape.new_autoshape_sp(950, "Presentation space", "rect",
                                   int(left_in * IN), int(top_in * IN),
                                   int(w_in * IN), int(h_in * IN))
    master.shapes._spTree.insert_element_before(sp, "p:extLst")
    shape = next(s for s in master.shapes if s.name == "Presentation space")
    shape.fill.background()
    shape.line.fill.background()
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _saved_profile(client) -> str:
    client.post("/master", files={"master": ("brand.pptx", _master_bytes(),
                                            "app/x")})
    spec_id = next(iter(web._specs))
    r = client.post(f"/spec/{spec_id}/profile", data={"name": "Frame Client"},
                    follow_redirects=False)
    return r.headers["location"].split("/")[2]


def test_replacing_the_master_stores_the_new_file_and_re_reads_its_frame(
        tmp_path, monkeypatch):
    from qc.profile import PROFILES_DIR
    from qc.templates import load_master

    client = _client(tmp_path, monkeypatch)
    _as_lead(client)
    pid = _saved_profile(client)
    try:
        before = Profile.load(pid).get("geometry.safe_zone_margins_emu")

        new_master = _master_with_space()
        r = client.post(f"/profiles/{pid}/master",
                        files={"master": ("brand-v2.pptx", new_master, "app/x")})
        assert r.status_code == 200
        assert "Master replaced" in r.text
        assert "presentation space" in r.text

        # the FILE is what gets applied, so that is what must have changed
        assert load_master(pid) == new_master
        after = Profile.load(pid).get("geometry.safe_zone_margins_emu")
        assert after != before
        assert after["left"] == int(1.2 * 914400)
        assert Profile.load(pid).get("style_spec_source.grid_source") == \
            "presentation_space"
    finally:
        (PROFILES_DIR / f"{pid}.json").unlink(missing_ok=True)


def test_replacing_the_master_keeps_hand_edited_rules(tmp_path, monkeypatch):
    """Fonts, palette and tolerances are decisions about the client, not
    readings of the file. Reverting them to a fresh projection would punish
    every edit ever made in the editor."""
    from qc.profile import PROFILES_DIR

    client = _client(tmp_path, monkeypatch)
    _as_lead(client)
    pid = _saved_profile(client)
    path = PROFILES_DIR / f"{pid}.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        data["config"]["geometry"]["alignment"]["edge_tolerance_emu"] = 28575
        data["config"]["font"]["roles"]["title"]["latin"] = ["Prezlab Display"]
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")

        client.post(f"/profiles/{pid}/master",
                    files={"master": ("brand-v2.pptx", _master_with_space(),
                                      "app/x")})
        after = Profile.load(pid)
        assert after.get("geometry.alignment.edge_tolerance_emu") == 28575
        assert after.get("font.roles.title.latin") == ["Prezlab Display"]
        assert after.version > 1, "the version must move: the rules changed"
    finally:
        path.unlink(missing_ok=True)


def test_replacing_the_master_needs_a_lead(tmp_path, monkeypatch):
    from qc.profile import PROFILES_DIR
    from qc.templates import load_master

    client = _client(tmp_path, monkeypatch)
    _as_lead(client)
    pid = _saved_profile(client)
    path = PROFILES_DIR / f"{pid}.json"
    try:
        stored = load_master(pid)
        anonymous = _client(tmp_path, monkeypatch)   # its own cookie jar
        r = anonymous.post(f"/profiles/{pid}/master",
                           files={"master": ("brand-v2.pptx",
                                             _master_with_space(), "app/x")})
        assert r.status_code == 403
        assert load_master(pid) == stored, "the stored file must be untouched"
    finally:
        path.unlink(missing_ok=True)


def test_a_broken_master_upload_is_refused_without_losing_the_old_one(
        tmp_path, monkeypatch):
    from qc.profile import PROFILES_DIR
    from qc.templates import load_master

    client = _client(tmp_path, monkeypatch)
    _as_lead(client)
    pid = _saved_profile(client)
    path = PROFILES_DIR / f"{pid}.json"
    try:
        stored = load_master(pid)
        r = client.post(f"/profiles/{pid}/master",
                        files={"master": ("junk.pptx", b"not a pptx", "app/x")})
        assert r.status_code == 422
        assert "Could not read that master" in r.text
        assert load_master(pid) == stored
    finally:
        path.unlink(missing_ok=True)


def test_expired_spec_says_so_instead_of_500(tmp_path, monkeypatch):
    client = _client(tmp_path, monkeypatch)
    _as_lead(client)
    r = client.post("/spec/gone/profile", data={"name": "X"},
                    follow_redirects=False)
    assert r.status_code == 404
    assert "expired" in r.text


# --------------------------------------------------------- the spec -> profile


def test_profile_takes_major_font_for_headings_and_minor_for_body():
    spec = extract_style_spec(Presentation(), source="m.pptx")
    theme = spec["theme"]["fonts"]
    cfg = spec_to_profile(spec, "p", "P")["config"]

    assert theme["major"]["latin"] in cfg["font"]["roles"]["title"]["latin"]
    assert theme["minor"]["latin"] in cfg["font"]["roles"]["body"]["latin"]


def test_palette_entries_carry_their_theme_slot():
    """The whole point of sourcing the palette from the theme: a later theme
    edit has to carry the palette with it."""
    spec = extract_style_spec(Presentation(), source="m.pptx")
    named = spec_to_profile(spec, "p", "P")["config"]["color_palette"]["named_colors"]

    assert named
    assert all(c["theme_ref"] == c["name"] for c in named)
    assert not any(c["name"] in ("hlink", "folHlink") for c in named)


def test_grid_columns_enable_the_profile_grid():
    spec = extract_style_spec(Presentation(), source="m.pptx")
    spec["grid"] = {"guides": {"vertical_emu": [], "horizontal_emu": []},
                    "margins_emu": {"left": 1, "right": 2, "top": 3, "bottom": 4},
                    "columns": 6, "gutter_emu": 304800, "source": "guides"}
    geo = spec_to_profile(spec, "p", "P")["config"]["geometry"]

    assert geo["safe_zone_margins_emu"]["left"] == 1
    assert geo["grid"] == {"columns": 6, "gutter_emu": 304800, "enabled": True}


def test_profile_falls_back_to_canvas_ratios_without_a_grid():
    spec = extract_style_spec(Presentation(), source="m.pptx")
    spec["grid"] = {"guides": {"vertical_emu": [], "horizontal_emu": []},
                    "margins_emu": None, "columns": None,
                    "gutter_emu": None, "source": None}
    geo = spec_to_profile(spec, "p", "P")["config"]["geometry"]

    assert geo["safe_zone_margins_emu"]["left"] > 0
    assert geo["grid"]["enabled"] is False


def test_profile_is_json_serialisable():
    spec = extract_style_spec(Presentation(), source="m.pptx")
    profile = spec_to_profile(spec, "p", "P")
    assert json.loads(json.dumps(profile)) == profile
