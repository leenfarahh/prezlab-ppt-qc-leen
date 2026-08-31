"""Editing a profile without re-reading a master.

The page exists because the alternative was a full rebuild for a one-field
change, so the claims that matter are about not losing anything:

  - A SAVE THAT CHANGES NOTHING CHANGES NOTHING. Loading a profile into the
    form and saving it back must leave every field this module owns exactly as
    it was, units and all. That is the round trip, and it is what makes the
    inches-and-points display safe: EMU in, inches on screen, the same EMU out.
  - keys the form does not own survive it, because a profile carries things no
    input should be able to reach;
  - a bad value is REFUSED and named, with the rest of the edit still on screen;
  - and saving is a lead-or-admin action, checked at the route rather than
    hidden in the page.
"""

import json

import pytest
from fastapi.testclient import TestClient

from qc import web
from qc import web_admin
from qc.profileform import EMU_IN, EMU_PT, Invalid, apply_form, display, groups


# ------------------------------------------------------------------ helpers


class _Form(dict):
    """A form the parser can read. Starlette's FormData is a multidict and the
    weight/palette fields post repeated keys, so `getlist` is part of the
    contract this stands in for."""

    def __init__(self, pairs):
        super().__init__()
        self._pairs = list(pairs)
        for key, value in self._pairs:
            self.setdefault(key, value)

    def get(self, key, default=None):
        return super().get(key, default)

    def getlist(self, key):
        return [v for k, v in self._pairs if k == key]


def _form_from(profile: dict):
    """The form a browser would post from an untouched render of `profile`.
    Built off the same FIELDS the page renders, so it cannot drift from it."""
    pairs = []
    for group in groups():
        for f in group.fields:
            value = display(profile, f)
            if f.kind == "bool":
                if value:
                    pairs.append((f.name, "1"))
            elif isinstance(value, list):
                pairs.extend((f.name, v) for v in value)
            else:
                pairs.append((f.name, str(value)))
    for color in ((profile.get("config") or {})
                  .get("color_palette") or {}).get("named_colors") or []:
        pairs.append(("color_name", color["name"]))
        pairs.append(("color_hex", color["hex"]))
    return _Form(pairs)


def _payload(pairs) -> dict:
    """The same pairs as httpx will post them.

    httpx wants repeated form keys as {key: [v1, v2]}; handed a list of tuples
    it encodes the whole thing as a raw body and the server sees no fields at
    all - which shows up as a validation error about a field the test never
    touched."""
    out: dict = {}
    for key, value in pairs:
        if key in out:
            out[key] = (out[key] if isinstance(out[key], list)
                        else [out[key]]) + [value]
        else:
            out[key] = value
    return out


@pytest.fixture()
def bilingual():
    from qc.profile import PROFILES_DIR

    return json.loads((PROFILES_DIR / "prezlab_bilingual.json")
                      .read_text(encoding="utf-8"))


@pytest.fixture()
def client(monkeypatch):
    monkeypatch.setattr(web, "AUTH_REQUIRED", False)
    monkeypatch.setattr(web.app.state, "auth_required", False)
    return TestClient(web.app)


@pytest.fixture()
def as_lead(monkeypatch):
    monkeypatch.setattr(web_admin, "_editor",
                        lambda request: {"name": "Leen", "role": "lead"})


@pytest.fixture()
def as_viewer(monkeypatch):
    monkeypatch.setattr(web_admin, "_editor", lambda request: None)


# ------------------------------------------------------------ the round trip


def test_a_save_that_changes_nothing_changes_nothing(bilingual):
    """The property the whole units layer rests on. EMU into the form, inches
    on screen, the same EMU back out - or every open-and-save quietly drifts a
    margin by a rounding error."""
    saved = apply_form(bilingual, _form_from(bilingual))

    before, after = bilingual["config"], saved["config"]
    assert after["geometry"] == before["geometry"]
    assert after["font"] == before["font"]
    assert after["color_palette"] == before["color_palette"]
    assert after["shape_size"] == before["shape_size"]
    assert after["master_slide"] == before["master_slide"]
    assert after["header_footer"] == before["header_footer"]


def test_saving_bumps_the_version_and_nothing_else_at_the_top(bilingual):
    saved = apply_form(bilingual, _form_from(bilingual))
    assert saved["version"] == bilingual["version"] + 1
    assert saved["id"] == bilingual["id"]
    assert saved["name"] == bilingual["name"]


def test_keys_the_form_does_not_own_survive_a_save(bilingual):
    """A profile carries things no input reaches - the id, whatever a later
    version adds. Rewriting the document from the form alone deletes them."""
    bilingual["config"]["typography"] = {"widow_control": True}
    bilingual["provenance"] = {"read_from": "brand-v4.pptx"}

    saved = apply_form(bilingual, _form_from(bilingual))
    assert saved["config"]["typography"] == {"widow_control": True}
    assert saved["provenance"] == {"read_from": "brand-v4.pptx"}


def test_the_original_is_not_edited_in_place(bilingual):
    """A save that fails validation halfway must not leave the loaded profile
    half-written; the page re-renders from it."""
    before = json.dumps(bilingual, sort_keys=True)
    apply_form(bilingual, _form_from(bilingual))
    assert json.dumps(bilingual, sort_keys=True) == before


# ---------------------------------------------------------------- the units


def test_margins_are_edited_in_inches_and_stored_in_emu(bilingual):
    form = _form_from(bilingual)
    form._pairs = [("geometry__safe_zone_margins_emu__left", "0.75")
                   if k == "geometry__safe_zone_margins_emu__left" else (k, v)
                   for k, v in form._pairs]
    form["geometry__safe_zone_margins_emu__left"] = "0.75"

    saved = apply_form(bilingual, form)
    assert saved["config"]["geometry"]["safe_zone_margins_emu"]["left"] == \
        int(0.75 * EMU_IN)


def test_tolerances_are_edited_in_points_and_stored_in_emu(bilingual):
    form = _form_from(bilingual)
    form["geometry__alignment__edge_tolerance_emu"] = "2"

    saved = apply_form(bilingual, form)
    assert saved["config"]["geometry"]["alignment"]["edge_tolerance_emu"] == \
        2 * EMU_PT


# ------------------------------------------------------------- what it refuses


def test_a_hex_code_that_is_not_one_is_refused_and_named(bilingual):
    form = _form_from(bilingual)
    form._pairs = [(k, "ZZTOP!" if k == "color_hex" else v)
                   for k, v in form._pairs]

    with pytest.raises(Invalid) as exc:
        apply_form(bilingual, form)
    assert exc.value.field_name == "color_hex"
    assert "six hex digits" in str(exc.value)


def test_a_colour_with_no_name_is_refused(bilingual):
    form = _form_from(bilingual)
    form._pairs.append(("color_name", ""))
    form._pairs.append(("color_hex", "ABCDEF"))

    with pytest.raises(Invalid) as exc:
        apply_form(bilingual, form)
    assert exc.value.field_name == "color_name"


def test_two_colours_with_the_same_name_are_refused(bilingual):
    """A name is how a fix picks one. Two of them makes the pick arbitrary."""
    form = _form_from(bilingual)
    form._pairs.append(("color_name", "prezlab_navy"))
    form._pairs.append(("color_hex", "123456"))

    with pytest.raises(Invalid) as exc:
        apply_form(bilingual, form)
    assert "both called" in str(exc.value)


def test_an_emptied_colour_row_is_dropped_rather_than_refused(bilingual):
    form = _form_from(bilingual)
    form._pairs.append(("color_name", "  "))
    form._pairs.append(("color_hex", ""))

    saved = apply_form(bilingual, form)
    assert len(saved["config"]["color_palette"]["named_colors"]) == \
        len(bilingual["config"]["color_palette"]["named_colors"])


def test_a_size_that_is_not_a_number_is_refused_and_named(bilingual):
    form = _form_from(bilingual)
    form["font__roles__body__size_pt"] = "biggish"

    with pytest.raises(Invalid) as exc:
        apply_form(bilingual, form)
    assert exc.value.field_name == "font__roles__body__size_pt"
    assert "has to be a number" in str(exc.value)


def test_a_value_outside_its_bounds_is_refused(bilingual):
    form = _form_from(bilingual)
    form["font__roles__body__size_pt"] = "0"

    with pytest.raises(Invalid) as exc:
        apply_form(bilingual, form)
    assert "cannot be below" in str(exc.value)


def test_a_role_with_no_weights_is_refused(bilingual):
    """An empty weight set allows nothing, which silently makes every run on
    that role an error. Almost certainly not what anyone meant."""
    form = _form_from(bilingual)
    form._pairs = [(k, v) for k, v in form._pairs
                   if k != "font__roles__body__allowed_weights"]

    with pytest.raises(Invalid) as exc:
        apply_form(bilingual, form)
    assert "at least one weight" in str(exc.value)


def test_a_profile_with_no_name_is_refused(bilingual):
    form = _form_from(bilingual)
    form["top:name"] = "   "

    with pytest.raises(Invalid) as exc:
        apply_form(bilingual, form)
    assert "needs a name" in str(exc.value)


def test_a_refused_save_keeps_the_rest_of_the_edit(bilingual):
    """Nine good changes and one typo must come back as nine good changes and
    one typo, not as the file on disk."""
    from qc.profileform import partial

    form = _form_from(bilingual)
    form["font__roles__body__size_pt"] = "17"
    form["geometry__grid__columns"] = "not a number"

    as_typed = partial(bilingual, form)
    assert as_typed["config"]["font"]["roles"]["body"]["size_pt"] == 17
    assert as_typed["config"]["geometry"]["grid"]["columns"] == \
        bilingual["config"]["geometry"]["grid"]["columns"]


# ----------------------------------------------------------------- the routes


def test_the_list_page_names_every_profile(client, as_viewer):
    page = client.get("/profiles")
    assert page.status_code == 200
    assert "Prezlab default (bilingual EN/AR)" in page.text
    assert "prezlab_bilingual" in page.text


def test_the_editor_shows_the_values_in_the_units_it_says(client, as_lead):
    page = client.get("/profiles/prezlab_bilingual")
    assert page.status_code == 200
    assert "Left margin (in)" in page.text
    # 457200 EMU is exactly half an inch, and that is what has to be on screen.
    assert 'value="0.5"' in page.text
    assert "Georgia" in page.text
    assert "1F4E79" in page.text


def test_an_unknown_profile_is_a_404_not_a_stack_trace(client, as_lead):
    page = client.get("/profiles/no_such_client")
    assert page.status_code == 404
    assert "no profile with the id" in page.text


def test_a_viewer_gets_the_form_read_only(client, as_viewer):
    page = client.get("/profiles/prezlab_bilingual")
    assert page.status_code == 200
    assert "Read-only" in page.text
    assert "Save profile" not in page.text


def test_a_viewer_cannot_save(client, as_viewer):
    reply = client.post("/profiles/prezlab_bilingual",
                        data={"top:name": "Mine"})
    assert reply.status_code == 403
    assert "lead or admin" in reply.text


def test_a_lead_can_change_one_field_and_the_rest_survives(client, as_lead,
                                                           bilingual):
    from qc.profile import PROFILES_DIR

    payload = _payload((k, "17" if k == "font__roles__body__size_pt" else v)
                       for k, v in _form_from(bilingual)._pairs)

    reply = client.post("/profiles/prezlab_bilingual", data=payload,
                        follow_redirects=False)
    assert reply.status_code == 303

    saved = json.loads((PROFILES_DIR / "prezlab_bilingual.json")
                       .read_text(encoding="utf-8"))
    assert saved["config"]["font"]["roles"]["body"]["size_pt"] == 17
    assert saved["config"]["geometry"] == bilingual["config"]["geometry"]
    assert saved["owner"] == "Leen"
    assert saved["version"] == bilingual["version"] + 1


def test_a_bad_save_is_a_400_that_keeps_the_form(client, as_lead, bilingual):
    from qc.profile import PROFILES_DIR

    payload = _payload((k, "nope" if k == "font__roles__body__size_pt" else v)
                       for k, v in _form_from(bilingual)._pairs)
    reply = client.post("/profiles/prezlab_bilingual", data=payload)

    assert reply.status_code == 400
    assert "has to be a number" in reply.text
    assert "Nothing was saved" in reply.text
    unchanged = json.loads((PROFILES_DIR / "prezlab_bilingual.json")
                           .read_text(encoding="utf-8"))
    assert unchanged["version"] == bilingual["version"]


def test_only_one_profile_can_be_the_default(client, as_lead, bilingual):
    """Two defaults is a picker that chooses arbitrarily, so the save clears
    the others rather than trusting whoever set them."""
    from qc.profile import PROFILES_DIR

    other = PROFILES_DIR / "prezlab_en.json"
    seeded = json.loads(other.read_text(encoding="utf-8"))
    seeded["is_default"] = True
    other.write_text(json.dumps(seeded), encoding="utf-8")

    payload = _payload(list(_form_from(bilingual)._pairs)
                       + [("top:is_default", "1")])
    client.post("/profiles/prezlab_bilingual", data=payload,
                follow_redirects=False)

    assert json.loads(other.read_text(encoding="utf-8"))["is_default"] is False
    assert json.loads((PROFILES_DIR / "prezlab_bilingual.json")
                      .read_text(encoding="utf-8"))["is_default"] is True


def test_duplicating_makes_a_new_id_that_is_not_the_default(client, as_lead):
    from qc.profile import PROFILES_DIR

    reply = client.post("/profiles/prezlab_bilingual/duplicate",
                        follow_redirects=False)
    assert reply.status_code == 303
    new_id = reply.headers["location"].rsplit("/", 1)[-1]
    assert new_id != "prezlab_bilingual"

    copy = json.loads((PROFILES_DIR / f"{new_id}.json")
                      .read_text(encoding="utf-8"))
    assert copy["id"] == new_id
    assert copy["name"].endswith("(copy)")
    assert copy["is_default"] is False
    assert copy["version"] == 1


def test_deleting_removes_the_profile(client, as_lead):
    from qc.profile import PROFILES_DIR

    reply = client.post("/profiles/prezlab_en/delete", follow_redirects=False)
    assert reply.status_code == 303
    assert not (PROFILES_DIR / "prezlab_en.json").exists()


def test_a_viewer_cannot_delete(client, as_viewer):
    from qc.profile import PROFILES_DIR

    reply = client.post("/profiles/prezlab_en/delete")
    assert reply.status_code == 403
    assert (PROFILES_DIR / "prezlab_en.json").exists()


def test_the_nav_reaches_the_profiles_from_anywhere(client):
    """The page is worthless if it is not linked. This is the whole complaint
    the rebuild started from: pages nobody could find from the page they were
    on."""
    for path in ("/", "/prep", "/profiles"):
        page = client.get(path)
        assert 'href="/profiles"' in page.text, f"no way to profiles from {path}"
