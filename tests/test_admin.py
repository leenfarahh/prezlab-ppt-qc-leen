"""Admin surfaces: team roster and profile editor (role-gated saves)."""

import json
import shutil

import pytest
from fastapi.testclient import TestClient

import qc.store as store_mod
from qc.web import app

client = TestClient(app)


@pytest.fixture()
def store(monkeypatch, tmp_path):
    monkeypatch.setattr(store_mod, "DATA_DIR", tmp_path)
    monkeypatch.setattr(store_mod, "DB_PATH", tmp_path / "qc.db")
    return store_mod


@pytest.fixture()
def profiles_dir(monkeypatch, tmp_path):
    import qc.profile as profile_mod

    src = profile_mod.PROFILES_DIR / "prezlab_en.json"
    pdir = tmp_path / "profiles"
    pdir.mkdir()
    shutil.copy(src, pdir / "prezlab_en.json")
    monkeypatch.setattr(profile_mod, "PROFILES_DIR", pdir)
    return pdir


def test_team_page_and_add_user(store):
    r = client.get("/team")
    assert r.status_code == 200 and "Team" in r.text
    r = client.post("/team/add", data={"name": "Razan", "role": "lead"},
                    follow_redirects=False)
    assert r.status_code == 303
    assert "Razan" in client.get("/team").text


def test_profiles_list(profiles_dir):
    r = client.get("/profiles")
    assert r.status_code == 200
    assert "prezlab_en" in r.text


def test_edit_page_open_to_all(profiles_dir):
    assert client.get("/profiles/prezlab_en/edit").status_code == 200


def test_save_requires_lead_or_admin(store, profiles_dir):
    # anonymous
    r = client.post("/profiles/prezlab_en/edit", data={"name": "x"})
    assert r.status_code == 403
    # designer role: still forbidden
    store.add_user("Dana", "designer")
    c2 = TestClient(app, cookies={"qc_user": "Dana"})
    assert c2.post("/profiles/prezlab_en/edit", data={"name": "x"}).status_code == 403


def _lead_client(store):
    store.add_user("Razan", "lead")
    return TestClient(app, cookies={"qc_user": "Razan"})


def test_save_as_lead_bumps_version(store, profiles_dir):
    c = _lead_client(store)
    before = json.loads((profiles_dir / "prezlab_en.json").read_text(encoding="utf-8"))
    r = c.post("/profiles/prezlab_en/edit",
               data={"name": "Prezlab EN edited", "palette": "navy 1F4E79",
                     "title_latin": "Georgia", "footer_text": ""},
               follow_redirects=False)
    assert r.status_code == 303
    after = json.loads((profiles_dir / "prezlab_en.json").read_text(encoding="utf-8"))
    assert after["version"] == before["version"] + 1
    assert after["name"] == "Prezlab EN edited"
    assert "Razan" in after["owner"]
    assert after["config"]["color_palette"]["named_colors"] == [
        {"name": "navy", "hex": "1F4E79", "theme_ref": None,
         "allowed_tints": [], "allowed_shades": []}]


def test_invalid_hex_rejected(store, profiles_dir):
    c = _lead_client(store)
    r = c.post("/profiles/prezlab_en/edit",
               data={"name": "x", "palette": "navy notahex"})
    assert r.status_code == 400
    assert "hex" in r.text


def test_raw_json_override(store, profiles_dir):
    c = _lead_client(store)
    raw = json.dumps({"font": {"roles": {}}, "color_palette": {"named_colors": []}})
    r = c.post("/profiles/prezlab_en/edit", data={"name": "x", "raw": raw},
               follow_redirects=False)
    assert r.status_code == 303
    after = json.loads((profiles_dir / "prezlab_en.json").read_text(encoding="utf-8"))
    assert after["config"] == json.loads(raw)


def test_invalid_raw_json_rejected(store, profiles_dir):
    c = _lead_client(store)
    r = c.post("/profiles/prezlab_en/edit", data={"name": "x", "raw": "{nope"})
    assert r.status_code == 400


def test_unknown_profile_404(profiles_dir):
    assert client.get("/profiles/nope/edit").status_code == 404


def test_create_profile_from_reference(store, profiles_dir, fixtures_dir):
    c = _lead_client(store)
    deck = (fixtures_dir / "mixed_layouts.pptx").read_bytes()
    r = c.post("/profiles/new",
               files={"deck": ("Client ABC ref.pptx", deck, "application/octet-stream")},
               data={"name": "Client ABC"}, follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"] == "/profiles/client_abc/edit"
    created = json.loads((profiles_dir / "client_abc.json").read_text(encoding="utf-8"))
    assert created["name"] == "Client ABC"
    assert "Razan" in created["owner"] and "reference" in created["owner"]
    # bootstrap actually learned something from the deck
    assert created["config"]["font"]["roles"]["title"]["latin"]


def test_create_profile_requires_lead(store, profiles_dir, fixtures_dir):
    deck = (fixtures_dir / "clean.pptx").read_bytes()
    # anonymous
    r = client.post("/profiles/new",
                    files={"deck": ("r.pptx", deck, "application/octet-stream")},
                    data={"name": "X"})
    assert r.status_code == 403
    # designer
    store.add_user("Dana", "designer")
    dana = TestClient(app, cookies={"qc_user": "Dana"})
    r = dana.post("/profiles/new",
                  files={"deck": ("r.pptx", deck, "application/octet-stream")},
                  data={"name": "X"})
    assert r.status_code == 403


def test_create_profile_rejects_non_pptx(store, profiles_dir):
    c = _lead_client(store)
    r = c.post("/profiles/new",
               files={"deck": ("notes.txt", b"hi", "text/plain")},
               data={"name": "X"})
    assert r.status_code == 400


def test_create_profile_unique_slug(store, profiles_dir, fixtures_dir):
    c = _lead_client(store)
    deck = (fixtures_dir / "clean.pptx").read_bytes()
    for _ in range(2):
        c.post("/profiles/new",
               files={"deck": ("r.pptx", deck, "application/octet-stream")},
               data={"name": "Repeat Client"}, follow_redirects=False)
    assert (profiles_dir / "repeat_client.json").exists()
    assert (profiles_dir / "repeat_client_2.json").exists()
