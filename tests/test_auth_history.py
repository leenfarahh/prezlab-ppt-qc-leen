"""Sign-in (PIN + sessions) and audit history."""

import pytest
from fastapi.testclient import TestClient

import qc.store as store_mod
from qc.web import app


@pytest.fixture()
def store(monkeypatch, tmp_path):
    monkeypatch.setattr(store_mod, "DATA_DIR", tmp_path)
    monkeypatch.setattr(store_mod, "DB_PATH", tmp_path / "qc.db")
    return store_mod


def test_first_signin_sets_pin_and_session(store):
    store.add_user("Sanad", "lead")
    client = TestClient(app)
    r = client.post("/signin", data={"name": "Sanad", "pin": "4321"},
                    follow_redirects=False)
    assert r.status_code == 303
    assert "qc_session" in r.cookies
    assert store.has_pin("Sanad")
    me = client.get("/me").json()
    assert me["user"]["name"] == "Sanad"
    assert "pin_hash" not in me["user"]


def test_wrong_pin_rejected(store):
    store.add_user("Sanad", "lead")
    client = TestClient(app)
    client.post("/signin", data={"name": "Sanad", "pin": "4321"})
    r = TestClient(app).post("/signin", data={"name": "Sanad", "pin": "9999"})
    assert r.status_code == 401
    assert "Wrong PIN" in r.text


def test_short_pin_rejected(store):
    store.add_user("Sanad", "lead")
    r = TestClient(app).post("/signin", data={"name": "Sanad", "pin": "12"})
    assert r.status_code == 400


def test_signout_kills_session(store):
    store.add_user("Sanad", "lead")
    client = TestClient(app)
    client.post("/signin", data={"name": "Sanad", "pin": "4321"})
    assert client.get("/me").json()["user"] is not None
    client.post("/signout")
    assert client.get("/me").json()["user"] is None


def test_signin_page_renders(store):
    r = TestClient(app).get("/signin")
    assert r.status_code == 200 and "Who is working?" in r.text


def test_audit_recorded_in_history(store, fixtures_dir):
    client = TestClient(app)
    deck = (fixtures_dir / "mixed_layouts.pptx").read_bytes()
    r = client.post("/audit",
                    files={"deck": ("mixed_layouts.pptx", deck, "application/octet-stream")},
                    data={"profile": "prezlab_en"})
    assert r.status_code == 200
    audits = store.list_audits()
    assert len(audits) == 1
    assert audits[0]["deck"] == "mixed_layouts.pptx"
    assert audits[0]["user_name"] == "anonymous"
    assert audits[0]["errors"] > 0

    h = client.get("/history")
    assert h.status_code == 200 and "mixed_layouts.pptx" in h.text

    view = client.get(f"/history/{audits[0]['id']}")
    assert view.status_code == 200
    assert "Archived audit" in view.text
    assert "font.family_out_of_set" in view.text
    # archived views are read-only: no apply form, no live exports
    assert 'id="applyform"' not in view.text
    assert "/report/" not in view.text


def test_history_attributes_signed_in_user(store, fixtures_dir):
    store.add_user("Sanad", "lead")
    client = TestClient(app)
    client.post("/signin", data={"name": "Sanad", "pin": "4321"})
    deck = (fixtures_dir / "clean.pptx").read_bytes()
    client.post("/audit",
                files={"deck": ("clean.pptx", deck, "application/octet-stream")},
                data={"profile": "prezlab_en"})
    assert store.list_audits()[0]["user_name"] == "Sanad"


def test_unknown_history_id_404(store):
    assert TestClient(app).get("/history/99999").status_code == 404


def test_lan_mode_enforces_signin(store, monkeypatch, fixtures_dir):
    import qc.auth as auth_mod
    import qc.web as web_mod

    monkeypatch.setattr(web_mod, "AUTH_REQUIRED", True)
    monkeypatch.setattr(auth_mod, "STRICT_SESSIONS", True)
    client = TestClient(app)

    # anonymous GET redirects to sign-in; public paths stay open
    r = client.get("/", follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"] == "/signin"
    assert client.get("/health").status_code == 200
    assert client.get("/signin").status_code == 200

    # anonymous POST is a clean 401, not a redirect
    assert client.post("/triage", json={}).status_code == 401

    # the legacy name cookie no longer counts in strict mode
    store.add_user("Sanad", "lead")
    spoof = TestClient(app, cookies={"qc_user": "Sanad"})
    assert spoof.get("/", follow_redirects=False).status_code == 303

    # a real session gets through
    signed = TestClient(app)
    signed.post("/signin", data={"name": "Sanad", "pin": "4321"})
    assert signed.get("/", follow_redirects=False).status_code == 200


def test_pin_reset_by_lead(store):
    store.add_user("Razan", "lead")
    store.add_user("Dana", "designer")
    dana = TestClient(app)
    dana.post("/signin", data={"name": "Dana", "pin": "1111"})
    assert store.has_pin("Dana")

    # designers cannot reset PINs
    r = dana.post("/team/reset-pin", data={"name": "Razan"})
    assert r.status_code == 403

    razan = TestClient(app)
    razan.post("/signin", data={"name": "Razan", "pin": "2222"})
    r = razan.post("/team/reset-pin", data={"name": "Dana"},
                   follow_redirects=False)
    assert r.status_code == 303
    assert not store.has_pin("Dana")
    # Dana's sessions are dead too
    assert dana.get("/me").json()["user"] is None
