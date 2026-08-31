"""The one admin surface left: the team roster.

The profile list and editor are gone. A profile is created from a master, and
re-pointed at a revised one, in step 1 of Prepare a deck; those paths are
covered in tests/test_master_ui.py."""

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
