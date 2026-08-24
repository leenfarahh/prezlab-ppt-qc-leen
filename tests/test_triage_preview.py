"""Triage endpoint, stats page, and slide preview routes.

Every test that logs a judgment patches qc.triage.TRIAGE_LOG (and DATA_DIR)
to tmp_path so the repo's real data/triage-log.jsonl is never touched.
"""

import json

import pytest
from fastapi.testclient import TestClient

import qc.triage as triage
from qc.web import _jobs, app
from tests.conftest import job_id_of

client = TestClient(app)


def _isolate_log(monkeypatch, tmp_path):
    monkeypatch.setattr(triage, "DATA_DIR", tmp_path)
    monkeypatch.setattr(triage, "TRIAGE_LOG", tmp_path / "triage-log.jsonl")
    return tmp_path / "triage-log.jsonl"


def _audit_and_get_job(fixtures_dir, deck_name="mixed_layouts.pptx",
                       profile="prezlab_en"):
    deck = (fixtures_dir / deck_name).read_bytes()
    r = client.post("/audit",
                    files={"deck": (deck_name, deck, "application/octet-stream")},
                    data={"profile": profile})
    assert r.status_code == 200
    job_id = job_id_of(r)
    return job_id, client.get(f"/manifest/{job_id}").json()


def _powerpoint_available() -> bool:
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application").Close()
        return True
    except OSError:
        return False


def test_triage_confirm_logs_and_counts(fixtures_dir, tmp_path, monkeypatch):
    log = _isolate_log(monkeypatch, tmp_path)
    job_id, manifest = _audit_and_get_job(fixtures_dir)
    record_id = manifest["records"][0]["record_id"]

    r = client.post("/triage", json={"job_id": job_id, "record_id": record_id,
                                     "state": "confirmed"})
    assert r.status_code == 200
    assert r.json()["counts"]["confirmed"] == 1

    lines = log.read_text(encoding="utf-8").splitlines()
    assert len(lines) == 1
    entry = json.loads(lines[0])
    assert entry["record_id"] == record_id
    assert entry["state"] == "confirmed"


def test_triage_toggle_to_cleared(fixtures_dir, tmp_path, monkeypatch):
    log = _isolate_log(monkeypatch, tmp_path)
    job_id, manifest = _audit_and_get_job(fixtures_dir)
    record_id = manifest["records"][0]["record_id"]

    r = client.post("/triage", json={"job_id": job_id, "record_id": record_id,
                                     "state": "confirmed"})
    assert r.status_code == 200 and r.json()["counts"]["confirmed"] == 1

    r = client.post("/triage", json={"job_id": job_id, "record_id": record_id,
                                     "state": "cleared"})
    assert r.status_code == 200
    assert r.json()["counts"]["confirmed"] == 0

    assert len(log.read_text(encoding="utf-8").splitlines()) == 2
    # latest state cleared: the record drops out of the aggregate entirely
    assert triage.stats() == []


def test_triage_false_positive_fp_rate(fixtures_dir, tmp_path, monkeypatch):
    _isolate_log(monkeypatch, tmp_path)
    job_id, manifest = _audit_and_get_job(fixtures_dir)
    record = manifest["records"][0]

    r = client.post("/triage", json={"job_id": job_id,
                                     "record_id": record["record_id"],
                                     "state": "false_positive"})
    assert r.status_code == 200
    assert r.json()["counts"]["false_positive"] == 1

    rows = triage.stats()
    row = next(x for x in rows if x["issue_type"] == record["issue_type"])
    assert row["fp_rate"] == 1.0
    assert row["false_alarms"] == 1 and row["confirmed"] == 0


def test_triage_bad_state_400(fixtures_dir, tmp_path, monkeypatch):
    _isolate_log(monkeypatch, tmp_path)
    job_id, manifest = _audit_and_get_job(fixtures_dir)
    r = client.post("/triage", json={"job_id": job_id,
                                     "record_id": manifest["records"][0]["record_id"],
                                     "state": "maybe"})
    assert r.status_code == 400


def test_triage_unknown_job_400(tmp_path, monkeypatch):
    _isolate_log(monkeypatch, tmp_path)
    r = client.post("/triage", json={"job_id": "deadbeef", "record_id": "x",
                                     "state": "confirmed"})
    assert r.status_code == 400


def test_triage_unknown_record_404(fixtures_dir, tmp_path, monkeypatch):
    _isolate_log(monkeypatch, tmp_path)
    job_id, _ = _audit_and_get_job(fixtures_dir)
    r = client.post("/triage", json={"job_id": job_id, "record_id": "no-such-record",
                                     "state": "confirmed"})
    assert r.status_code == 404


def test_stats_page_renders_with_empty_log(tmp_path, monkeypatch):
    _isolate_log(monkeypatch, tmp_path)  # nonexistent file: stats() -> []
    r = client.get("/stats")
    assert r.status_code == 200
    assert "text/html" in r.headers["content-type"]
    assert "Triage stats" in r.text


def test_triage_persists_across_requests(fixtures_dir, tmp_path, monkeypatch):
    # The report HTML is only produced at audit/apply time, so persistence is
    # verified via the POST response counts and the JSONL file instead of a
    # re-rendered page.
    log = _isolate_log(monkeypatch, tmp_path)
    job_id, manifest = _audit_and_get_job(fixtures_dir)
    ids = [r["record_id"] for r in manifest["records"][:2]]

    for record_id in ids:
        r = client.post("/triage", json={"job_id": job_id, "record_id": record_id,
                                         "state": "confirmed"})
        assert r.status_code == 200
    assert r.json()["counts"]["confirmed"] == 2

    entries = [json.loads(line) for line in
               log.read_text(encoding="utf-8").splitlines()]
    assert [e["record_id"] for e in entries] == ids
    assert all(e["state"] == "confirmed" for e in entries)


def test_slide_preview_and_thumb_png(fixtures_dir):
    if not _powerpoint_available():
        pytest.skip("PowerPoint not installed; slide rendering is desktop-only")

    job_id, _ = _audit_and_get_job(fixtures_dir)
    r = client.get(f"/slide/{job_id}/0")
    assert r.status_code == 200
    body = r.json()
    assert body["png"] == f"/thumb/{job_id}/0.png"
    assert isinstance(body["rects"], list)

    img = client.get(body["png"])
    assert img.status_code == 200
    assert img.content[:8] == b"\x89PNG\r\n\x1a\n"

    # second hit serves from the in-memory cache
    assert client.get(f"/slide/{job_id}/0").status_code == 200


def test_slide_unknown_job_404():
    assert client.get("/slide/deadbeef/0").status_code == 404


def test_thumb_unrendered_job_404(fixtures_dir):
    job_id, _ = _audit_and_get_job(fixtures_dir, "clean.pptx")
    assert client.get(f"/thumb/{job_id}/0.png").status_code == 404


def test_apply_invalidates_cached_thumbs(fixtures_dir, tmp_path, monkeypatch):
    from qc.fixer import is_fixable

    _isolate_log(monkeypatch, tmp_path)
    job_id, manifest = _audit_and_get_job(fixtures_dir)
    # seed a fake cache so invalidation is observable without COM rendering
    _jobs[job_id]["thumbs"] = {0: b"fake-png"}
    _jobs[job_id]["rects"] = {0: []}

    fixable_ids = [r["record_id"] for r in manifest["records"] if is_fixable(r)]
    assert fixable_ids
    r = client.post("/apply", data={"job_id": job_id, "record_ids": fixable_ids})
    assert r.status_code == 200
    assert _jobs[job_id].get("thumbs") is None
    assert _jobs[job_id].get("rects") is None
