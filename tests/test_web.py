"""Web pilot: upload -> audit -> report -> manifest, plus rejection paths."""

from fastapi.testclient import TestClient

from qc.web import app
from tests.conftest import job_id_of

client = TestClient(app)


def test_index_lists_profiles_and_modules():
    r = client.get("/")
    assert r.status_code == 200
    assert "prezlab_en" in r.text and "prezlab_bilingual" in r.text
    assert "color_palette" in r.text


def test_health():
    assert client.get("/health").json() == {"status": "ok"}


def test_an_upload_lands_on_the_slides_not_on_the_occurrence_list(fixtures_dir):
    """The report is a list of occurrences with no picture, and a designer's
    first question is "what is wrong with THIS slide" (design lead,
    24/08/2026). It is still there - Design QC links back to it, and the
    exports, triage and comments live on it - it is just not the doorway.

    A 303 rather than a rendered page, so refreshing does not re-upload the
    deck and re-run the audit."""
    deck = (fixtures_dir / "clean.pptx").read_bytes()
    r = client.post("/audit", follow_redirects=False,
                    files={"deck": ("clean.pptx", deck, "application/octet-stream")},
                    data={"profile": "prezlab_en"})
    assert r.status_code == 303
    assert r.headers["location"].startswith("/design/")

    job = r.headers["location"].rsplit("/", 1)[1]
    assert "Design QC" in client.get(f"/design/{job}").text
    back = client.get(f"/audit/{job}")
    assert back.status_code == 200 and "clean.pptx" in back.text


def test_every_page_offers_the_way_back_to_a_new_audit():
    """Landing on Design QC makes the header the only route back to the upload
    form, so it has to be in the header."""
    assert '<a href="/">Run an audit</a>' in client.get("/").text


def test_audit_bilingual_fixture_end_to_end(fixtures_dir):
    deck = (fixtures_dir / "bilingual_ar.pptx").read_bytes()
    r = client.post(
        "/audit",
        files={"deck": ("bilingual_ar.pptx", deck,
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        data={"profile": "prezlab_bilingual"},
    )
    assert r.status_code == 200
    # the upload lands on Design QC; the occurrence list is the report's job
    report = client.get(f"/audit/{job_id_of(r)}").text
    assert "bilingual_ar.pptx" in report
    assert "font.cs_typeface_missing" in report  # the Arabic guard is visible
    assert ">AR<" in report                      # Arabic badge rendered

    # manifest link present and serves the JSON
    job_id = job_id_of(r)
    m = client.get(f"/manifest/{job_id}")
    assert m.status_code == 200
    body = m.json()
    assert body["profile_id"] == "prezlab_bilingual"
    assert body["summary"]["arabic_flagged"] >= 1


def test_rejects_non_pptx():
    r = client.post("/audit", files={"deck": ("notes.txt", b"hello", "text/plain")},
                    data={"profile": "prezlab_en"})
    assert r.status_code == 400


def test_rejects_unknown_profile(fixtures_dir):
    deck = (fixtures_dir / "clean.pptx").read_bytes()
    r = client.post("/audit", files={"deck": ("clean.pptx", deck, "application/octet-stream")},
                    data={"profile": "nope"})
    assert r.status_code == 400


def test_corrupt_pptx_handled_gracefully():
    r = client.post("/audit", files={"deck": ("bad.pptx", b"not a zip", "application/octet-stream")},
                    data={"profile": "prezlab_en"})
    assert r.status_code == 422
    assert "Could not audit" in r.text


def test_unknown_manifest_404():
    assert client.get("/manifest/deadbeef").status_code == 404


def test_zip_bomb_rejected(monkeypatch):
    import io
    import zipfile

    import qc.web as web

    # Lower the ceiling so the test archive stays small: 2 MB of zeros
    # declared against a 1 MB uncompressed limit must be rejected pre-parse.
    monkeypatch.setattr(web, "MAX_UNCOMPRESSED_BYTES", 1 * 1024 * 1024)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("ppt/slides/slide1.xml", b"\x00" * (2 * 1024 * 1024))
    r = client.post("/audit",
                    files={"deck": ("bomb.pptx", buf.getvalue(), "application/octet-stream")},
                    data={"profile": "prezlab_en"})
    assert r.status_code == 413
    assert "rejected" in r.text


def test_zip_bomb_reason_passes_normal_deck(fixtures_dir):
    from qc.web import _zip_bomb_reason

    assert _zip_bomb_reason((fixtures_dir / "large_200.pptx").read_bytes()) is None
    assert _zip_bomb_reason(b"not a zip at all") is None  # 422 path handles it


def _audit_and_get_job(fixtures_dir, deck_name="mixed_layouts.pptx", profile="prezlab_en"):
    deck = (fixtures_dir / deck_name).read_bytes()
    r = client.post("/audit",
                    files={"deck": (deck_name, deck, "application/octet-stream")},
                    data={"profile": profile})
    assert r.status_code == 200
    job_id = job_id_of(r)
    return job_id, client.get(f"/manifest/{job_id}").json()


def test_apply_fixes_end_to_end(fixtures_dir, tmp_path):
    import io

    from pptx import Presentation

    from qc.fixer import is_fixable

    job_id, manifest = _audit_and_get_job(fixtures_dir)
    fixable_ids = [r["record_id"] for r in manifest["records"] if is_fixable(r)]
    assert fixable_ids

    r = client.post("/apply", data={"job_id": job_id, "record_ids": fixable_ids})
    assert r.status_code == 200
    # verify-after-write: the response is a re-audit of the cleaned deck
    assert "Applied" in r.text and "Re-audit" in r.text
    assert "Download cleaned" in r.text  # action visible in the sticky bar

    # cleaned deck downloads and opens
    d = client.get(f"/download/{job_id}")
    assert d.status_code == 200
    assert '.cleaned.pptx"' in d.headers["content-disposition"]
    prs = Presentation(io.BytesIO(d.content))
    assert len(prs.slides) == 10

    # re-audit the cleaned bytes: applied issue types are gone
    from qc.engine import run_audit

    out = tmp_path / "cleaned.pptx"
    out.write_bytes(d.content)
    again = run_audit(out, "prezlab_en")
    assert "master_slide.placeholder_geometry_off" not in again.summary["by_issue_type"]
    assert "font.family_out_of_set" not in again.summary["by_issue_type"]


def test_apply_with_no_selection_is_graceful(fixtures_dir):
    job_id, _ = _audit_and_get_job(fixtures_dir)
    r = client.post("/apply", data={"job_id": job_id})
    assert r.status_code == 200
    assert "No fixes selected" in r.text


def test_apply_unknown_job_404(fixtures_dir):
    r = client.post("/apply", data={"job_id": "nope", "record_ids": ["x"]})
    assert r.status_code == 404


def test_download_before_apply_404(fixtures_dir):
    job_id, _ = _audit_and_get_job(fixtures_dir)
    assert client.get(f"/download/{job_id}").status_code == 404


def test_report_pdf_and_csv_routes(fixtures_dir):
    job_id, manifest = _audit_and_get_job(fixtures_dir, "bilingual_ar.pptx",
                                          "prezlab_bilingual")
    pdf = client.get(f"/report/{job_id}.pdf")
    assert pdf.status_code == 200
    assert pdf.content.startswith(b"%PDF-")
    csv_r = client.get(f"/report/{job_id}.csv")
    assert csv_r.status_code == 200
    assert "issue_type" in csv_r.text
    assert csv_r.text.count("\n") >= len(manifest["records"])


def test_deck_bytes_evicted_beyond_cap(fixtures_dir):
    # the oldest job loses its deck bytes once newer audits arrive; fixes
    # then degrade gracefully to 410
    first_job, manifest = _audit_and_get_job(fixtures_dir)
    from qc.fixer import is_fixable

    ids = [r["record_id"] for r in manifest["records"] if is_fixable(r)][:1]
    for _ in range(6):  # push past MAX_DECKS_IN_MEMORY
        _audit_and_get_job(fixtures_dir, "clean.pptx")
    r = client.post("/apply", data={"job_id": first_job, "record_ids": ids})
    assert r.status_code == 410
    assert "no longer held in memory" in r.text


def _powerpoint_available() -> bool:
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application").Close()
        return True
    except OSError:
        return False


def test_diff_before_after_renders(fixtures_dir):
    import pytest as _pytest

    if not _powerpoint_available():
        _pytest.skip("PowerPoint not installed; diff rendering is desktop-only")

    from qc.fixer import is_fixable

    job_id, manifest = _audit_and_get_job(fixtures_dir)
    ids = [r["record_id"] for r in manifest["records"] if is_fixable(r)]
    r = client.post("/apply", data={"job_id": job_id, "record_ids": ids})
    assert r.status_code == 200
    assert f"/diff/{job_id}" in r.text  # Review changes offered post-apply

    d = client.get(f"/diff/{job_id}")
    assert d.status_code == 200
    assert "Before" in d.text and "After" in d.text
    assert f"/render/{job_id}/before-" in d.text

    # first advertised image serves real PNG bytes
    key = d.text.split(f"/render/{job_id}/", 1)[1].split('"', 1)[0]
    img = client.get(f"/render/{job_id}/{key}")
    assert img.status_code == 200
    assert img.content[:8] == b"\x89PNG\r\n\x1a\n"

    # the side-by-side review exports as a PDF (from the cached render)
    pdf = client.get(f"/diff/{job_id}.pdf")
    assert pdf.status_code == 200
    assert pdf.content.startswith(b"%PDF-")
    assert len(pdf.content) > 20_000  # contains the slide images
    assert "review-" in pdf.headers["content-disposition"]


def test_diff_pdf_without_apply_404(fixtures_dir):
    job_id, _ = _audit_and_get_job(fixtures_dir, "clean.pptx")
    assert client.get(f"/diff/{job_id}.pdf").status_code == 404


def test_diff_without_apply_explains(fixtures_dir):
    job_id, _ = _audit_and_get_job(fixtures_dir, "clean.pptx")
    d = client.get(f"/diff/{job_id}")
    assert d.status_code == 200
    assert "Apply fixes first" in d.text


def test_module_subset_respected(fixtures_dir):
    deck = (fixtures_dir / "mixed_layouts.pptx").read_bytes()
    r = client.post("/audit",
                    files={"deck": ("mixed_layouts.pptx", deck, "application/octet-stream")},
                    data={"profile": "prezlab_en", "modules": ["master_slide"]})
    assert r.status_code == 200
    report = client.get(f"/audit/{job_id_of(r)}").text
    assert "master_slide.layout_outlier" in report
    assert "font.family_out_of_set" not in report


def test_self_consistency_audit(fixtures_dir):
    """Sense 1: audit a deck against rules derived from itself (no profile)."""
    deck = (fixtures_dir / "mixed_layouts.pptx").read_bytes()
    r = client.post("/audit",
                    files={"deck": ("mixed_layouts.pptx", deck, "application/octet-stream")},
                    data={"profile": "__self__"})
    assert r.status_code == 200
    job_id = job_id_of(r)
    m = client.get(f"/manifest/{job_id}").json()
    # profile is the self-derived one, not a saved file
    assert m["profile_id"] == "self"
    assert m["summary"]["total"] >= 0


def test_self_option_offered_on_upload_page():
    r = client.get("/")
    assert "__self__" in r.text and "Match the deck itself" in r.text


def test_downloads_survive_arabic_filenames(fixtures_dir):
    """HTTP headers are latin-1; an Arabic deck name must not crash the
    download/export responses (real-deck finding, 12/08/2026). RFC 5987:
    ASCII fallback plus filename* carries the real name."""
    from qc.web import _attachment

    headers = _attachment("عينة من العرض.cleaned.pptx")
    disp = headers["Content-Disposition"]
    disp.encode("latin-1")  # must not raise
    assert "filename*=UTF-8''" in disp

    deck = (fixtures_dir / "mixed_layouts.pptx").read_bytes()
    r = client.post("/audit",
                    files={"deck": ("عينة من العرض.pptx", deck,
                                    "application/octet-stream")},
                    data={"profile": "prezlab_en"})
    job_id = job_id_of(r)
    r = client.get(f"/report/{job_id}.csv")
    assert r.status_code == 200
    assert "filename*=UTF-8''" in r.headers["content-disposition"]
