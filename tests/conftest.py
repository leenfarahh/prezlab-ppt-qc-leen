"""Shared pytest scaffolding.

Provides the synthetic fixture corpus (generated on demand) and small
factories for building planted-violation decks inside tests.
"""

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402

from qc.profile import Profile  # noqa: E402

FIXTURES = ROOT / "fixtures"


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    """Generate the spike corpus once per session if absent."""
    if not (FIXTURES / "clean.pptx").exists():
        from spike.fixtures import main as build

        build()
    return FIXTURES


@pytest.fixture()
def make_prs():
    """Factory for a 16:9 presentation matching the profile safe zones."""

    def _make():
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        return prs

    return _make


@pytest.fixture()
def en_profile() -> Profile:
    return Profile.load("prezlab_en")


@pytest.fixture()
def bilingual_profile() -> Profile:
    return Profile.load("prezlab_bilingual")


def job_id_of(response) -> str:
    """The job an upload created.

    ONE place that knows how, because the answer changed: an upload used to
    respond with the audit report and the id was read out of its markup, and it
    now redirects to Design QC (design lead, 24/08/2026 - the report is a list
    of two thousand occurrences with no picture, and a designer's first question
    is about a slide). Reading it off the final URL works for both, so a test
    that wants the report asks for /audit/<id> and says so."""
    import re

    m = re.search(r"/(?:design|audit)/([0-9a-f]{32})", str(response.url))
    if m:
        return m.group(1)
    # pre-redirect shape: the id in a link on the returned page
    m = re.search(r"/manifest/([0-9a-f]{32})", response.text)
    assert m, f"no job id in the response or its url ({response.url})"
    return m.group(1)


def save_and_ctx(prs, tmp_path: Path, profile: Profile, name: str = "case.pptx"):
    """Save a built deck and return a ready AuditContext for direct module
    detect() calls (reopens the file so tests exercise the real read path)."""
    from qc.engine import AuditContext, _build_arabic_index

    path = tmp_path / name
    prs.save(path)
    reopened = Presentation(path)
    ctx = AuditContext(prs=reopened, profile=profile, deck_path=path)
    _build_arabic_index(reopened, ctx)
    return ctx


@pytest.fixture(autouse=True)
def _isolate_local_data(monkeypatch, tmp_path):
    """No test may inherit this machine's state.

    History, users, and triage redirect to the test's tmp dir unless a test
    re-patches. Feature switches are pinned to their REPO defaults, not to
    whatever the developer's gitignored .env says: a local QC_AI=0 must not
    silently turn the assistant and copilot tests green-by-absence. Tests
    that exercise the off state patch the switch themselves."""
    import qc.store as _store
    import qc.triage as _triage
    import qc.web as _web

    monkeypatch.setattr(_store, "DATA_DIR", tmp_path)
    monkeypatch.setattr(_store, "DB_PATH", tmp_path / "qc.db")
    monkeypatch.setattr(_triage, "DATA_DIR", tmp_path)
    monkeypatch.setattr(_triage, "TRIAGE_LOG", tmp_path / "triage-log.jsonl")
    monkeypatch.setattr(_web, "AI_ENABLED", True, raising=False)
