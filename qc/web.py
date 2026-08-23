"""Local pilot web UI for Flow B (audit only).

    python -m qc.web            # http://127.0.0.1:8000

Deliberately a LOCAL PILOT, not the production deployment: binds loopback by
default, no auth (PRD forbids network exposure without the Entra sign-in
gate), uploads are audited from a temp file and deleted immediately after
processing, manifests are held in memory only. The production web tier (job
queue, storage, SSO) is PRD Section 6 and comes after the residency ruling.
"""

import io
import json
import os
import tempfile
import threading
import uuid
import zipfile
from collections import OrderedDict
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .engine import run_audit
from .profile import PROFILES_DIR, Profile
from .records import MODULES
from .ui import render_index, render_report

from .config import DEMO_BANNER
from .config import MAX_UPLOAD_MB as _MAX_UPLOAD_MB

MAX_UPLOAD_BYTES = _MAX_UPLOAD_MB * 1024 * 1024  # configurable (QC_MAX_UPLOAD_MB)
# A .pptx is a zip; XML compresses ~100x, so a small upload can expand to GBs
# in RAM when parsed (review finding: 0.26MB deck -> ~206MB peak). Bound the
# declared uncompressed size and the compression ratio before parsing. The
# absolute uncompressed cap scales with the upload cap; the ratio check is the
# real zip-bomb defense regardless of absolute size.
MAX_UNCOMPRESSED_BYTES = MAX_UPLOAD_BYTES * 6
MAX_COMPRESSION_RATIO = 120.0
RATIO_CHECK_FLOOR = 64 * 1024 * 1024  # small decks legitimately compress well
MAX_STORED_MANIFESTS = 50
# Deck bytes are retained in memory for the newest jobs only, so "apply
# selected fixes" can run without re-upload. Older jobs keep their manifest
# but lose the deck (fix option expires). Nothing is written to disk.
MAX_DECKS_IN_MEMORY = 5

# Every route outside the public set requires a signed-in session when set
# (PRD: no network-reachable endpoint without an auth gate). Driven by env
# (cloud) or the --lan flag (LAN pilot).
from .config import AI_ENABLED
from .config import AUTH_REQUIRED as _ENV_AUTH_REQUIRED

AUTH_REQUIRED = _ENV_AUTH_REQUIRED
_PUBLIC_PATHS = ("/signin", "/signout", "/health", "/me")
_PUBLIC_PREFIXES = ("/static/",)

app = FastAPI(title="Prezlab PPT QC", docs_url=None, redoc_url=None)
app.state.auth_required = _ENV_AUTH_REQUIRED


@app.on_event("startup")
def _bootstrap_admin():
    from .config import BOOTSTRAP_ADMIN
    from .store import add_user, list_users

    if BOOTSTRAP_ADMIN and not list_users():
        add_user(BOOTSTRAP_ADMIN, "admin")


@app.middleware("http")
async def require_signin(request: Request, call_next):
    # app.state has one identity even under `python -m` (which loads this
    # module twice, as __main__ and qc.web); the module global alone would
    # miss that. The global is still honored for the test monkeypatch path.
    active = getattr(request.app.state, "auth_required", False) or AUTH_REQUIRED
    if active:
        path = request.url.path
        public = path in _PUBLIC_PATHS or any(
            path.startswith(p) for p in _PUBLIC_PREFIXES)
        if not public:
            from .auth import current_user

            if current_user(request) is None:
                if request.method == "GET":
                    from fastapi.responses import RedirectResponse

                    return RedirectResponse("/signin", status_code=303)
                return JSONResponse({"error": "sign in required"}, status_code=401)
    return await call_next(request)
app.mount("/static", StaticFiles(directory=str(Path(__file__).parent / "static")),
          name="static")
# job_id -> {"manifest": dict, "deck": bytes|None, "cleaned": bytes|None,
#            "filename": str, "profile": str}
_jobs: OrderedDict[str, dict] = OrderedDict()
_jobs_lock = threading.Lock()


from .web_admin import register as _register_admin

_register_admin(app)


def _profiles_meta() -> list[dict]:
    """[{id, name}] for the profile picker, read from the profile JSONs."""
    out = []
    for p in sorted(PROFILES_DIR.glob("*.json")):
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            out.append({"id": p.stem, "name": data.get("name", p.stem)})
        except Exception:
            out.append({"id": p.stem, "name": p.stem})
    return out


@app.middleware("http")
async def reject_oversized_bodies(request: Request, call_next):
    """Reject oversized uploads from the declared Content-Length BEFORE
    Starlette spools the whole body to disk (review finding: the in-handler
    cap only fires after the full body has already been received)."""
    if request.method == "POST":
        declared = request.headers.get("content-length")
        if declared and declared.isdigit() and int(declared) > MAX_UPLOAD_BYTES + 1_048_576:
            mb = int(declared) // 1_048_576
            msg = (f"That upload is about {mb} MB, over this instance's "
                   f"{_MAX_UPLOAD_MB} MB cap.")
            if DEMO_BANNER:
                msg += (" Large client decks belong on the office LAN "
                        "instance, which accepts up to 250 MB and keeps "
                        "files on-site.")
            return HTMLResponse(
                render_index(_pickable_profiles(), MODULES, msg),
                status_code=413)
    return await call_next(request)


def _zip_bomb_reason(data: bytes) -> str | None:
    """Non-None when the archive declares an implausible expansion. A broken
    zip returns None here; python-pptx raises and the 422 path handles it."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            total = sum(i.file_size for i in z.infolist())
    except zipfile.BadZipFile:
        return None
    if total > MAX_UNCOMPRESSED_BYTES:
        return f"declares {total / 1e6:.0f} MB uncompressed (limit {MAX_UNCOMPRESSED_BYTES / 1e6:.0f} MB)"
    if total > RATIO_CHECK_FLOOR and total / max(len(data), 1) > MAX_COMPRESSION_RATIO:
        return "implausible compression ratio for a presentation"
    return None


def _profiles() -> list[str]:
    return sorted(p.stem for p in PROFILES_DIR.glob("*.json"))


# Audit a deck against rules derived from the deck itself (Sense 1): no brand
# needed, flags slides that break from the deck's own dominant conventions.
SELF_PROFILE = "__self__"
_SELF_META = {"id": SELF_PROFILE, "label": "Match the deck itself",
              "name": "No brand needed; flags slides that break from the "
                      "deck's own conventions"}

# Audit a deck against a master submitted in the same request: Stage 1 reads
# the master's design surface, and the resulting Style Spec becomes the rules
# for this one job. Nothing is written to disk, so a designer with no admin
# rights can do it; saving it as a reusable profile is a separate, gated step
# on the /master page.
MASTER_PROFILE = "__master__"
_MASTER_META = {"id": MASTER_PROFILE, "label": "Match a master I'll upload",
                "name": "Reads the master's theme, layouts, and grid, then "
                        "checks the deck against them"}

# Neither sentinel corresponds to a profile file, so anything that tunes or
# reloads a SAVED profile has to refuse them.
_EPHEMERAL = (SELF_PROFILE, MASTER_PROFILE)


def _ephemeral(profile_key: str) -> bool:
    return profile_key in _EPHEMERAL


# Anthropic-backed features are refused at the route, not just hidden in the
# UI: hiding a button is a presentation choice, and QC_AI=0 has to mean no
# request leaves this machine even for someone POSTing the endpoint directly.
_AI_OFF = {"error": "AI features are disabled on this instance (QC_AI=0). "
                    "No request is sent to the Anthropic API."}


def _ai_disabled_response():
    return JSONResponse(_AI_OFF, status_code=503) if not AI_ENABLED else None


def _pickable_profiles() -> list[dict]:
    return [_SELF_META, _MASTER_META] + _profiles_meta()


def _attachment(filename: str) -> dict:
    """Content-Disposition that survives non-ASCII names (RFC 5987): an
    ASCII fallback plus filename* in UTF-8. HTTP headers are latin-1, so a
    bare Arabic filename crashes the response (real-deck finding,
    12/08/2026: every download/export died on an Arabic deck name)."""
    from urllib.parse import quote

    fallback = "".join(c if 31 < ord(c) < 127 and c not in '"\\' else "_"
                       for c in filename).strip() or "download"
    return {"Content-Disposition":
            f'attachment; filename="{fallback}"; '
            f"filename*=UTF-8''{quote(filename, safe='')}"}


def _resolve_profile(profile: str, data: bytes) -> Profile:
    """A Profile object for the chosen id, or one bootstrapped from the deck
    itself when the self-consistency option is picked."""
    if profile == SELF_PROFILE:
        from pptx import Presentation

        from .bootstrap import build_profile

        prs = Presentation(io.BytesIO(data))
        return Profile(build_profile(prs, "self", "This deck (self-consistency)"))
    return Profile.load(profile)


@app.get("/", response_class=HTMLResponse)
def index():
    return render_index(_pickable_profiles(), MODULES)


@app.get("/health")
def health():
    return {"status": "ok"}


# ------------------------------------------------------- Stage 1: the master
# Reading a master is a different job from auditing a deck: the input is the
# design surface, the output is a Style Spec, and the spec (not the file) is
# what everything downstream consumes. Specs are held in memory beside jobs,
# under the same rule as decks: nothing is written to disk.
MAX_SPECS_IN_MEMORY = 20
_specs: OrderedDict[str, dict] = OrderedDict()
_specs_lock = threading.Lock()


def _remember_spec(spec: dict, master_bytes: bytes | None = None) -> str:
    """Hold the spec and the master it came from. The master is kept because a
    spec DESCRIBES a design system and cannot be applied as one: restyling a
    slide needs real slideLayout parts for PowerPoint to match placeholders
    against. Saving the spec as a profile hands these bytes to the template
    store; until then they live in memory like everything else here."""
    spec_id = uuid.uuid4().hex
    with _specs_lock:
        _specs[spec_id] = {"spec": spec, "master": master_bytes}
        while len(_specs) > MAX_SPECS_IN_MEMORY:
            _specs.popitem(last=False)
    return spec_id


def _get_spec(spec_id: str) -> dict | None:
    with _specs_lock:
        held = _specs.get(spec_id)
    return held["spec"] if held else None


def _get_spec_master(spec_id: str) -> bytes | None:
    with _specs_lock:
        held = _specs.get(spec_id)
    return held["master"] if held else None


@app.get("/master", response_class=HTMLResponse)
def master_intake():
    from .ui_master import render_master_intake

    return HTMLResponse(render_master_intake())


@app.post("/master", response_class=HTMLResponse)
def master_read(request: Request, master: UploadFile = File(...)):
    from pptx import Presentation

    from .stylespec import extract_style_spec
    from .ui_master import render_master_intake, render_style_spec
    from .web_admin import _editor

    filename = master.filename or "master.pptx"
    if not filename.lower().endswith(".pptx"):
        return HTMLResponse(render_master_intake("Only .pptx files are accepted."),
                            status_code=400)

    data = master.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        return HTMLResponse(
            render_master_intake(f"File exceeds the {_MAX_UPLOAD_MB} MB cap."),
            status_code=413)
    bomb = _zip_bomb_reason(data)
    if bomb:
        return HTMLResponse(render_master_intake(f"File rejected: {bomb}."),
                            status_code=413)

    try:
        spec = extract_style_spec(Presentation(io.BytesIO(data)), source=filename)
    except Exception as exc:
        return HTMLResponse(
            render_master_intake(f"Could not read that master: "
                                 f"{type(exc).__name__}: {exc}"), status_code=422)

    spec_id = _remember_spec(spec, data)
    return HTMLResponse(render_style_spec(
        spec, spec_id, can_save=_editor(request) is not None))


@app.get("/spec/{spec_id}.json")
def spec_json(spec_id: str):
    """The Style Spec itself: the canonical artifact, downloadable so it can be
    archived and replayed against a future deck without the master."""
    spec = _get_spec(spec_id)
    if spec is None:
        return JSONResponse({"error": "spec expired or unknown"}, status_code=404)
    source = (spec.get("meta") or {}).get("source_file") or "master"
    stem = Path(source).stem or "master"
    return JSONResponse(spec, headers={
        "Content-Disposition": f'attachment; filename="{stem}-stylespec.json"'})


@app.post("/spec/{spec_id}/profile", response_class=HTMLResponse)
def spec_to_profile_route(request: Request, spec_id: str, name: str = Form(...)):
    from datetime import date

    from fastapi.responses import RedirectResponse

    from .profile import PROFILES_DIR
    from .stylespec import spec_to_profile
    from .ui_admin import render_admin_error
    from .ui_master import render_style_spec
    from .web_admin import _editor, _slugify, _unique_pid

    spec = _get_spec(spec_id)
    if spec is None:
        return HTMLResponse(render_admin_error(
            "Style Spec", "That spec has expired",
            "Specs are held in memory only. Read the master again.",
            back="/master"), status_code=404)

    editor = _editor(request)
    if editor is None:
        return HTMLResponse(render_admin_error(
            "Style Spec", "Saving a profile needs a lead or admin",
            "Sign in as a lead or admin to save this spec as a profile.",
            back="/master"), status_code=403)

    slug = _slugify(name)
    if not slug:
        return HTMLResponse(render_style_spec(
            spec, spec_id, can_save=True,
            message="Give the profile a name with letters or numbers."),
            status_code=400)

    pid = _unique_pid(slug)
    profile = spec_to_profile(spec, pid, name.strip())
    profile["owner"] = (f"{editor['name']} (from master "
                        f"{(spec.get('meta') or {}).get('source_file')}, "
                        f"{date.today().strftime('%d/%m/%Y')})")
    (PROFILES_DIR / f"{pid}.json").write_text(
        json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")

    # The rules alone cannot restyle a slide; keep the master file so this
    # profile can be applied, not just audited against.
    master_bytes = _get_spec_master(spec_id)
    if master_bytes:
        from .templates import save_master

        save_master(pid, master_bytes)
    return RedirectResponse(f"/profiles/{pid}/edit", status_code=303)


def _profile_from_master(master: UploadFile | None):
    """(Profile, spec, error_message) for the upload-a-master option.

    The master is read for its design surface, projected onto a profile, and
    then dropped: only the derived rules survive into the job. That is the
    Style Spec contract holding at the web tier, not just in the library."""
    from pptx import Presentation

    from .stylespec import extract_style_spec, spec_to_profile

    if master is None or not master.filename:
        return None, None, ("Pick the master .pptx to check this deck against, "
                            "or choose a different rule source.")
    if not master.filename.lower().endswith(".pptx"):
        return None, None, "The master must be a .pptx file."

    raw = master.file.read(MAX_UPLOAD_BYTES + 1)
    if len(raw) > MAX_UPLOAD_BYTES:
        return None, None, f"The master exceeds the {_MAX_UPLOAD_MB} MB cap."
    bomb = _zip_bomb_reason(raw)
    if bomb:
        return None, None, f"The master was rejected: {bomb}."

    try:
        spec = extract_style_spec(Presentation(io.BytesIO(raw)),
                                  source=master.filename)
    except Exception as exc:
        return None, None, (f"Could not read that master: "
                            f"{type(exc).__name__}: {exc}")

    stem = Path(master.filename).stem or "master"
    # A readable id, because it lands in the report header and in history.
    return (Profile(spec_to_profile(spec, f"master:{stem}",
                                    f"Master: {master.filename}")),
            spec, None)


# ----------------------------------------------- Stage 3: apply the master
# Auditing reads a deck; this REWRITES it, so it gets its own page rather than
# hiding behind the audit form. Results are held in memory like every other
# job: nothing is written to disk except the profile's stored master.
MAX_FORMAT_JOBS = 10
_format_jobs: OrderedDict[str, dict] = OrderedDict()
_format_lock = threading.Lock()


def _formattable_profiles() -> list[dict]:
    """Every saved profile, flagged with whether it carries a master. Profiles
    without one are still listed by the page so the absence is explained
    rather than looking like a missing option.

    Each one also reports WHICH master it would apply: when the stored file was
    stored, and what frame that file states. A profile whose stored copy
    predates the designer's latest master formats decks on the old one, and the
    only symptom is something missing from the output - a presentation-space
    rectangle, in the case that made this necessary (design lead, 21/08/2026).
    Read at the point of use, because that is where the surprise happens."""
    from datetime import datetime

    from .stylespec import dominant_master, extract_layouts, infer_grid
    from .templates import has_master, load_master, master_info

    out = []
    for meta in _profiles_meta():
        pid = meta["id"]
        n_layouts = 0
        frame = stored = None
        if has_master(pid):
            try:
                from pptx import Presentation

                prs = Presentation(io.BytesIO(load_master(pid)))
                master = dominant_master(prs)
                n_layouts = len(extract_layouts(master, embed_assets=False)) \
                    if master is not None else 0
                if master is not None:
                    frame = infer_grid(prs, master).get("source")
                info = master_info(pid)
                if info:
                    stored = datetime.fromtimestamp(
                        info["modified"]).strftime("%d/%m/%Y")
            except Exception:
                n_layouts = 0
        out.append({"id": pid, "name": meta["name"],
                    "has_master": has_master(pid), "layouts": n_layouts,
                    "frame": frame, "master_stored": stored})
    return out


@app.get("/format", response_class=HTMLResponse)
def format_intake():
    from .ui_format import render_format_intake
    from .unify import com_available

    return HTMLResponse(render_format_intake(_formattable_profiles(),
                                             com_ready=com_available()))


@app.post("/format", response_class=HTMLResponse)
def format_deck(request: Request, deck: UploadFile = File(...),
                profile: str = Form(...)):
    from pptx import Presentation

    from .applymaster import apply_master, plan_assignments
    from .stylespec import dominant_master, extract_layouts, infer_grid
    from .templates import load_master
    from .ui_format import render_format_intake, render_format_result
    from .unify import com_available

    def _fail(msg, code=400):
        return HTMLResponse(
            render_format_intake(_formattable_profiles(), msg,
                                 com_ready=com_available()), status_code=code)

    filename = deck.filename or "deck.pptx"
    if not filename.lower().endswith(".pptx"):
        return _fail("Only .pptx files are accepted.")
    if profile not in _profiles():
        return _fail("Unknown profile.")
    master_bytes = load_master(profile)
    if not master_bytes:
        return _fail("That profile carries no master file, so there is nothing "
                     "to apply. Create it from a master on the Read a master page.")

    data = deck.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        return _fail(f"File exceeds the {_MAX_UPLOAD_MB} MB cap.", 413)
    bomb = _zip_bomb_reason(data)
    if bomb:
        return _fail(f"File rejected: {bomb}.", 413)

    try:
        deck_prs = Presentation(io.BytesIO(data))
        master_prs = Presentation(io.BytesIO(master_bytes))
        target = dominant_master(master_prs)
        if target is None:
            return _fail("That profile's master file has no slide master.")
        # Layouts are read from the stored master rather than from an archived
        # spec: the file is the source of truth and cannot drift from itself.
        plans = plan_assignments(deck_prs,
                                 extract_layouts(target, embed_assets=False))
        # The submitted master's own frame, kept only as a fallback: the output
        # deck is asked first, because PowerPoint resizes a loaded design to the
        # deck's slide size and the file it writes is the only place the frame's
        # real numbers are (qc.pspace).
        master_space = (infer_grid(master_prs, target)
                        or {}).get("presentation_space")
        master_size = (master_prs.slide_width, master_prs.slide_height)
    except Exception as exc:
        return _fail(f"Could not read that deck: {type(exc).__name__}: {exc}", 422)

    if not plans:
        return _fail("That deck has no slides to format.")

    result = apply_master(data, master_bytes, plans)
    if result.fatal:
        return _fail(result.fatal, 503)

    # The presentation space goes in BEFORE the content moves. The migration
    # seats every slide's body on the frame it reads from that slide's own
    # master, so a deck whose masters carry the marker seats its stragglers on
    # the new frame too, instead of on whatever the original design implied.
    from .pspace import ensure_presentation_space

    try:
        result.deck, space_notes = ensure_presentation_space(
            result.deck,
            fallback_box=(master_space or {}).get("box_emu")
            if not (master_space or {}).get("problem") else None,
            fallback_size=master_size)
    except Exception as exc:
        space_notes = [f"The presentation space could not be written into the "
                       f"deck ({type(exc).__name__}: {exc}). The layouts were "
                       f"still applied; the marker is missing from the masters."]

    # Applying the layout is only half of it: PowerPoint remaps placeholder
    # content, but a deck of free-floating shapes keeps them exactly where
    # they were, leaving the master's empty placeholders on top of the old
    # content. The migration pass moves the content into the master.
    from .migrate import migrate_deck

    try:
        deck_out, content_changes = migrate_deck(result.deck)
    except Exception as exc:
        deck_out, content_changes = result.deck, []
        content_changes = [type("C", (), {
            "slide_index": 0, "action": "migration skipped",
            "detail": f"{type(exc).__name__}: {exc}; layouts applied, content "
                      f"left where it was"})()]
    result.deck = deck_out

    job_id = uuid.uuid4().hex
    with _format_lock:
        # The job keeps the changes as well as the deck: putting a removed
        # piece back re-renders this same page, and re-deriving the change list
        # would mean running the whole migration again.
        # `source` is the deck exactly as it was uploaded, kept for the review
        # page: "before" is not derivable from the output, and re-uploading the
        # file to see what changed is the thing the review exists to avoid.
        _format_jobs[job_id] = {"deck": result.deck, "source": data,
                                "filename": filename,
                                "profile": profile, "plans": result.plans,
                                "errors": result.errors,
                                "applied": result.applied,
                                "masters": result.masters,
                                "stragglers": result.stragglers,
                                "space_notes": space_notes,
                                "changes": content_changes, "restored": [],
                                "undone": [], "undo_notes": {}}
        while len(_format_jobs) > MAX_FORMAT_JOBS:
            _format_jobs.popitem(last=False)

    profile_meta = next((p for p in _profiles_meta() if p["id"] == profile), None)
    return HTMLResponse(render_format_result(
        deck_name=filename, profile_name=(profile_meta or {}).get("name", profile),
        job_id=job_id, plans=result.plans, errors=result.errors,
        applied=result.applied, content_changes=content_changes,
        masters=result.masters, stragglers=result.stragglers,
        space_notes=space_notes))


@app.post("/format/{job_id}/restore", response_class=HTMLResponse)
def format_restore(job_id: str, restore_ids: list[str] = Form(None)):
    """Put selected removed pieces back into the rebuilt deck.

    The migration removes header text the master has no placeholder for, and
    says so loudly. This is the other half of that promise: the designer, not
    the tool, decides whether the removal was right, and the piece goes back
    exactly as it was rather than being retyped from the report."""
    from .migrate import restore_shapes
    from .ui_format import render_format_result

    with _format_lock:
        job = _format_jobs.get(job_id)
    if job is None or job.get("deck") is None:
        return JSONResponse({"error": "unknown or expired job"}, status_code=404)

    # Anything already back is skipped rather than inserted twice: a browser
    # resubmitting this POST (refresh, back button) would otherwise put a second
    # copy of the same shape on the slide.
    wanted = set(restore_ids or []) - set(job.get("restored") or [])
    items = [{"slide_index": c.slide_index,
              "removed_xml": getattr(c, "removed_xml", None),
              "restore_id": getattr(c, "restore_id", None)}
             for c in job.get("changes") or []
             if getattr(c, "restore_id", None) in wanted]
    if items:
        try:
            deck, outcomes = restore_shapes(job["deck"], items)
        except Exception as exc:
            job["restore_error"] = f"{type(exc).__name__}: {exc}"
        else:
            with _format_lock:
                job["deck"] = deck
                # Kept per piece, not as a bare list of ids: a piece nudged clear
                # of the master's header is back in a different place than it
                # left, and a designer has to be told which.
                notes = dict(job.get("restored_notes") or {})
                for o in outcomes:
                    if o.get("restore_id"):
                        notes[o["restore_id"]] = o["detail"]
                job["restored_notes"] = notes
                job["restored"] = sorted(notes)

    profile = job.get("profile")
    profile_meta = next((p for p in _profiles_meta() if p["id"] == profile), None)
    return HTMLResponse(render_format_result(
        deck_name=job["filename"],
        profile_name=(profile_meta or {}).get("name", profile),
        job_id=job_id, plans=job.get("plans") or [],
        errors=job.get("errors") or {}, applied=job.get("applied") or 0,
        content_changes=job.get("changes") or [],
        restored=job.get("restored") or [],
        restored_notes=job.get("restored_notes") or {},
        restore_error=job.get("restore_error"),
        masters=job.get("masters") or 1,
        stragglers=job.get("stragglers") or [],
        space_notes=job.get("space_notes") or []))


# How many slides the deck view renders on ONE PAGE. Every render is a
# PowerPoint export, so an uncapped page on a 200-slide deck is a several-minute
# wait. This is a page size, not a limit on what can be reviewed: the page pages
# through the rest, because a review that silently ends at slide 20 of a 26-slide
# deck reads as six slides the tool declined to show (design lead, 23/08/2026).
REVIEW_PAGE_SIZE = 20


def _reviewable(job) -> list[int]:
    """Every slide worth reviewing: the ones this run changed, plus any that
    failed. A slide the content pass did not touch has nothing to review and
    would only push the ones that do further down the page."""
    changed = {c.slide_index for c in (job.get("changes") or [])}
    changed |= set((job.get("errors") or {}).keys())
    return sorted(changed)


def _review_slides(job, page: int = 0) -> list[int]:
    """One page of reviewable slides."""
    every = _reviewable(job)
    start = max(0, page) * REVIEW_PAGE_SIZE
    return every[start:start + REVIEW_PAGE_SIZE]


def _layout_use(plans, key: str) -> dict:
    """{layout name: slides on it} from the assignment plans, for the badges on
    the master view. A layout nothing sits on is worth seeing as unused."""
    counts: dict[str, int] = {}
    for p in plans or []:
        name = getattr(p, key, None)
        if name:
            counts[name] = counts.get(name, 0) + 1
    return counts


def _url_keys(images: dict, prefix: str) -> dict:
    """Re-key renders as the page asks for them: "before:7" -> "slide_before_7".

    One namespace, URL-safe, and prefixed per view. Without the prefix a
    12-layout master and a 12-slide deck collide on "before:3" and the master
    view starts serving slide pictures - which is a wrong answer that looks like
    a right one, the worst kind for a review page."""
    return {f"{prefix}_{k.replace(':', '_', 1)}": png
            for k, png in (images or {}).items()}


def _ensure_review(job, view: str, page: int = 0) -> str | None:
    """Render what this view needs, once per page, and cache it on the job.
    Returns a readable reason when nothing could be rendered - the page then
    shows the change list and the Undo buttons without pictures, which is the
    part a designer cannot do without."""
    from .render import layout_previews, slide_previews

    key = "review_master" if view == "master" else f"review_deck_{page}"
    if job.get(key) is not None:
        return job.get(key + "_error")
    if job.get("source") is None:
        job[key] = {}
        job[key + "_error"] = ("the uploaded deck is no longer in memory, so "
                              "there is nothing to compare against")
        return job[key + "_error"]
    try:
        if view == "master":
            out = layout_previews(job["source"], job["deck"])
            out["images"] = _url_keys(out.get("images"), "layout")
        else:
            out = slide_previews(job["source"], job["deck"],
                                 _review_slides(job, page))
            out["images"] = _url_keys(out.get("images"), "slide")
    except Exception as exc:  # reading the DECK failed: there is no answer at all
        job[key] = {}
        job[key + "_error"] = f"{type(exc).__name__}: {exc}."
        return job[key + "_error"]
    # A render failure is kept WITH the result, not instead of it. Which layouts
    # the deck arrived with and which it has now needs no PowerPoint; discarding
    # the whole answer made the page say "No layouts to show" about a master
    # carrying twelve (design lead, 23/08/2026).
    job[key] = out
    job[key + "_error"] = out.get("error")
    return out.get("error")


def _drop_review_renders(job) -> None:
    """Forget every cached render. Called after an undo: the deck has changed,
    and a stale "after" picture beside a row marked undone is the one thing this
    page must never show."""
    for key in [k for k in list(job) if k.startswith(("review_master",
                                                      "review_deck"))]:
        job[key] = None


def _review_page(job_id: str, job, view: str, undo_error: str | None = None,
                 page: int = 0):
    from .ui_review import render_review

    error = _ensure_review(job, view, page)
    key = "review_master" if view == "master" else f"review_deck_{page}"
    previews = job.get(key) or {}
    every = _reviewable(job)
    profile = job.get("profile")
    meta = next((p for p in _profiles_meta() if p["id"] == profile), None)
    return HTMLResponse(render_review(
        deck_name=job["filename"],
        profile_name=(meta or {}).get("name", profile or ""),
        job_id=job_id, view=view, previews=previews,
        changes=job.get("changes") or [], plans=job.get("plans") or [],
        errors=job.get("errors") or {},
        undone=set(job.get("undone") or []),
        notes=job.get("undo_notes") or {},
        shown=_review_slides(job, page),
        reviewable=len(every),
        page=page, page_size=REVIEW_PAGE_SIZE,
        total_slides=len(job.get("plans") or []),
        masters=job.get("masters") or 1,
        used_before=_layout_use(job.get("plans"), "source_layout"),
        used_after=_layout_use(job.get("plans"), "target_layout"),
        truncated=bool(previews.get("truncated")),
        render_error=error, undo_error=undo_error))


@app.get("/format/{job_id}/review", response_class=HTMLResponse)
def format_review(job_id: str, view: str = "master", page: int = 0):
    """Before and after for one format run, as the template and as the deck."""
    with _format_lock:
        job = _format_jobs.get(job_id)
    if job is None or job.get("deck") is None:
        return HTMLResponse(
            "<p>Unknown or expired job. Format the deck again.</p>",
            status_code=404)
    return _review_page(job_id, job, "deck" if view == "deck" else "master",
                        page=max(0, page))


@app.post("/format/{job_id}/undo", response_class=HTMLResponse)
def format_undo(job_id: str, change_ids: list[str] = Form(None),
                page: int = Form(0)):
    """Take one reported change back, exactly as it was.

    Anything already undone is skipped rather than replayed: a browser
    resubmitting this POST would otherwise insert a second copy of a restored
    shape, or reset a shape a later undo has since moved."""
    from .undo import apply_undo, expand

    with _format_lock:
        job = _format_jobs.get(job_id)
    if job is None or job.get("deck") is None:
        return HTMLResponse(
            "<p>Unknown or expired job. Format the deck again.</p>",
            status_code=404)

    # Expanded before it is filtered: undoing a change means undoing what came
    # after it on that slide too, or the slide lands in a state the run never
    # produced (qc.undo.followers).
    changes = job.get("changes") or []
    wanted = set(change_ids or []) - set(job.get("undone") or [])
    already = set(job.get("undone") or [])
    items = [{"change_id": c.change_id, "slide_index": c.slide_index,
              "action": c.action, "ops": c.undo}
             for c in expand(changes, wanted)
             if c.change_id not in already]
    error = None
    if items:
        try:
            deck, outcomes = apply_undo(job["deck"], items)
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
        else:
            with _format_lock:
                job["deck"] = deck
                notes = dict(job.get("undo_notes") or {})
                for o in outcomes:
                    if o.get("change_id"):
                        notes[o["change_id"]] = o["detail"]
                job["undo_notes"] = notes
                job["undone"] = sorted(
                    set(job.get("undone") or [])
                    | {o["change_id"] for o in outcomes
                       if o.get("done") and o.get("change_id")})
                _drop_review_renders(job)
    return _review_page(job_id, job, "deck", undo_error=error,
                        page=max(0, page))


@app.get("/review-img/{job_id}/{key}.png")
def review_img(job_id: str, key: str):
    from fastapi.responses import Response

    with _format_lock:
        job = _format_jobs.get(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job"}, status_code=404)
    # Every review bucket is searched by PREFIX, not by an exact name: the deck
    # view caches per page ("review_deck_0", "review_deck_1"), and naming the two
    # buckets literally here meant page 0's images were cached under a key this
    # route never looked in, so the deck view rendered 18 pictures and served
    # none of them. The image keys themselves are already unique across views
    # (_url_keys), so a prefix scan cannot return the wrong picture.
    png = None
    for bucket, value in job.items():
        if not str(bucket).startswith("review_") or not isinstance(value, dict):
            continue
        png = (value.get("images") or {}).get(key)
        if png is not None:
            break
    if png is None:
        return JSONResponse({"error": "no such render"}, status_code=404)
    return Response(content=png, media_type="image/png",
                    headers={"Cache-Control": "private, max-age=600"})


@app.get("/format/{job_id}/download")
def format_download(job_id: str):
    from fastapi.responses import Response

    with _format_lock:
        job = _format_jobs.get(job_id)
    if job is None or job["deck"] is None:
        return JSONResponse({"error": "unknown or expired job"}, status_code=404)
    stem = Path(job["filename"]).stem or "deck"
    return Response(
        job["deck"],
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=_attachment(f"{stem} (master applied).pptx"))


@app.post("/audit", response_class=HTMLResponse)
def audit(request: Request, deck: UploadFile = File(...),
          profile: str = Form("prezlab_en"),
          master: UploadFile = File(None),
          modules: list[str] = Form(None)):
    filename = deck.filename or "upload.pptx"
    if not filename.lower().endswith(".pptx"):
        return HTMLResponse(render_index(_pickable_profiles(), MODULES, "Only .pptx files are accepted."), status_code=400)
    if not _ephemeral(profile) and profile not in _profiles():
        return HTMLResponse(render_index(_pickable_profiles(), MODULES, "Unknown profile."), status_code=400)
    selected = [m for m in (modules or []) if m in MODULES] or None

    master_profile, master_spec = None, None
    if profile == MASTER_PROFILE:
        master_profile, master_spec, err = _profile_from_master(master)
        if err:
            return HTMLResponse(
                render_index(_pickable_profiles(), MODULES, err), status_code=400)

    data = deck.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES, f"File exceeds the {_MAX_UPLOAD_MB} MB cap."), status_code=413)
    bomb = _zip_bomb_reason(data)
    if bomb:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES, f"File rejected: {bomb}."), status_code=413)

    fd, tmp_name = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)  # Windows cannot delete a file while this handle is open
    tmp = Path(tmp_name)
    try:
        tmp.write_bytes(data)
        try:
            profile_obj = master_profile or _resolve_profile(profile, data)
            result = run_audit(tmp, profile_obj, selected)
        except Exception as exc:
            return HTMLResponse(
                render_index(_pickable_profiles(), MODULES, f"Could not audit this file: {type(exc).__name__}: {exc}"),
                status_code=422)
    finally:
        tmp.unlink(missing_ok=True)  # PRD: uploads auto-delete after processing

    manifest = result.to_manifest()
    manifest["deck"] = filename  # temp path is meaningless to the user
    job_id = uuid.uuid4().hex
    with _jobs_lock:
        _jobs[job_id] = {"manifest": manifest, "deck": data,
                         "cleaned": None, "filename": filename,
                         "profile": profile, "profile_obj": profile_obj,
                         # kept so the report can offer the spec it used; the
                         # master's own bytes are already gone
                         "master_spec": master_spec}
        while len(_jobs) > MAX_STORED_MANIFESTS:
            _jobs.popitem(last=False)
        # drop deck bytes beyond the newest N jobs (manifests stay)
        with_deck = [k for k, v in _jobs.items() if v["deck"] is not None]
        for k in with_deck[:-MAX_DECKS_IN_MEMORY]:
            _jobs[k]["deck"] = None
            _jobs[k].pop("prev_deck", None)
            _jobs[k].pop("diff", None)
            _jobs[k].pop("thumbs", None)
            _jobs[k].pop("rects", None)
    from .auth import current_user
    from .promotion import promoted_issue_types
    from .store import comment_counts, record_audit

    user = current_user(request)
    record_audit(manifest, user["name"] if user else "anonymous")
    return render_report(manifest, job_id, can_fix=True,
                         promoted=promoted_issue_types(),
                         comments=comment_counts(filename),
                         design=design_count(_jobs[job_id]),
                         assist=AI_ENABLED and not _ephemeral(profile))


def _job(job_id: str) -> dict | None:
    return _jobs.get(job_id)


@app.post("/apply", response_class=HTMLResponse)
def apply(request: Request, job_id: str = Form(...),
          record_ids: list[str] = Form(None)):
    from .fixer import apply_fixes

    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."), status_code=404)
    if job["deck"] is None:
        return HTMLResponse(render_report(
            job["manifest"], job_id, can_fix=False,
            banner="The deck is no longer held in memory (newer audits replaced it). "
                   "Re-upload to apply fixes."), status_code=410)
    selected = set(record_ids or [])
    if not selected:
        return render_report(job["manifest"], job_id, can_fix=True,
                             banner="No fixes selected.",
                             assist=AI_ENABLED and not _ephemeral(job["profile"]))

    fx = apply_fixes(job["deck"], job["manifest"]["records"], selected)
    skipped = [o for o in fx.outcomes if o.outcome == "skipped"]
    before_total = job["manifest"]["summary"]["total"]

    # Verify-after-write: re-audit the cleaned bytes so the report reflects
    # the actual new state of the deck, not an assumption.
    fd, tmp_name = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        tmp.write_bytes(fx.cleaned_bytes)
        prof = job.get("profile_obj") or Profile.load(job["profile"])
        result = run_audit(tmp, prof)
    finally:
        tmp.unlink(missing_ok=True)

    manifest = result.to_manifest()
    manifest["deck"] = job["filename"]
    changed_ids = {o.record_id for o in fx.outcomes if o.outcome == "changed"}
    applied = [r for r in job["manifest"]["records"] if r["record_id"] in changed_ids]
    with _jobs_lock:
        job["prev_deck"] = job["deck"]   # pre-fix bytes for the visual diff
        job["applied_records"] = job.get("applied_records", []) + applied
        job["diff"] = None               # invalidate any cached render
        job["thumbs"] = None             # slides changed: thumbnails are stale
        job["rects"] = None
        job["design"] = None             # and so are the design findings: a fix
                                         # that moved a shape may have created
                                         # or cleared an overlap
        job["manifest"] = manifest
        job["deck"] = fx.cleaned_bytes   # further fixes build on the cleaned deck
        job["cleaned"] = fx.cleaned_bytes
    after_total = manifest["summary"]["total"]
    note = (f"Applied {fx.applied} fix{'es' if fx.applied != 1 else ''}. "
            f"Re-audit of the cleaned deck: {after_total} findings remain "
            f"(was {before_total}).")
    if skipped:
        note += f" Skipped {len(skipped)}."
    from .auth import current_user
    from .promotion import promoted_issue_types
    from .store import comment_counts, record_audit

    user = current_user(request)
    record_audit(manifest, user["name"] if user else "anonymous", kind="fix")
    return render_report(manifest, job_id, can_fix=True, banner=note,
                         has_cleaned=True, diff_href=f"/diff/{job_id}",
                         promoted=promoted_issue_types(),
                         comments=comment_counts(job["filename"]),
                         design=design_count(job),
                         assist=AI_ENABLED and not _ephemeral(job["profile"]))


# ------------------------------------------------------------- design QC
#
# The audit's own modules answer "does this deck match the profile". These are
# the questions that only have answers once a master is on the deck - a palette
# meeting the deck's own colours, text meeting a new background, the master's
# furniture meeting content that was already there - and each of them has more
# than one right fix, so the designer picks and the pick is reversible. Detection
# is qc.design, applying a pick is qc.remedy, taking it back is qc.undo (the same
# machinery the format review page uses, reaching one step further).


def _palette_cfg(job) -> dict:
    profile = job.get("profile_obj")
    try:
        return profile.module_config("color_palette") if profile else {}
    except Exception:
        return {}


def _ensure_design(job) -> str | None:
    """Run the design pass once per state of the deck, and cache it on the job.

    Cleared (design=None) after every apply and every undo, because the findings
    ARE a description of the current bytes: a card offering to fix something the
    designer just fixed is the one thing this page must not show."""
    if job.get("design") is not None:
        return job.get("design_error")
    if job.get("deck") is None:
        job["design"] = []
        return None
    from .design import scan

    try:
        job["design"] = scan(job["deck"], _palette_cfg(job))
        job["design_error"] = None
    except Exception as exc:
        job["design"] = []
        job["design_error"] = (f"The design checks could not run on this deck: "
                               f"{type(exc).__name__}: {exc}")
    return job.get("design_error")


def design_count(job) -> int | None:
    """How many open design decisions this deck has, for the badge on the audit
    report. None when the pass could not run - the link still appears, because a
    missing link reads as a missing feature."""
    if _ensure_design(job):
        return None
    answered = {a.finding_id for a in job.get("design_applied") or []}
    return sum(1 for f in job.get("design") or []
               if f.finding_id not in answered)


def _reaudit_in_place(job) -> str | None:
    """Re-run the audit over the deck as the design pass has left it, keeping
    the SAME job.

    In place, and that is the whole point: /reaudit mints a new job id, which
    would strand every design decision recorded against this one. The rule the
    audit flow already follows - verify after write, never assume - applies
    just as much to a change made from this page."""
    if job.get("deck") is None:
        return None
    fd, tmp_name = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        tmp.write_bytes(job["deck"])
        profile = job.get("profile_obj") or Profile.load(job["profile"])
        result = run_audit(tmp, profile)
    except Exception as exc:
        return (f"The deck was changed, but it could not be re-audited "
                f"({type(exc).__name__}: {exc}), so the audit report's counts "
                f"are from before these changes.")
    finally:
        tmp.unlink(missing_ok=True)
    manifest = result.to_manifest()
    manifest["deck"] = job["filename"]
    job["manifest"] = manifest
    return None


# How many slides are rendered per PowerPoint call. Flipping one slide at a time
# would mean one COM round trip per click (a second or more each); rendering the
# whole deck up front would mean minutes of wait on a 200-slide deck before the
# first slide appears. A window is neither: the first click pays for eight
# slides, the next seven are instant.
DESIGN_WINDOW = 8


def _ensure_design_shot(job, index: int) -> str | None:
    """Render the window of slides containing `index`, once. Returns a readable
    reason when nothing could be rendered - the page then shows every finding
    and every remedy without a picture, which is the part a designer cannot do
    without."""
    shots = job.setdefault("design_shots", {})
    if index in shots:
        return None
    if job.get("deck") is None:
        return "the deck is no longer held in memory."
    start = (index // DESIGN_WINDOW) * DESIGN_WINDOW
    total = job["manifest"]["slides"]
    wanted = [i for i in range(start, min(total, start + DESIGN_WINDOW))
              if i not in shots]
    if not wanted:
        return None
    from .render import export_decks_png

    lock = _thumb_locks.setdefault(f"design:{id(job)}", threading.Lock())
    with lock:
        if index in shots:
            return None
        try:
            images = export_decks_png({"deck": job["deck"]}, wanted,
                                      width=THUMB_WIDTH)
        except Exception as exc:
            job["design_shot_error"] = f"{type(exc).__name__}: {exc}."
            return job["design_shot_error"]
        for key, png in (images or {}).items():
            shots[int(key.split(":", 1)[1])] = png
    if index not in shots:
        return ("PowerPoint returned no image for this slide. Rendering needs "
                "desktop PowerPoint or LibreOffice on this machine.")
    return None


def _design_severity_map(findings, records, answered) -> dict:
    """{slide_index: {severity: n}} for the dots on the slide strip, from both
    passes: a slide the designer should stop at is one with anything on it, and
    which check found it is not the question the strip answers."""
    out: dict[int, dict] = {}
    for f in findings:
        if f.finding_id in answered:
            continue
        for index in f.slides:
            slot = out.setdefault(index, {})
            slot[f.severity] = slot.get(f.severity, 0) + 1
    for r in records or []:
        if r.get("module") == "preflight":
            continue
        slot = out.setdefault(r["slide_index"], {})
        slot[r["severity"]] = slot.get(r["severity"], 0) + 1
    return out


def _design_page(job_id: str, job, banner: str = "", error: str | None = None,
                 view: str = "slide", current: int = 0):
    from .design import slide_rects
    from .ui_design import render_design

    err = _ensure_design(job) or error
    meta = next((p for p in _profiles_meta() if p["id"] == job.get("profile")),
                None)
    applied = job.get("design_applied") or []
    answered = {a.finding_id for a in applied}
    every = [f for f in (job.get("design") or [])
             if f.finding_id not in answered]
    total = job["manifest"]["slides"]
    current = max(0, min(current, max(0, total - 1)))
    records = [r for r in job["manifest"]["records"]
               if r.get("module") != "preflight"]

    # A finding that touches ONE slide belongs to that slide's page. One that
    # spans several is a single decision about the deck and cannot honestly be
    # asked on any one of them, so it goes to the deck view (and is counted
    # there, on the tab, so it is not lost).
    on_slide = [f for f in every if f.slides == [current]]
    deck_wide = [f for f in every if len(f.slides) != 1]

    shot_error = None
    rects: list = []
    if view != "deck":
        shot_error = _ensure_design_shot(job, current)
        if shot_error is None and job.get("deck") is not None:
            try:
                rects = slide_rects(job["deck"], on_slide).get(current, [])
            except Exception:
                rects = []

    return HTMLResponse(render_design(
        deck_name=job["filename"],
        profile_name=(meta or {}).get("name") or job.get("profile") or "",
        job_id=job_id, view=view, current=current, total_slides=total,
        findings=on_slide, deck_findings=deck_wide, applied=applied,
        audit_records=[r for r in records if r["slide_index"] == current],
        rects=rects,
        per_slide=_design_severity_map(every, records, answered),
        banner=banner, error=err, render_error=shot_error,
        has_deck=job.get("deck") is not None))


def _invalidate_renders(job) -> None:
    """Everything derived from the deck bytes, dropped. A cached thumbnail beside
    a row marked applied is a picture of the deck before the fix."""
    job["diff"] = None
    job["thumbs"] = None
    job["rects"] = None
    job["design"] = None
    job["design_shots"] = {}
    job.pop("design_shot_error", None)


@app.get("/audit/{job_id}", response_class=HTMLResponse)
def audit_view(request: Request, job_id: str):
    """The audit report for a job that already ran.

    The report used to exist only as the response to the POST that produced it,
    so navigating away from it - to the design page, say - was one-way. A page a
    designer cannot get back to is a page they will not leave."""
    from .promotion import promoted_issue_types
    from .store import comment_counts

    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."),
                            status_code=404)
    return render_report(job["manifest"], job_id,
                         can_fix=job.get("deck") is not None,
                         has_cleaned=job.get("cleaned") is not None,
                         promoted=promoted_issue_types(),
                         comments=comment_counts(job["filename"]),
                         design=design_count(job),
                         assist=AI_ENABLED and not _ephemeral(job["profile"]))


@app.get("/design/{job_id}", response_class=HTMLResponse)
def design_page(job_id: str, view: str = "slide", n: int = 0):
    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."),
                            status_code=404)
    return _design_page(job_id, job,
                        view="deck" if view == "deck" else "slide",
                        current=max(0, n))


@app.get("/design-img/{job_id}/{idx}.png")
def design_img(job_id: str, idx: int):
    """One rendered slide for the design page.

    Its own route and its own cache rather than /thumb: that one is filled by
    _ensure_thumbs, which renders the WHOLE deck and whose callers (the visual
    PDF, the preview overlay) rely on it being complete. This page fills a
    window at a time, and a half-filled thumbs cache would make those callers
    think they had every slide."""
    from fastapi.responses import Response

    job = _job(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job"}, status_code=404)
    error = _ensure_design_shot(job, idx)
    png = (job.get("design_shots") or {}).get(idx)
    if png is None:
        return JSONResponse({"error": error or "no such slide"}, status_code=503)
    return Response(content=png, media_type="image/png",
                    headers={"Cache-Control": "private, max-age=600"})


@app.post("/design/{job_id}/apply", response_class=HTMLResponse)
async def design_apply(request: Request, job_id: str):
    """Perform the remedy picked on each card, and nothing else.

    The radio group per finding is named pick_<finding_id>, so the form is read
    rather than declared: the set of findings is not known until the deck is
    scanned, and a fixed signature would have to be a list of opaque pairs."""
    from .remedy import apply as apply_remedies

    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."),
                            status_code=404)
    if job.get("deck") is None:
        return _design_page(job_id, job)

    _ensure_design(job)
    by_id = {f.finding_id: f for f in job.get("design") or []}
    already = {a.finding_id for a in job.get("design_applied") or []}
    form = await request.form()
    try:
        back_to = max(0, int(form.get("n") or 0))
    except (TypeError, ValueError):
        back_to = 0
    picks = []
    for key, value in form.multi_items():
        if not str(key).startswith("pick_"):
            continue
        finding = by_id.get(str(key)[5:])
        if finding is None or finding.finding_id in already:
            continue
        remedy = next((o for o in finding.options
                       if o.remedy_id == str(value)), None)
        if remedy is not None:
            picks.append((finding, remedy))
    if not picks:
        return _design_page(job_id, job, current=back_to,
                            banner="No card had an answer picked, so nothing "
                                   "was changed.")

    try:
        deck, applied = apply_remedies(job["deck"], picks)
    except Exception as exc:
        return _design_page(job_id, job, current=back_to,
                            error=f"Nothing was changed: {type(exc).__name__}: "
                                  f"{exc}")

    acted = [a for a in applied if a.done and a.undo]
    left = [a for a in applied if a.done and not a.undo]
    failed = [a for a in applied if not a.done]
    with _jobs_lock:
        if acted:
            job["prev_deck"] = job["deck"]   # pre-change bytes, for the diff
            job["deck"] = deck
            job["cleaned"] = deck
        job["design_applied"] = (job.get("design_applied") or []) + applied
        _invalidate_renders(job)

    note = []
    if acted:
        note.append(f"Applied {len(acted)} change"
                    f"{'s' if len(acted) != 1 else ''}")
    if left:
        note.append(f"recorded {len(left)} as deliberate")
    if failed:
        note.append(f"{len(failed)} could not be applied and say why below")
    stale = _reaudit_in_place(job) if acted else None
    return _design_page(job_id, job, current=back_to,
                        banner="; ".join(note) + ".", error=stale)


@app.post("/design/{job_id}/undo", response_class=HTMLResponse)
def design_undo(job_id: str, finding_ids: list[str] = Form(None),
                n: int = Form(None)):
    """Take one decision back, exactly, and say what came with it."""
    from .remedy import followers, undo_items
    from .undo import apply_undo

    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."),
                            status_code=404)
    # Which view to answer on: the slide the button was pressed on, or the
    # deck-wide tab when it carried no slide.
    view = "slide" if n is not None else "deck"
    back_to = max(0, n or 0)
    applied = list(job.get("design_applied") or [])
    wanted = set(finding_ids or [])
    if not wanted:
        return _design_page(job_id, job, view=view, current=back_to)

    # Chains are unioned before anything is replayed: two requested decisions
    # can drag the same third one, and undoing it twice would put back a state
    # from before the first replay.
    chain_ids: set = set()
    for finding_id in wanted:
        chain_ids.update(a.finding_id for a in followers(applied, finding_id))
    chain = [a for a in applied if a.finding_id in chain_ids]
    dragged = [a for a in chain if a.finding_id not in wanted]

    error = None
    outcomes = []
    items = undo_items([a for a in chain if a.undo])
    if items and job.get("deck") is not None:
        try:
            deck, outcomes = apply_undo(job["deck"], items)
        except Exception as exc:
            error = f"The undo failed and nothing was changed: {type(exc).__name__}: {exc}"
        else:
            with _jobs_lock:
                job["deck"] = deck
                job["cleaned"] = deck
                _invalidate_renders(job)
    elif items:
        error = ("The deck is no longer in memory, so these changes cannot be "
                 "reversed here. The decision has been cleared from the list.")

    if error is None or not items:
        with _jobs_lock:
            job["design_applied"] = [a for a in applied
                                     if a.finding_id not in chain_ids]
        if items:
            _reaudit_in_place(job)

    put_back = sum(1 for o in outcomes if o.get("done"))
    note = [f"{len(chain)} decision{'s' if len(chain) != 1 else ''} reopened"]
    if put_back:
        note.append(f"{put_back} change{'s' if put_back != 1 else ''} put back "
                    f"exactly as {'they were' if put_back != 1 else 'it was'}")
    if dragged:
        note.append(f"including {len(dragged)} that touched the same shape and "
                    f"could not come back on its own")
    # Back to the view the Undo button was on. Returning a designer to slide 1
    # after they pressed a button on slide 7 is the small rudeness that makes a
    # tool tiring to use.
    return _design_page(job_id, job, view=view, current=back_to,
                        banner="; ".join(note) + ".", error=error)


_thumb_locks: dict[str, threading.Lock] = {}
THUMB_WIDTH = 1100


def _ensure_thumbs(job_id: str, job: dict) -> None:
    """Render every slide of the job's current deck once (PowerPoint COM),
    cache in memory. Raises RuntimeError when rendering is unavailable."""
    if job.get("thumbs"):
        return
    if job.get("deck") is None:
        raise RuntimeError("deck bytes no longer in memory; re-upload to preview")
    from .render import export_decks_png

    lock = _thumb_locks.setdefault(job_id, threading.Lock())
    with lock:
        if job.get("thumbs"):
            return
        n = job["manifest"]["slides"]
        images = export_decks_png({"deck": job["deck"]}, list(range(n)),
                                  width=THUMB_WIDTH)
        job["thumbs"] = {int(k.split(":", 1)[1]): v for k, v in images.items()}


def _ensure_rects(job: dict) -> dict:
    if job.get("rects") is None:
        from .render import audit_rects

        flagged = [r for r in job["manifest"]["records"]
                   if r["module"] != "preflight"]
        job["rects"] = audit_rects(job["deck"], flagged)
    return job["rects"]


@app.get("/signin", response_class=HTMLResponse)
def signin_page():
    from .store import list_users
    from .ui import render_signin

    return render_signin(list_users())


@app.post("/signin", response_class=HTMLResponse)
def signin(request: Request, name: str = Form(...), pin: str = Form(...)):
    from fastapi.responses import RedirectResponse

    from .store import (create_session, get_user, has_pin, list_users,
                        set_pin, verify_pin)
    from .ui import render_signin

    if get_user(name) is None:
        return HTMLResponse(render_signin(list_users(), "Unknown name."),
                            status_code=404)
    if not has_pin(name):
        try:
            set_pin(name, pin)  # first sign-in sets the PIN
        except ValueError as exc:
            return HTMLResponse(render_signin(list_users(), str(exc)),
                                status_code=400)
    elif not verify_pin(name, pin):
        return HTMLResponse(render_signin(list_users(), "Wrong PIN."),
                            status_code=401)
    resp = RedirectResponse("/", status_code=303)
    resp.set_cookie("qc_session", create_session(name), max_age=30 * 24 * 3600,
                    httponly=True)
    resp.delete_cookie("qc_user")
    return resp


@app.post("/signout")
def signout(request: Request):
    from .store import delete_session

    delete_session(request.cookies.get("qc_session", ""))
    resp = JSONResponse({"ok": True})
    resp.delete_cookie("qc_session")
    resp.delete_cookie("qc_user")
    return resp


@app.get("/history", response_class=HTMLResponse)
def history_page():
    from .store import list_audits
    from .ui import render_history

    return render_history(list_audits())


@app.get("/history/{audit_id}", response_class=HTMLResponse)
def history_view(audit_id: int):
    from .store import comment_counts, get_audit
    from .ui import render_history

    audit = get_audit(audit_id)
    if audit is None:
        return HTMLResponse(render_history([]), status_code=404)
    manifest = audit["manifest"]
    when = audit["created_at"].replace("T", " ")[:16]
    return render_report(
        manifest, job_id="", can_fix=False, archived=True,
        banner=(f"Archived {audit['kind']} from {when} UTC by "
                f"{audit['user_name']}. Read-only record; re-upload the deck "
                f"to work on it."),
        comments=comment_counts(manifest["deck"]))


@app.get("/me")
def me(request: Request):
    from .auth import current_user
    from .store import list_users

    user = current_user(request)
    if user:
        user = {k: v for k, v in user.items() if k != "pin_hash"}
    return JSONResponse({"user": user, "users": [
        {k: v for k, v in u.items() if k != "pin_hash"} for u in list_users()]})


@app.post("/whoami")
async def whoami(request: Request):
    from .store import get_user

    data = await request.json()
    user = get_user(data.get("name", ""))
    if user is None:
        return JSONResponse({"error": "unknown user"}, status_code=404)
    resp = JSONResponse({"ok": True, "user": user})
    resp.set_cookie("qc_user", user["name"], max_age=90 * 24 * 3600)
    return resp


@app.get("/comments")
def get_comments(deck: str, slide: int):
    from .store import comments_for

    return JSONResponse({"comments": comments_for(deck, slide)})


@app.post("/comments")
async def post_comment(request: Request):
    from .store import add_comment, get_user

    from .auth import current_user

    user = current_user(request)
    if user is None:
        return JSONResponse({"error": "sign in first (top right)"},
                            status_code=401)
    author = user["name"]
    data = await request.json()
    try:
        comment = add_comment(deck=data["deck"], slide_index=int(data["slide_index"]),
                              author=author, text=data.get("text", ""),
                              record_id=data.get("record_id"))
    except (KeyError, ValueError) as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)
    return JSONResponse({"ok": True, "comment": comment})


@app.post("/triage")
async def triage(request: Request):
    from .triage import log_triage

    data = await request.json()
    job = _job(data.get("job_id", ""))
    state = data.get("state", "")
    if job is None or state not in ("confirmed", "false_positive", "cleared"):
        return JSONResponse({"error": "bad request"}, status_code=400)
    record = next((r for r in job["manifest"]["records"]
                   if r["record_id"] == data.get("record_id")), None)
    if record is None:
        return JSONResponse({"error": "unknown record"}, status_code=404)

    states = job.setdefault("triage", {})
    if state == "cleared":
        states.pop(record["record_id"], None)
    else:
        states[record["record_id"]] = state
    log_triage(record, state, job["filename"], job["manifest"]["profile_id"])
    counts = {"confirmed": sum(1 for v in states.values() if v == "confirmed"),
              "false_positive": sum(1 for v in states.values() if v == "false_positive")}
    return JSONResponse({"ok": True, "counts": counts})


@app.get("/stats", response_class=HTMLResponse)
def stats_page():
    from .triage import stats
    from .ui import render_stats

    return render_stats(stats())


@app.post("/assist/{job_id}")
def assist_questions(job_id: str):
    """Build clarifying questions from this job's findings. The issued
    question set (with actions) is held server-side; /assist/apply accepts
    only question ids, never client-supplied action payloads."""
    from .assist import aggregate, generate_questions

    off = _ai_disabled_response()
    if off is not None:
        return off

    job = _job(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job id"}, status_code=404)
    if _ephemeral(job["profile"]):
        which = ("self-consistency audit" if job["profile"] == SELF_PROFILE
                 else "master you uploaded")
        extra = ("" if job["profile"] == SELF_PROFILE else
                 " Save the master as a profile from the Read a master page "
                 "first, then audit against it.")
        return JSONResponse(
            {"error": f"The assistant tunes a saved profile; the {which} has "
                      f"none.{extra}"}, status_code=400)
    profile_obj = job.get("profile_obj")
    agg = aggregate(job["manifest"], job.get("deck"),
                    profile_obj.config if profile_obj is not None else None)
    questions, source = generate_questions(agg)
    job["assist"] = {q["id"]: q for q in questions}
    public = [{k: q[k] for k in ("id", "question", "rationale", "impact")}
              for q in questions]
    return {"questions": public, "source": source,
            "profile": job["profile"]}


@app.post("/assist/{job_id}/apply")
def assist_apply(request: Request, job_id: str,
                 accepted: list[str] = Form(None)):
    from .assist import apply_actions
    from .web_admin import _editor

    off = _ai_disabled_response()
    if off is not None:
        return off

    job = _job(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job id"}, status_code=404)
    editor = _editor(request)
    if editor is None:
        return JSONResponse(
            {"error": "Profile changes need a lead or admin; sign in with "
                      "such a role first."}, status_code=403)
    issued = job.get("assist") or {}
    actions = [issued[qid]["action"] for qid in (accepted or [])
               if qid in issued]
    if not actions:
        return JSONResponse({"error": "no accepted questions"}, status_code=400)
    result = apply_actions(job["profile"], actions, editor["name"])
    return {"profile": job["profile"], **result}


@app.post("/reaudit/{job_id}", response_class=HTMLResponse)
def reaudit(request: Request, job_id: str):
    """Re-run the audit on the retained deck bytes with the profile as it
    is NOW (e.g. right after assistant answers updated it)."""
    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."), status_code=404)
    if job["deck"] is None:
        return HTMLResponse(render_index(
            _pickable_profiles(), MODULES,
            "The deck is no longer held in memory; re-upload it."), status_code=410)

    fd, tmp_name = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        tmp.write_bytes(job["deck"])
        # A master-derived profile cannot be re-resolved: the master's bytes
        # were dropped after extraction, by design. Reuse the profile the job
        # already carries. Saved profiles still re-resolve, which is the whole
        # point of this route (it picks up assistant edits).
        profile_obj = (job["profile_obj"] if job["profile"] == MASTER_PROFILE
                       else _resolve_profile(job["profile"], job["deck"]))
        result = run_audit(tmp, profile_obj)
    finally:
        tmp.unlink(missing_ok=True)

    manifest = result.to_manifest()
    manifest["deck"] = job["filename"]
    new_id = uuid.uuid4().hex
    with _jobs_lock:
        _jobs[new_id] = {"manifest": manifest, "deck": job["deck"],
                         "cleaned": None, "filename": job["filename"],
                         "profile": job["profile"], "profile_obj": profile_obj,
                         "master_spec": job.get("master_spec")}
    from .auth import current_user
    from .promotion import promoted_issue_types
    from .store import comment_counts, record_audit

    user = current_user(request)
    record_audit(manifest, user["name"] if user else "anonymous")
    return render_report(manifest, new_id, can_fix=True,
                         promoted=promoted_issue_types(),
                         comments=comment_counts(job["filename"]),
                         assist=AI_ENABLED and not _ephemeral(job["profile"]),
                         banner="Re-audited with the updated profile.")


@app.post("/copilot/{job_id}", response_class=HTMLResponse)
def copilot(request: Request, job_id: str):
    """Design copilot: render the slides, let Claude (vision) propose layout
    actions, verify them in code, and merge the survivors into the report
    as ordinary tickable suggestions."""
    from .assist import api_configured
    off = _ai_disabled_response()
    if off is not None:
        return off

    from .copilot import run_copilot

    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."), status_code=404)
    if job["deck"] is None:
        return HTMLResponse(render_index(
            _pickable_profiles(), MODULES,
            "The deck is no longer held in memory; re-upload it."), status_code=410)

    def _report(banner):
        from .promotion import promoted_issue_types
        from .store import comment_counts

        return render_report(job["manifest"], job_id, can_fix=True,
                             banner=banner, promoted=promoted_issue_types(),
                             comments=comment_counts(job["filename"]),
                             assist=AI_ENABLED and not _ephemeral(job["profile"]))

    if not api_configured():
        return _report("Design copilot needs an Anthropic API key "
                       "(ANTHROPIC_API_KEY) on the server.")
    try:
        _ensure_thumbs(job_id, job)
    except RuntimeError as exc:
        return _report(f"Design copilot needs slide renders: {exc}")

    new_records, reviewed = run_copilot(job["deck"], job["thumbs"],
                                        job["manifest"])
    if new_records:
        from collections import Counter

        manifest = job["manifest"]
        manifest["records"].extend(new_records)
        records = manifest["records"]
        manifest["summary"] = {
            "by_severity": dict(Counter(r["severity"] for r in records)),
            "by_issue_type": dict(Counter(r["issue_type"] for r in records)),
            "by_module": dict(Counter(r["module"] for r in records)),
            "arabic_flagged": sum(1 for r in records if r["arabic_flag"]),
            "total": len(records),
        }
        job["rects"] = None  # pins are renumbered on next preview
    return _report(f"Design copilot reviewed {reviewed} slide"
                   f"{'s' if reviewed != 1 else ''} and added "
                   f"{len(new_records)} suggestion"
                   f"{'s' if len(new_records) != 1 else ''} "
                   "(tickable, never pre-selected).")


@app.get("/slide/{job_id}/{idx}")
def slide_preview(job_id: str, idx: int):
    job = _job(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job id"}, status_code=404)
    try:
        _ensure_thumbs(job_id, job)
    except RuntimeError as exc:
        return JSONResponse({"error": str(exc)}, status_code=503)
    if idx not in job["thumbs"]:
        return JSONResponse({"error": "no such slide"}, status_code=404)
    rects = _ensure_rects(job).get(idx, [])
    return JSONResponse({"png": f"/thumb/{job_id}/{idx}.png", "rects": rects})


@app.get("/thumb/{job_id}/{idx}.png")
def thumb_png(job_id: str, idx: int):
    from fastapi.responses import Response

    job = _job(job_id)
    png = ((job or {}).get("thumbs") or {}).get(idx)
    if png is None:
        return JSONResponse({"error": "no such render"}, status_code=404)
    return Response(content=png, media_type="image/png",
                    headers={"Cache-Control": "private, max-age=600"})


@app.get("/diff/{job_id}.pdf")
def diff_pdf(job_id: str):
    from fastapi.responses import Response

    from .render import build_diff
    from .report import render_diff_pdf

    job = _job(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job id"}, status_code=404)
    if not job.get("applied_records") or job.get("prev_deck") is None:
        return JSONResponse({"error": "no applied fixes to compare yet"},
                            status_code=404)
    if job.get("diff") is None:
        try:
            job["diff"] = build_diff(job["prev_deck"], job["cleaned"],
                                     job["applied_records"])
        except RuntimeError as exc:
            return JSONResponse({"error": str(exc)}, status_code=503)
    pdf = render_diff_pdf(job["filename"], job["diff"])
    name = Path(job["filename"]).stem
    return Response(content=pdf, media_type="application/pdf",
                    headers=_attachment(f"review-{name}.pdf"))


@app.get("/diff/{job_id}", response_class=HTMLResponse)
def diff(job_id: str):
    from .render import build_diff
    from .ui import render_diff

    job = _job(job_id)
    if job is None:
        return HTMLResponse(render_index(_pickable_profiles(), MODULES,
                                         "Unknown or expired job."), status_code=404)
    if not job.get("applied_records") or job.get("prev_deck") is None:
        return render_diff(job["filename"], job_id, None,
                           error="No applied fixes to compare yet. Apply fixes first.")
    if job.get("diff") is None:
        try:
            # first visit renders via desktop PowerPoint; cached afterwards
            job["diff"] = build_diff(job["prev_deck"], job["cleaned"],
                                     job["applied_records"])
        except RuntimeError as exc:
            return render_diff(job["filename"], job_id, None, error=str(exc))
    return render_diff(job["filename"], job_id, job["diff"])


@app.get("/render/{job_id}/{key}.png")
def render_png(job_id: str, key: str):
    from fastapi.responses import Response

    job = _job(job_id)
    images = (job or {}).get("diff", {}) or {}
    png = images.get("images", {}).get(key.replace("-", ":", 1))
    if png is None:
        return JSONResponse({"error": "no such render"}, status_code=404)
    return Response(content=png, media_type="image/png",
                    headers={"Cache-Control": "private, max-age=3600"})


@app.get("/annotated/{job_id}")
def annotated(job_id: str):
    from fastapi.responses import Response

    from .annotate import build_annotated
    from .store import comments_for

    job = _job(job_id)
    if job is None or job.get("deck") is None:
        return JSONResponse({"error": "deck no longer in memory; re-upload"},
                            status_code=404)
    out = build_annotated(job["deck"], job["manifest"],
                          comments_for(job["filename"]))
    name = Path(job["filename"]).stem
    return Response(
        content=out,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=_attachment(f"{name}.annotated.pptx"))


@app.get("/download/{job_id}")
def download(job_id: str):
    from fastapi.responses import Response

    job = _job(job_id)
    if job is None or job["cleaned"] is None:
        return JSONResponse({"error": "no cleaned deck for this job"}, status_code=404)
    name = Path(job["filename"]).stem
    return Response(
        content=job["cleaned"],
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=_attachment(f"{name}.cleaned.pptx"))


@app.get("/report/{job_id}.pdf")
def report_pdf(job_id: str):
    from fastapi.responses import Response

    from .report import render_pdf, render_visual_audit_pdf

    job = _job(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job id"}, status_code=404)
    content = None
    if job.get("deck") is not None:
        try:
            # visual report: rendered slides with the findings highlighted
            _ensure_thumbs(job_id, job)
            content = render_visual_audit_pdf(job["manifest"], job["thumbs"],
                                              _ensure_rects(job))
        except RuntimeError:
            content = None  # rendering unavailable: text-only fallback
    if content is None:
        content = render_pdf(job["manifest"])
    return Response(content=content, media_type="application/pdf",
                    headers=_attachment(f"audit-{Path(job['filename']).stem}.pdf"))


@app.get("/report/{job_id}.csv")
def report_csv(job_id: str):
    from fastapi.responses import Response

    from .report import render_csv

    job = _job(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job id"}, status_code=404)
    return Response(content=render_csv(job["manifest"]), media_type="text/csv",
                    headers=_attachment(f"audit-{Path(job['filename']).stem}.csv"))


@app.get("/manifest/{job_id}")
def manifest(job_id: str):
    # single atomic lookup: eviction by a concurrent /audit must yield a
    # clean 404, not a KeyError 500 (review finding)
    job = _jobs.get(job_id)
    if job is None:
        return JSONResponse({"error": "unknown or expired job id"}, status_code=404)
    return JSONResponse(job["manifest"])


def main():
    import argparse
    import socket

    import uvicorn

    from . import auth, web

    ap = argparse.ArgumentParser(description="Prezlab PPT QC web app.")
    ap.add_argument("--host", default=None,
                    help="Bind address (default 127.0.0.1, or 0.0.0.0 with --lan).")
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--lan", action="store_true",
                    help="LAN pilot deployment: bind all interfaces, REQUIRE "
                         "sign-in on every route, and disable the legacy "
                         "name-cookie identity.")
    ap.add_argument("--reload", action="store_true",
                    help="Restart on source changes. Use while developing: "
                         "without it a running server keeps serving the code "
                         "it was started with, and edits appear to do nothing.")
    args = ap.parse_args()

    host = args.host or ("0.0.0.0" if args.lan else "127.0.0.1")
    if args.lan:
        app.state.auth_required = True  # single identity across -m double-import
        web.AUTH_REQUIRED = True
        auth.STRICT_SESSIONS = True
        try:
            probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            probe.connect(("10.255.255.255", 1))
            ip = probe.getsockname()[0]
            probe.close()
        except OSError:
            ip = "this-machine"
        print(f"LAN mode: sign-in enforced on every route.")
        print(f"Team URL: http://{ip}:{args.port}")
    elif host != "127.0.0.1":
        print("WARNING: binding beyond loopback WITHOUT --lan leaves routes "
              "open to anonymous use. Use --lan for network exposure.")
    # Printed unconditionally: log_level="warning" suppresses uvicorn's own
    # startup line, so without this the server looks hung, and there is no
    # visible marker of WHEN it started or whether it predates a code change.
    print(f"Serving http://{'localhost' if host == '127.0.0.1' else host}:"
          f"{args.port}   (Ctrl+C to stop)")
    if args.reload:
        # Reload needs an import string, not the app object, so uvicorn can
        # re-import the module in a fresh worker.
        uvicorn.run("qc.web:app", host=host, port=args.port,
                    log_level="warning", reload=True)
    else:
        print("Code changes need a restart; start with --reload to pick them "
              "up automatically.")
        uvicorn.run(app, host=host, port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
