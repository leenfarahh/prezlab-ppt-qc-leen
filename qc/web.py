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
        _format_jobs[job_id] = {"deck": result.deck, "filename": filename,
                                "profile": profile, "plans": result.plans,
                                "errors": result.errors,
                                "applied": result.applied,
                                "masters": result.masters,
                                "stragglers": result.stragglers,
                                "space_notes": space_notes,
                                "changes": content_changes, "restored": []}
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
                         assist=AI_ENABLED and not _ephemeral(job["profile"]))


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
