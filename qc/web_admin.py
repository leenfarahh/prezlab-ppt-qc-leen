"""The team roster and the profile editor.

Attached to the app via register(app) from qc/web.py. Saving anything here is
role-gated to leads/admins through the pilot identity cookie; this is
attribution-grade control (real auth is Entra ID in the production tier).

THE PROFILE EDITOR IS BACK (design lead, 31/08/2026). It was removed on the
argument that a profile is created where a designer already is - step 1 of
Prepare a deck, with the master in front of them - and that is true of CREATING
one. It is not true of changing one. Every edit after the first went through
re-reading the master, which rebuilds the whole document and throws away
everything decided against it since, so a profile that needed one number changed
either got a full rebuild or got left wrong.

Creation still happens on Prepare a deck, and the helpers below are still called
from that route. What is new is /profiles: a list, and a typed form per profile
(qc.profileform, qc.ui_profiles).
"""

import json
import re

from fastapi import Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

from .ui import esc
from .ui_admin import render_admin_error, render_team

_HEX = re.compile(r"^[0-9A-Fa-f]{6}$")


def _slugify(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "_", name.strip().lower()).strip("_")
    return slug or ""


def _unique_pid(base: str) -> str:
    """A slug not already taken by an existing profile file."""
    from .profile import PROFILES_DIR

    pid = base
    n = 2
    while (PROFILES_DIR / f"{pid}.json").exists():
        pid = f"{base}_{n}"
        n += 1
    return pid


def _profile_path(pid: str):
    from .profile import PROFILES_DIR

    if not re.match(r"^[A-Za-z0-9_\-]+$", pid):
        return None
    return PROFILES_DIR / f"{pid}.json"


def _load_profile(pid: str) -> dict | None:
    path = _profile_path(pid)
    if path is None or not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _editor(request: Request) -> dict | None:
    """The signed-in pilot user, only if their role may save profiles."""
    from .auth import current_user

    user = current_user(request)
    if user and user["role"] in ("lead", "admin"):
        return user
    return None


def replace_profile_master(pid: str, blob: bytes, filename: str,
                           editor_name: str) -> tuple[bool, str]:
    """Point an existing profile at a revised master, and re-read its frame.

    A profile's master is stored once, when the profile is created, and the
    rebuild hands PowerPoint that stored copy. Everything the designer adds to
    their master afterwards - a presentation-space rectangle, a moved guide, a
    new layout - reached no deck at all until this existed (design lead,
    21/08/2026).

    Only what the master STATES is re-read. Fonts, palette and tolerances stay
    exactly as they were edited: those are decisions about the client, not
    readings of the file, and silently reverting them to a fresh projection
    would punish every edit ever made.

    A module-level function with two callers: step 1 of Prepare a deck, where a
    designer already has the master open, and the profile editor's own replace
    button. One implementation, because "which fields does replacing a master
    touch" is a question that must have exactly one answer.

    Returns (saved, note). `note` carries HTML, and on a failure it is the
    reason.
    """
    import io
    import json
    from datetime import date

    from pptx import Presentation

    from .stylespec import (dominant_master, extract_style_spec,
                            spec_to_profile)
    from .templates import save_master

    data = _load_profile(pid)
    if data is None:
        return False, f"No profile named '{esc(pid)}'."

    try:
        prs = Presentation(io.BytesIO(blob))
        if dominant_master(prs) is None:
            raise ValueError("the file carries no slide master")
        spec = extract_style_spec(prs, source=filename, embed_assets=False)
        fresh = spec_to_profile(spec, pid, data.get("name", pid))["config"]
    except Exception as exc:
        return False, (f"Could not read that master: "
                       f"{esc(type(exc).__name__)}: {esc(str(exc))}")

    cfg = data.setdefault("config", {})
    geo = cfg.setdefault("geometry", {})
    fresh_geo = fresh.get("geometry", {})
    changes = []

    old_margins = geo.get("safe_zone_margins_emu") or {}
    new_margins = fresh_geo.get("safe_zone_margins_emu") or {}
    if new_margins and new_margins != old_margins:
        geo["safe_zone_margins_emu"] = new_margins
        changes.append("margins " + ", ".join(
            f"{side} {old_margins.get(side, 0) / 914400:.2f}in "
            f"&rarr; {new_margins[side] / 914400:.2f}in"
            for side in ("left", "top", "right", "bottom")
            if old_margins.get(side) != new_margins.get(side)))

    if fresh_geo.get("body_band_emu") != geo.get("body_band_emu"):
        geo["body_band_emu"] = fresh_geo.get("body_band_emu")
        band = geo["body_band_emu"]
        changes.append(
            f"reserved header band "
            f"{band['subtitle_floor'] / 914400:.2f}in&ndash;"
            f"{band['body_top'] / 914400:.2f}in" if band
            else "reserved header band dropped: the master no longer states one")

    if fresh_geo.get("grid") != geo.get("grid"):
        geo["grid"] = fresh_geo.get("grid")
        changes.append("column grid")

    names = [lay["name"] for lay in spec.get("layouts", [])]
    ms = cfg.setdefault("master_slide", {})
    if names != ms.get("layout_allowlist"):
        added = [n for n in names if n not in (ms.get("layout_allowlist") or [])]
        dropped = [n for n in (ms.get("layout_allowlist") or [])
                   if n not in names]
        ms["layout_allowlist"] = names
        changes.append(f"layouts ({len(added)} added, {len(dropped)} removed)")

    source = fresh.get("style_spec_source") or {}
    was = (cfg.get("style_spec_source") or {}).get("grid_source")
    cfg["style_spec_source"] = source
    if source.get("grid_source") != was:
        changes.append(f"frame now read from "
                       f"{str(source.get('grid_source')).replace('_', ' ')}")

    save_master(pid, blob)
    data["version"] = int(data.get("version", 1)) + 1
    data["owner"] = (f"{editor_name} (master replaced from {filename}, "
                     f"{date.today().strftime('%d/%m/%Y')})")
    _profile_path(pid).write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    return True, (f"Master replaced on <b>{esc(data.get('name', pid))}</b> from "
                  f"<b>{esc(filename)}</b>. "
                  + ("Re-read: " + "; ".join(changes) + "."
                     if changes else
                     "The rules it states are unchanged, so only the file was "
                     "replaced - which is still what every deck prepared "
                     "against this profile will now be built on."))


def _all_profiles() -> list:
    """Every saved profile as the list page reads it, by name."""
    from .profile import PROFILES_DIR
    from .templates import has_master
    from .ui_profiles import profile_row

    out = []
    for path in sorted(PROFILES_DIR.glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue          # a hand-edited profile must not blank the page
        out.append(profile_row(path.stem, data, has_master(path.stem)))
    return sorted(out, key=lambda p: p["name"].casefold())


def _clear_other_defaults(pid: str) -> None:
    """Only one profile is offered first. Enforced on save rather than trusted,
    because two defaults is a picker that chooses arbitrarily."""
    from .profile import PROFILES_DIR

    for path in PROFILES_DIR.glob("*.json"):
        if path.stem == pid:
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if data.get("is_default"):
            data["is_default"] = False
            path.write_text(json.dumps(data, indent=2, ensure_ascii=False),
                            encoding="utf-8")


def register(app) -> None:
    # ------------------------------------------------------------- profiles

    def _denied(what: str, back: str):
        return HTMLResponse(render_admin_error(
            "Profiles", f"{what} needs a lead or admin",
            "Sign in as a lead or admin to change a profile. Reading one does "
            "not need a role.", back=back), status_code=403)

    def _missing(pid: str):
        return HTMLResponse(render_admin_error(
            "Profiles", "No such profile",
            f"There is no profile with the id '{pid}'.", back="/profiles"),
            status_code=404)

    @app.get("/profiles", response_class=HTMLResponse)
    def profiles(request: Request, saved: str = "", deleted: str = ""):
        from .ui_profiles import render_profiles

        ok = ""
        if saved:
            ok = f"Saved {saved}."
        elif deleted:
            ok = (f"Deleted {deleted}. Decks already audited against it keep "
                  f"their reports; nothing else changed.")
        return HTMLResponse(render_profiles(
            _all_profiles(), can_edit=_editor(request) is not None, ok=ok))

    @app.get("/profiles/{pid}", response_class=HTMLResponse)
    def profile_edit(request: Request, pid: str, saved: str = ""):
        from .templates import has_master, master_info
        from .ui_profiles import render_profile

        data = _load_profile(pid)
        if data is None:
            return _missing(pid)
        info = master_info(pid)
        note = ("This profile carries no master file, so it can audit a deck "
                "but cannot be applied to one.")
        if info:
            note = (f"{info['bytes'] // 1024} KB, fingerprint "
                    f"{info['sha1'][:10]}. Replacing it re-reads the frame, "
                    f"reserved bands, grid and layout names from the new file "
                    f"and leaves everything on this page alone.")
        return HTMLResponse(render_profile(
            data, pid=pid, can_edit=_editor(request) is not None,
            has_master=has_master(pid), master_note=note,
            ok="Saved." if saved else ""))

    @app.post("/profiles/{pid}", response_class=HTMLResponse)
    async def profile_save(request: Request, pid: str):
        from .profileform import Invalid, apply_form, partial
        from .templates import has_master, master_info
        from .ui_profiles import render_profile

        editor = _editor(request)
        if editor is None:
            return _denied("Saving a profile", f"/profiles/{pid}")
        data = _load_profile(pid)
        if data is None:
            return _missing(pid)

        form = await request.form()
        try:
            updated = apply_form(data, form)
        except Invalid as exc:
            # The EDITED document is re-rendered, not the stored one: a designer
            # who mistyped one hex code must not lose the other nine changes
            # they made in the same pass. apply_form works on a copy, so what
            # comes back here is the profile as they left it plus the fault.
            as_typed = partial(data, form)
            info = master_info(pid)
            return HTMLResponse(render_profile(
                as_typed, pid=pid, can_edit=True, has_master=has_master(pid),
                master_note=(f"{info['bytes'] // 1024} KB" if info else ""),
                errors={exc.field_name: str(exc)},
                message="Nothing was saved. One field needs fixing."),
                status_code=400)

        updated["owner"] = editor["name"]
        path = _profile_path(pid)
        path.write_text(json.dumps(updated, indent=2, ensure_ascii=False),
                        encoding="utf-8")
        if updated.get("is_default"):
            _clear_other_defaults(pid)
        return RedirectResponse(f"/profiles/{pid}?saved=1", status_code=303)

    @app.post("/profiles/{pid}/duplicate")
    def profile_duplicate(request: Request, pid: str):
        if _editor(request) is None:
            return _denied("Duplicating a profile", f"/profiles/{pid}")
        data = _load_profile(pid)
        if data is None:
            return _missing(pid)
        from .templates import load_master, save_master

        new_id = _unique_pid(_slugify(f"{data.get('name', pid)} copy")
                             or f"{pid}_copy")
        data["id"] = new_id
        data["name"] = f"{data.get('name', pid)} (copy)"
        data["version"] = 1
        # A duplicate that shares the original's default flag would give the
        # picker two answers on the press that created it.
        data["is_default"] = False
        _profile_path(new_id).write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        blob = load_master(pid)
        if blob:
            save_master(new_id, blob)
        return RedirectResponse(f"/profiles/{new_id}", status_code=303)

    @app.post("/profiles/{pid}/delete")
    def profile_delete(request: Request, pid: str):
        if _editor(request) is None:
            return _denied("Deleting a profile", f"/profiles/{pid}")
        data = _load_profile(pid)
        if data is None:
            return _missing(pid)
        from .templates import delete_master

        name = data.get("name") or pid
        delete_master(pid)
        _profile_path(pid).unlink(missing_ok=True)
        return RedirectResponse(f"/profiles?deleted={name}", status_code=303)

    @app.post("/profiles/{pid}/master")
    async def profile_replace_master(request: Request, pid: str):
        if _editor(request) is None:
            return _denied("Replacing a master", f"/profiles/{pid}")
        form = await request.form()
        upload = form.get("master")
        if upload is None or not getattr(upload, "filename", ""):
            return RedirectResponse(f"/profiles/{pid}", status_code=303)
        blob = await upload.read()
        ok, note = replace_profile_master(
            pid, blob, upload.filename, _editor(request)["name"])
        if not ok:
            return HTMLResponse(render_admin_error(
                "Profiles", "That master could not be read", note,
                back=f"/profiles/{pid}"), status_code=400)
        return RedirectResponse(f"/profiles/{pid}?saved=1", status_code=303)

    # ----------------------------------------------------------------- team

    @app.get("/team", response_class=HTMLResponse)
    def team():
        from .store import ROLES, list_users

        return render_team(list_users(), ROLES)

    @app.post("/team/add")
    def team_add(name: str = Form(...), role: str = Form(...)):
        from .store import ROLES, add_user, list_users

        try:
            add_user(name, role)
        except ValueError as exc:
            return HTMLResponse(render_team(list_users(), ROLES, str(exc)),
                                status_code=400)
        return RedirectResponse("/team", status_code=303)

    @app.post("/team/reset-pin")
    def team_reset_pin(request: Request, name: str = Form(...)):
        from .store import clear_pin, get_user

        if _editor(request) is None:
            return HTMLResponse(render_admin_error(
                "Team", "PIN reset needs a lead or admin",
                "Sign in as a lead or admin to reset PINs.",
                back="/team"), status_code=403)
        if get_user(name) is None:
            return HTMLResponse(render_admin_error(
                "Team", "Not found", f"No team member named '{name}'.",
                back="/team"), status_code=404)
        clear_pin(name)
        return RedirectResponse("/team", status_code=303)
