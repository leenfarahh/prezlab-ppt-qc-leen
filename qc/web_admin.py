"""Admin routes for the local pilot: team roster and profile editor.

Attached to the app via register(app) from qc/web.py. Saving a profile is
role-gated to leads/admins through the pilot identity cookie; this is
attribution-grade control (real auth is Entra ID in the production tier).
"""

import io
import json
import re
from datetime import date

from fastapi import File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse

from .ui import esc
from .ui_admin import (FONT_ROLES, render_admin_error, render_profile_edit,
                       render_profiles, render_team)

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


def _profile_rows() -> list[dict]:
    from .profile import PROFILES_DIR

    rows = []
    for path in sorted(PROFILES_DIR.glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        cfg = data.get("config", {})
        rows.append({
            "id": path.stem,
            "name": data.get("name", path.stem),
            "version": data.get("version", 1),
            "owner": data.get("owner", ""),
            "n_colors": len(cfg.get("color_palette", {}).get("named_colors", [])),
            "n_roles": len(cfg.get("font", {}).get("roles", {})),
        })
    return rows


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


def _apply_form_fields(cfg: dict, form) -> str | None:
    """Field-level updates onto cfg in place. Returns an error message or
    None. Only touches the keys the form exposes."""
    font = cfg.setdefault("font", {})
    roles = font.setdefault("roles", {})
    for role in FONT_ROLES:
        rc = roles.setdefault(role, {})
        latin = [f.strip() for f in form.get(f"{role}_latin", "").split(",") if f.strip()]
        cs = [f.strip() for f in form.get(f"{role}_complex", "").split(",") if f.strip()]
        rc["latin"] = latin
        rc["complex_script"] = cs
        size_raw = form.get(f"{role}_size", "").strip()
        if size_raw:
            try:
                rc["size_pt"] = float(size_raw)
            except ValueError:
                return f"{role} size must be a number"
        else:
            rc.pop("size_pt", None)

    named = []
    for line in form.get("palette", "").splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.rsplit(None, 1)
        if len(parts) != 2:
            return f"palette line needs 'name hex': {line[:40]}"
        name, hexval = parts[0].strip(), parts[1].strip().lstrip("#")
        if not _HEX.match(hexval):
            return f"'{hexval}' is not a 6-digit hex color"
        named.append({"name": name, "hex": hexval.upper(), "theme_ref": None,
                      "allowed_tints": [], "allowed_shades": []})
    cfg.setdefault("color_palette", {})["named_colors"] = named

    geometry = cfg.setdefault("geometry", {})
    margins = geometry.setdefault("safe_zone_margins_emu", {})
    numbers = (
        ("sz_left", margins, "left"), ("sz_right", margins, "right"),
        ("sz_top", margins, "top"), ("sz_bottom", margins, "bottom"),
    )
    for field, target, key in numbers:
        raw = form.get(field, "").strip()
        if raw:
            try:
                target[key] = int(float(raw))
            except ValueError:
                return f"{field} must be a number (EMU)"

    align = geometry.setdefault("alignment", {})
    shape = cfg.setdefault("shape_size", {})
    floats = (
        ("tol_font_size", font, "size_tolerance_pt", float),
        ("tol_edge", align, "edge_tolerance_emu", int),
        ("tol_spacing", align, "spacing_tolerance_emu", int),
        ("tol_shape_size", shape, "size_tolerance_emu", int),
        ("tol_near_miss", shape, "near_miss_ratio", float),
        ("tol_min_cohort", shape, "min_cohort_size", int),
    )
    for field, target, key, cast in floats:
        raw = form.get(field, "").strip()
        if raw:
            try:
                target[key] = cast(float(raw))
            except ValueError:
                return f"{field} must be a number"

    template = cfg.setdefault("header_footer", {}).setdefault("template", {})
    footer_text = form.get("footer_text", "").strip()
    template["footer_text"] = footer_text or None
    template["slide_number"] = "slide_number" in form
    return None


def register(app) -> None:
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

    @app.get("/profiles", response_class=HTMLResponse)
    def profiles():
        return render_profiles(_profile_rows())

    @app.post("/profiles/new", response_class=HTMLResponse)
    def profile_new(request: Request, deck: UploadFile = File(...),
                    name: str = Form(...)):
        from pptx import Presentation

        from .bootstrap import build_profile
        from .profile import PROFILES_DIR
        from .web import MAX_UPLOAD_BYTES, _zip_bomb_reason  # reuse audit guards

        def _fail(msg, code=400):
            from .store import list_users  # noqa: F401 (parity with list page)

            rows = _profile_rows()
            return HTMLResponse(render_profiles(rows, msg), status_code=code)

        editor = _editor(request)
        if editor is None:
            return HTMLResponse(render_admin_error(
                "Profiles", "Creating a profile needs a lead or admin",
                "Sign in as a lead or admin to create profiles.",
                back="/profiles"), status_code=403)

        filename = deck.filename or "reference.pptx"
        if not filename.lower().endswith(".pptx"):
            return _fail("Reference file must be a .pptx.")
        slug = _slugify(name)
        if not slug:
            return _fail("Give the profile a name with letters or numbers.")

        data = deck.file.read(MAX_UPLOAD_BYTES + 1)
        if len(data) > MAX_UPLOAD_BYTES:
            return _fail("Reference file exceeds the size cap.", 413)
        bomb = _zip_bomb_reason(data)
        if bomb:
            return _fail(f"Reference file rejected: {bomb}.", 413)

        try:
            prs = Presentation(io.BytesIO(data))
            pid = _unique_pid(slug)
            profile = build_profile(prs, pid, name.strip())
        except Exception as exc:
            return _fail(f"Could not read that reference deck: "
                         f"{type(exc).__name__}: {exc}", 422)

        profile["owner"] = (f"{editor['name']} (from reference {filename}, "
                            f"{date.today().strftime('%d/%m/%Y')})")
        (PROFILES_DIR / f"{pid}.json").write_text(
            json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")
        # land in the editor to review palette names and tweak before use
        return RedirectResponse(f"/profiles/{pid}/edit", status_code=303)

    @app.post("/profiles/{pid}/master", response_class=HTMLResponse)
    def profile_master_replace(pid: str, request: Request,
                               master: UploadFile = File(...)):
        """Replace the master file a profile applies, and re-read its frame.

        A profile's master is stored once, when the profile is created, and the
        format step hands PowerPoint that stored copy. Everything the designer
        adds to their master afterwards - a presentation-space rectangle, a
        moved guide, a new layout - reached no deck at all until this existed
        (design lead, 21/08/2026).

        Only what the master STATES is re-read. Fonts, palette and tolerances
        stay exactly as the lead edited them: those are decisions about the
        client, not readings of the file, and silently reverting them to a fresh
        projection would punish every edit ever made here."""
        from pptx import Presentation

        from .stylespec import (dominant_master, extract_layouts,
                                extract_style_spec, spec_to_profile)
        from .templates import save_master

        data = _load_profile(pid)
        if data is None:
            return HTMLResponse(render_admin_error(
                "Profile editor", "Not found",
                f"No profile named '{pid}'."), status_code=404)
        editor = _editor(request)
        if editor is None:
            return HTMLResponse(render_admin_error(
                "Profile editor", "Replacing the master needs a lead or admin",
                "Pick your name in the top right: the master is what every "
                "formatted deck is rebuilt on, so the change needs an owner.",
                back=f"/profiles/{pid}/edit"), status_code=403)

        filename = master.filename or ""
        if not filename.lower().endswith(".pptx"):
            return HTMLResponse(render_profile_edit(
                pid, data, master_note="Only .pptx files are accepted."),
                status_code=400)
        blob = master.file.read()
        try:
            prs = Presentation(io.BytesIO(blob))
            target = dominant_master(prs)
            if target is None:
                raise ValueError("the file carries no slide master")
            spec = extract_style_spec(prs, source=filename, embed_assets=False)
            fresh = spec_to_profile(spec, pid, data.get("name", pid))["config"]
        except Exception as exc:
            return HTMLResponse(render_profile_edit(
                pid, data,
                master_note=f"Could not read that master: "
                            f"{esc(type(exc).__name__)}: {esc(str(exc))}"),
                status_code=422)

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
                else "reserved header band dropped: the master no longer "
                     "states one")

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
            changes.append(
                f"layouts ({len(added)} added, {len(dropped)} removed)")

        source = (fresh.get("style_spec_source") or {})
        was = (cfg.get("style_spec_source") or {}).get("grid_source")
        cfg["style_spec_source"] = source
        if source.get("grid_source") != was:
            changes.append(f"frame now read from "
                           f"{str(source.get('grid_source')).replace('_', ' ')}")

        save_master(pid, blob)
        data["version"] = int(data.get("version", 1)) + 1
        data["owner"] = (f"{editor['name']} (master replaced from {filename}, "
                         f"{date.today().strftime('%d/%m/%Y')})")
        _profile_path(pid).write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

        note = (f"Master replaced from <b>{esc(filename)}</b>. "
                + ("Re-read: " + "; ".join(changes) + "."
                   if changes else
                   "The rules it states are unchanged, so only the file was "
                   "replaced - which is still what formatted decks will now be "
                   "built on."))
        return HTMLResponse(render_profile_edit(pid, _load_profile(pid),
                                               master_note=note))

    @app.get("/profiles/{pid}/edit", response_class=HTMLResponse)
    def profile_edit(pid: str):
        data = _load_profile(pid)
        if data is None:
            return HTMLResponse(render_admin_error(
                "Profile editor", "Not found",
                f"No profile named '{pid}'."), status_code=404)
        return render_profile_edit(pid, data)

    @app.post("/profiles/{pid}/edit", response_class=HTMLResponse)
    async def profile_save(pid: str, request: Request):
        data = _load_profile(pid)
        if data is None:
            return HTMLResponse(render_admin_error(
                "Profile editor", "Not found",
                f"No profile named '{pid}'."), status_code=404)
        editor = _editor(request)
        if editor is None:
            return HTMLResponse(render_admin_error(
                "Profile editor", "Saving needs a lead or admin",
                "Pick your name in the top right; profile changes are "
                "restricted to leads and admins so the rule set has an owner.",
                back=f"/profiles/{pid}/edit"), status_code=403)

        form = dict(await request.form())
        cfg = data.get("config", {})
        rendered_raw = json.dumps(cfg, indent=2, ensure_ascii=False)
        submitted_raw = form.get("raw", "")

        if submitted_raw.strip() and submitted_raw.strip() != rendered_raw.strip():
            # raw JSON wins wholesale when edited
            try:
                cfg = json.loads(submitted_raw)
            except json.JSONDecodeError as exc:
                return HTMLResponse(render_profile_edit(
                    pid, data, error=f"Raw JSON does not parse: {exc}"),
                    status_code=400)
            if not isinstance(cfg, dict):
                return HTMLResponse(render_profile_edit(
                    pid, data, error="Raw JSON must be an object"),
                    status_code=400)
        else:
            error = _apply_form_fields(cfg, form)
            if error:
                return HTMLResponse(render_profile_edit(pid, data, error=error),
                                    status_code=400)

        data["config"] = cfg
        data["name"] = form.get("name", data.get("name", pid)).strip() or pid
        data["version"] = int(data.get("version", 1)) + 1
        data["owner"] = f"{editor['name']} (edited {date.today().strftime('%d/%m/%Y')})"
        _profile_path(pid).write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        return RedirectResponse("/profiles", status_code=303)
