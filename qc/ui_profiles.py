"""The profiles: what the tool holds a deck to, and how to change it.

Two pages. A list of every profile, and a form for one of them.

WHY A FORM AND NOT A RE-READ. A profile could only be produced by reading a
master, and changing one number meant reading the master again - which rebuilds
the whole document and discards every decision anyone made against it since. The
common edits are one field wide: the body size is 17, this navy belongs in the
palette, this client's margins are tighter than ours.

WHAT IT IS NOT is a JSON editor. Every input is typed, bounded and labelled in
the unit a designer thinks in - inches for margins, points for type and
tolerances - and the conversion to EMU happens in qc.profileform, which is also
what renders the fields, so the page and the parser cannot disagree about what
exists. A value the form refuses comes back with the rest of the edit intact.

Rendering only. The spec is qc.profileform and the routes are qc/web_admin.py.
"""

from .profileform import display, groups, multiple, summary
from .ui import _shell, esc

_CSS = """
<style>
.plist { display: grid; gap: 0.7rem; margin: 1.2rem 0; }
.prow { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 0.9rem 1.1rem; display: flex; gap: 1rem; align-items: center;
  flex-wrap: wrap; }
.prow .who { flex: 1 1 18rem; }
.prow .nm { font-weight: 700; font-size: 1.05rem; }
.prow .meta { color: var(--slate-text); font-size: 0.85rem; margin-top: 0.15rem; }
.prow .acts { display: flex; gap: 0.5rem; flex-wrap: wrap; }
.badge { display: inline-block; font-size: 10px; font-weight: 800;
  letter-spacing: 0.08em; text-transform: uppercase; border-radius: 999px;
  padding: 0.12rem 0.5rem; margin-left: 0.4rem; }
.badge.default { background: var(--teal-light); color: var(--teal); }
.badge.nomaster { background: var(--sand); color: var(--burgundy); }

.grp { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 1.2rem 1.4rem; margin: 0 0 1rem; }
.grp > h2 { font-size: 1.1rem; margin: 0 0 0.2rem; }
.grp > .blurb { color: var(--slate-text); font-size: 0.9rem; margin: 0 0 1rem; }
.fields { display: grid; gap: 0.9rem;
  grid-template-columns: repeat(auto-fill, minmax(15rem, 1fr)); }
.f { display: flex; flex-direction: column; gap: 0.25rem; min-width: 0; }
.f > label { font-weight: 700; font-size: 0.88rem; }
.f .hint { color: var(--slate-text); font-size: 0.8rem; }
.f input[type=text], .f input[type=number], .f select {
  font: inherit; padding: 0.45rem 0.6rem; border-radius: 8px;
  border: 1px solid var(--line); background: #fff; color: var(--teal);
  width: 100%; min-width: 0; }
.f.bad input, .f.bad select { border-color: var(--orange); border-width: 2px; }
.f .err { color: var(--burgundy); font-size: 0.8rem; font-weight: 600; }
.f.wide { grid-column: 1 / -1; }
.f.check { flex-direction: row; align-items: center; gap: 0.5rem; }
.f.check > label { font-weight: 600; }
.weights { display: flex; gap: 0.6rem; flex-wrap: wrap; }
.weights label { display: inline-flex; align-items: center; gap: 0.3rem;
  font-weight: 600; font-size: 0.88rem; }

.palette { display: grid; gap: 0.5rem; }
.crow { display: flex; gap: 0.6rem; align-items: center; }
.crow input[type=text] { font: inherit; padding: 0.45rem 0.6rem;
  border-radius: 8px; border: 1px solid var(--line); background: #fff;
  color: var(--teal); }
.crow .cname { flex: 1 1 12rem; min-width: 0; }
.crow .chex { flex: 0 0 8rem; text-transform: uppercase; }
.sw { width: 2rem; height: 2rem; border-radius: 6px; border: 1px solid var(--line);
  flex: 0 0 auto; }
.crow .drop { border: 0; background: none; color: var(--slate-text);
  cursor: pointer; font: inherit; padding: 0.2rem 0.4rem; }
.crow .drop:hover { color: var(--burgundy); }

.savebar { position: sticky; bottom: 0; background: #fff;
  border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 0.7rem 1.1rem; margin-top: 1rem; display: flex; gap: 0.9rem;
  align-items: center; flex-wrap: wrap;
  box-shadow: 0 -2px 12px rgba(0, 37, 40, 0.06); }
.savebar .grow { flex: 1; }
.danger { color: var(--burgundy); }
</style>
"""


def _err(errors: dict, name: str) -> str:
    msg = (errors or {}).get(name)
    return f'<div class="err">{esc(msg)}</div>' if msg else ""


def _input(f, value, errors: dict) -> str:
    bad = " bad" if (errors or {}).get(f.name) else ""
    hint = f'<div class="hint">{esc(f.help)}</div>' if f.help else ""

    if f.kind == "bool":
        checked = " checked" if value else ""
        return f"""
<div class="f check{bad}">
 <input type="checkbox" id="{esc(f.name)}" name="{esc(f.name)}" value="1"{checked}>
 <label for="{esc(f.name)}">{esc(f.label)}</label>{hint}{_err(errors, f.name)}
</div>"""

    if f.kind == "select" and multiple(f):
        picked = set(value or [])
        boxes = "".join(
            f'<label><input type="checkbox" name="{esc(f.name)}" '
            f'value="{esc(opt)}"{" checked" if opt in picked else ""}>'
            f'{esc(opt)}</label>' for opt in f.options)
        return f"""
<div class="f wide{bad}"><label>{esc(f.label)}</label>
 <div class="weights">{boxes}</div>{hint}{_err(errors, f.name)}</div>"""

    if f.kind == "select":
        opts = "".join(
            f'<option value="{esc(o)}"'
            f'{" selected" if str(value) == o else ""}>{esc(o)}</option>'
            for o in f.options)
        return f"""
<div class="f{bad}"><label for="{esc(f.name)}">{esc(f.label)}</label>
 <select id="{esc(f.name)}" name="{esc(f.name)}">{opts}</select>
 {hint}{_err(errors, f.name)}</div>"""

    if f.kind in ("int", "number", "inches", "points"):
        step = "1" if f.kind == "int" else "any"
        bounds = ""
        if f.minimum is not None:
            bounds += f' min="{f.minimum}"'
        if f.maximum is not None:
            bounds += f' max="{f.maximum}"'
        return f"""
<div class="f{bad}"><label for="{esc(f.name)}">{esc(f.label)}</label>
 <input type="number" step="{step}"{bounds} id="{esc(f.name)}"
  name="{esc(f.name)}" value="{esc(str(value))}">{hint}{_err(errors, f.name)}</div>"""

    wide = " wide" if f.help and len(f.help) > 70 else ""
    ph = f' placeholder="{esc(f.placeholder)}"' if f.placeholder else ""
    return f"""
<div class="f{wide}{bad}"><label for="{esc(f.name)}">{esc(f.label)}</label>
 <input type="text" id="{esc(f.name)}" name="{esc(f.name)}"
  value="{esc(str(value))}"{ph}>{hint}{_err(errors, f.name)}</div>"""


def _palette(colors: list, errors: dict, can_edit: bool = True) -> str:
    rows = "".join(_color_row(c.get("name", ""), c.get("hex", ""), can_edit)
                   for c in (colors or []))
    err = _err(errors, "color_name") + _err(errors, "color_hex")
    head = f"""
<div class="grp">
 <h2>Palette</h2>
 <p class="blurb">The colours this client's work may use. A shape off the
  palette is measured against these and reported by how far off it is, so a
  name here is what a fix snaps a near-miss to.</p>
 <div class="palette" id="palette">{rows}</div>
 {err}"""
    # A reader gets the swatches and the values and no controls. A disabled
    # fieldset greys the inputs but leaves an "Add a colour" button that still
    # adds rows, which is a page offering an action it will not honour.
    if not can_edit:
        return head + "</div>"
    return head + f"""
 <p style="margin-top:0.8rem">
  <button type="button" class="btn ghost" id="addcolor">Add a colour</button></p>
</div>
<template id="crowtpl">{_color_row("", "")}</template>
<script>
document.getElementById('addcolor').addEventListener('click', () => {{
  const tpl = document.getElementById('crowtpl');
  document.getElementById('palette').appendChild(
    tpl.content.cloneNode(true));
}});
// One listener on the container rather than one per row, so a row added after
// load behaves like a row that was rendered with the page.
document.getElementById('palette').addEventListener('click', e => {{
  const drop = e.target.closest('.drop');
  if (drop) drop.closest('.crow').remove();
}});
document.getElementById('palette').addEventListener('input', e => {{
  if (!e.target.classList.contains('chex')) return;
  const row = e.target.closest('.crow');
  const hex = e.target.value.trim().replace(/^#/, '');
  row.querySelector('.sw').style.background =
    /^[0-9a-fA-F]{{6}}$/.test(hex) ? '#' + hex : 'transparent';
}});
</script>"""


def _color_row(name: str, hexval: str, can_edit: bool = True) -> str:
    swatch = f"#{esc(hexval)}" if hexval else "transparent"
    remove = ('<button type="button" class="drop" '
              'aria-label="Remove this colour">Remove</button>'
              if can_edit else "")
    return f"""<div class="crow">
 <span class="sw" style="background:{swatch}"></span>
 <input class="cname" type="text" name="color_name" value="{esc(name)}"
  placeholder="prezlab navy" aria-label="Colour name">
 <input class="chex" type="text" name="color_hex" value="{esc(hexval)}"
  placeholder="1F4E79" aria-label="Hex code" maxlength="7">
 {remove}
</div>"""


def render_profiles(profiles: list, *, can_edit: bool, message: str = "",
                    ok: str = "") -> str:
    """`profiles` are {id, name, summary, has_master, is_default, version}."""
    banner = ""
    if ok:
        banner += f'<div class="banner ok">{esc(ok)}</div>'
    if message:
        banner += f'<div class="banner warn">{esc(message)}</div>'

    rows = []
    for p in profiles:
        badges = ""
        if p.get("is_default"):
            badges += '<span class="badge default">Default</span>'
        if not p.get("has_master"):
            badges += '<span class="badge nomaster">No master</span>'
        acts = (f'<a class="btn ghost" href="/profiles/{esc(p["id"])}">Edit</a>'
                if can_edit else
                f'<a class="btn ghost" href="/profiles/{esc(p["id"])}">View</a>')
        rows.append(f"""
<div class="prow">
 <div class="who"><div class="nm">{esc(p['name'])}{badges}</div>
  <div class="meta">{p['summary']} &middot; v{esc(str(p.get('version', 1)))}
   &middot; <code>{esc(p['id'])}</code></div></div>
 <div class="acts">{acts}</div>
</div>""")

    note = ("" if can_edit else
            '<p class="note">Sign in as a lead or admin to change a profile. '
            'Anyone signed in can read one.</p>')
    body = f"""
<h1>Profiles.</h1>
<p class="sub">What a deck gets held to: the type, the palette, the margins and
 the tolerances. A profile is created by reading a master on
 <a href="/prep"><b>Prepare a deck</b></a>; this is where you change one
 afterwards without re-reading it.</p>
{banner}{note}
<div class="plist">{''.join(rows) or '<p class="note">No profiles yet.</p>'}</div>
"""
    return _shell("Profiles", _CSS + body)


def render_profile(profile: dict, *, pid: str, can_edit: bool,
                   has_master: bool, master_note: str = "",
                   errors: dict | None = None, message: str = "",
                   ok: str = "") -> str:
    errors = errors or {}
    banner = ""
    if ok:
        banner += f'<div class="banner ok">{esc(ok)}</div>'
    if message:
        banner += f'<div class="banner warn">{esc(message)}</div>'

    blocks = []
    for group in groups():
        fields = "".join(_input(f, display(profile, f), errors)
                         for f in group.fields)
        blocks.append(f"""
<div class="grp"><h2>{esc(group.title)}</h2>
 <p class="blurb">{esc(group.blurb)}</p>
 <div class="fields">{fields}</div></div>""")

    colors = ((profile.get("config") or {}).get("color_palette") or {}) \
        .get("named_colors") or []
    # After Identity, before Type: the palette is the thing people come here to
    # change, and putting it under seven collapsed sections is how a page grows
    # a reputation for not having the field you need.
    blocks.insert(1, _palette(colors, errors, can_edit))

    if not can_edit:
        return _shell(f"Profile: {profile.get('name', pid)}", _CSS + f"""
<h1 class="file">{esc(profile.get('name', pid))}</h1>
<p class="sub">Read-only. Sign in as a lead or admin to change it.</p>
{banner}<fieldset disabled style="border:0;padding:0;margin:0">
{''.join(blocks)}</fieldset>""")

    master_block = f"""
<div class="grp"><h2>The stored master</h2>
 <p class="blurb">{esc(master_note)}</p>
 <form method="post" action="/profiles/{esc(pid)}/master"
  enctype="multipart/form-data"
  data-busy="Reading the revised master"
  data-busysub="Its frame, reserved bands, grid and layout names are re-read.
   The type, palette and tolerances on this page are left exactly as they are.">
  <input type="file" name="master" accept=".pptx" required>
  <button class="btn ghost" type="submit">Replace the master</button>
 </form>
</div>"""

    body = f"""
<h1 class="file">{esc(profile.get('name', pid))}</h1>
<p class="sub">Editing <code>{esc(pid)}</code>, version
 {esc(str(profile.get('version', 1)))}. Saving bumps the version; decks already
 audited keep the version they were audited against.</p>
{banner}
<form method="post" action="/profiles/{esc(pid)}" id="pf" data-busy="Saving">
{''.join(blocks)}
{master_block if has_master else ''}
 <div class="savebar">
  <button class="btn primary" type="submit">Save profile</button>
  <span class="grow"></span>
  <a class="btn ghost" href="/profiles">Back to profiles</a>
 </div>
</form>
<form method="post" action="/profiles/{esc(pid)}/duplicate"
 style="display:inline-block;margin-top:1rem">
 <button class="btn ghost" type="submit">Duplicate this profile</button>
</form>
<form method="post" action="/profiles/{esc(pid)}/delete"
 style="display:inline-block;margin-top:1rem;margin-left:0.5rem"
 onsubmit="return confirm('Delete this profile and its stored master? Decks already audited against it keep their reports.')">
 <button class="btn ghost danger" type="submit">Delete</button>
</form>"""
    return _shell(f"Profile: {profile.get('name', pid)}", _CSS + body)


def profile_row(pid: str, profile: dict, has_master: bool) -> dict:
    """The list page's view of one profile. Here rather than in the route so
    the shape the template reads is defined beside the template."""
    return {"id": pid, "name": profile.get("name") or pid,
            "summary": summary(profile), "has_master": has_master,
            "is_default": bool(profile.get("is_default")),
            "version": profile.get("version", 1)}
