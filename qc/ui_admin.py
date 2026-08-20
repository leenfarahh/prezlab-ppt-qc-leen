"""Admin pages for the local pilot: team roster and profile editor.

Rendering only; route logic lives in qc/web_admin.py. Reuses the Prezlab
brand shell from qc/ui.py so admin pages read as one product.
"""

import json

from .ui import MODULE_LABELS, _shell, esc

FONT_ROLES = ("title", "subtitle", "body", "caption")

_INPUT = ("border:1px solid var(--line);border-radius:10px;"
          "padding:0.42rem 0.75rem;background:#fff;color:var(--teal);"
          "font-size:0.88rem;")
_MONO_AREA = ("width:100%;border:1px solid var(--line);border-radius:10px;"
              "padding:0.6rem 0.75rem;background:#fff;color:var(--teal);"
              "font-family:'Cascadia Mono',Consolas,monospace;"
              "font-size:0.8rem;line-height:1.45;")


def _warn(message: str) -> str:
    return f'<div class="banner warn">{esc(message)}</div>' if message else ""


def _fmt(value) -> str:
    """Numbers render without a spurious trailing .0; None renders empty."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def render_team(users: list[dict], roles: tuple, message: str = "") -> str:
    rows = "".join(
        f"<tr><td>{esc(u['name'])}</td><td>{esc(u['role'])}</td>"
        f"<td>{esc(str(u.get('created_at', ''))[:10])}</td>"
        f"<td><form method='post' action='/team/reset-pin' style='margin:0'>"
        f"<input type='hidden' name='name' value=\"{esc(u['name'])}\">"
        f"<button class='btn ghost' type='submit' "
        f"title='Clears the PIN and signs them out everywhere; next sign-in "
        f"sets a new one (lead/admin only)'>Reset PIN</button></form></td></tr>"
        for u in users) or ('<tr><td colspan="4" class="note">No one yet. '
                            'Add the first person below.</td></tr>')
    options = "".join(f'<option value="{esc(r)}">{esc(r)}</option>' for r in roles)
    body = f"""
<span class="kicker">Pilot team</span>
<h1>Team.</h1>
<p class="sub">Names here are attribution for triage, comments, and profile
 edits, not accounts. Leads and admins can save profile changes; everyone
 else reads and comments.</p>
{_warn(message)}
<table class="stats"><thead><tr><th>Name</th><th>Role</th><th>Added</th><th></th></tr></thead>
<tbody>{rows}</tbody></table>
<div class="card">
 <form method="post" action="/team/add">
  <fieldset style="border:0;padding:0;margin:0"><legend>Add someone</legend>
   <input type="text" name="name" placeholder="Full name" required
    aria-label="Name" style="{_INPUT}width:16rem">
   <select name="role" aria-label="Role" style="{_INPUT}">{options}</select>
   <button class="btn primary" type="submit">Add to team</button>
  </fieldset>
  <p class="note">Adding a name that already exists keeps the existing entry
   and role.</p>
 </form>
</div>
<p><a href="/">Back to audits</a> &middot; <a href="/profiles">Profiles</a></p>"""
    return _shell("Team", body)


def render_profiles(rows: list[dict], message: str = "") -> str:
    trs = "".join(
        f"""<tr><td class="code">{esc(r['id'])}</td><td>{esc(r['name'])}</td>
<td>v{esc(r['version'])}</td><td>{esc(r['owner'])}</td>
<td>{esc(r['n_colors'])}</td><td>{esc(r['n_roles'])}</td>
<td><a href="/profiles/{esc(r['id'])}/edit">Edit</a></td></tr>"""
        for r in rows) or ('<tr><td colspan="7" class="note">No profiles '
                           'found.</td></tr>')
    body = f"""
<span class="kicker">Formatting profiles</span>
<h1>Profiles.</h1>
<p class="sub">Every audit runs against one of these rule sets. Anyone can
 open a profile to read it; creating or saving one takes a lead or admin.</p>
{_warn(message)}
<table class="stats"><thead><tr><th>Profile</th><th>Name</th><th>Version</th>
<th>Owner</th><th>Palette colors</th><th>Font roles</th><th></th></tr></thead>
<tbody>{trs}</tbody></table>

<div class="card">
 <form method="post" action="/profiles/new" enctype="multipart/form-data"
   onsubmit="if(window.showBusy)showBusy('Reading the reference deck',
    'Learning its fonts, palette, and layout to build a profile.')">
  <b>New profile from a reference deck</b>
  <p class="note">Upload a client's reference file. The tool reads the fonts,
   colors, and layout it consistently uses and turns them into a starting
   profile, then drops you in the editor to name the colors and adjust.</p>
  <p><input type="file" name="deck" accept=".pptx" required></p>
  <p><input type="text" name="name" placeholder="Profile name, e.g. Client ABC"
    required style="width:22rem"></p>
  <button class="btn primary" type="submit">Create from reference</button>
  <p class="note">Needs a lead or admin. The generated profile is a starting
   point to review, not a final brand definition.</p>
 </form>
</div>
<p><a href="/">Back to audits</a> &middot; <a href="/team">Team</a></p>"""
    return _shell("Profiles", body)


def render_admin_error(kicker: str, heading: str, message: str,
                       back: str = "/profiles") -> str:
    body = f"""
<span class="kicker">{esc(kicker)}</span>
<h1>{esc(heading)}</h1>
{_warn(message)}
<p><a class="btn ghost" href="{esc(back)}">Back</a></p>"""
    return _shell(heading, body)


def _num_input(name: str, value, width: str = "7rem") -> str:
    return (f'<input type="number" step="any" name="{esc(name)}" '
            f'value="{esc(_fmt(value))}" style="{_INPUT}width:{width}">')


def _master_card(pid: str, data: dict, note: str = "") -> str:
    """What master file this profile carries, and how to replace it.

    A profile's master is stored once, when the profile is created, and applying
    the profile hands PowerPoint THAT file. So a rectangle, guide or layout the
    designer adds to their master afterwards reaches no deck until the stored
    copy is replaced, which is what this is for (design lead, 21/08/2026: the
    presentation-space box was missing from every formatted deck because the
    stored master predated it)."""
    from .templates import master_info

    info = master_info(pid)
    src = (data.get("config", {}).get("style_spec_source") or {})
    if info:
        from datetime import datetime

        when = datetime.fromtimestamp(info["modified"]).strftime("%d/%m/%Y %H:%M")
        state = (f"<b>{info['bytes'] / 1e6:.1f} MB</b>, stored {esc(when)} "
                 f"<span class='note'>({esc(info['sha1'][:12])})</span>")
        frame = src.get("grid_source")
        state += ("<br><span class='note'>Frame read from "
                  f"<b>{esc(str(frame).replace('_', ' '))}</b>."
                  " Re-read the master after adding a presentation-space "
                  "rectangle, moving a guide, or changing a layout, or decks "
                  "keep getting formatted on the old one.</span>"
                  if frame else "")
    else:
        state = ("<span class='note'>None. This profile can be audited "
                 "against but not applied: applying a master needs the file "
                 "itself.</span>")

    return f"""
<div class="card">
 <div class="tag">Master file</div>
 <h3 style="margin:0 0 0.3rem">The file this profile applies</h3>
 <p class="sub">{state}</p>
 {f'<div class="banner warn">{note}</div>' if note else ""}
 <form method="post" action="/profiles/{esc(pid)}/master"
   enctype="multipart/form-data">
  <input type="file" name="master" accept=".pptx" required>
  <button class="btn ghost" type="submit">Replace the master and re-read its
   frame</button>
  <p class="note">Replaces the stored file and re-reads what the master
   STATES: margins (presentation space, else guides, else placeholders), the
   reserved header band, the column grid and the layout allow-list. Fonts,
   palette and tolerances are left exactly as edited here, since those are
   yours. Needs a lead or admin.</p>
 </form>
</div>"""


def render_profile_edit(pid: str, data: dict, error: str = "",
                        master_note: str = "") -> str:
    cfg = data.get("config", {})
    roles_cfg = cfg.get("font", {}).get("roles", {})
    role_rows = []
    for role in FONT_ROLES:
        rc = roles_cfg.get(role, {})
        role_rows.append(f"""
 <div style="margin:0.55rem 0 0.9rem">
  <b style="text-transform:capitalize">{esc(role)}</b><br>
  <label class="note">Latin families
   <input type="text" name="{role}_latin"
    value="{esc(', '.join(rc.get('latin', [])))}" style="{_INPUT}width:12rem"></label>
  <label class="note">Complex script
   <input type="text" name="{role}_complex"
    value="{esc(', '.join(rc.get('complex_script', [])))}" style="{_INPUT}width:10rem"></label>
  <label class="note">Size pt {_num_input(f"{role}_size", rc.get('size_pt'), '5.5rem')}</label>
 </div>""")

    palette_lines = "\n".join(
        f"{c.get('name', '')} {c.get('hex', '')}"
        for c in cfg.get("color_palette", {}).get("named_colors", []))
    margins = cfg.get("geometry", {}).get("safe_zone_margins_emu", {})
    align = cfg.get("geometry", {}).get("alignment", {})
    shape = cfg.get("shape_size", {})
    footer = cfg.get("header_footer", {}).get("template", {})
    raw = json.dumps(cfg, indent=2, ensure_ascii=False)
    slide_no = " checked" if footer.get("slide_number") else ""

    body = f"""
<span class="kicker">Profile editor</span>
<h1 class="file">{esc(pid)}</h1>
<p class="sub">Version <b>v{esc(data.get('version', 1))}</b> &middot;
 owner <b>{esc(data.get('owner', ''))}</b>. Saving bumps the version so
 audit reports always name the exact rule set they ran against.</p>
{_warn(error)}
<form method="post" action="/profiles/{esc(pid)}/edit">
 <div class="card">
  <label>Profile name<br>
   <input type="text" name="name" value="{esc(data.get('name', pid))}"
    style="{_INPUT}width:26rem"></label>
 </div>
 <div class="config">
  <fieldset><legend>{esc(MODULE_LABELS['font'])}</legend>
   {''.join(role_rows)}
   <p class="note">Families are comma-separated; the first is the target.
    Leave a size empty to drop the size target for that role.</p>
   <label class="note">Size tolerance pt
    {_num_input('tol_font_size', cfg.get('font', {}).get('size_tolerance_pt'))}</label>
  </fieldset>
  <fieldset><legend>{esc(MODULE_LABELS['color_palette'])}</legend>
   <textarea name="palette" rows="8" spellcheck="false"
    style="{_MONO_AREA}">{esc(palette_lines)}</textarea>
   <p class="note">One color per line: <b>name hex</b>, e.g.
    <span style="font-family:Consolas,monospace">prezlab_navy 1F4E79</span>.
    Theme references and tint/shade allowances keep their defaults.</p>
  </fieldset>
  <fieldset><legend>{esc(MODULE_LABELS['margin_alignment'])}</legend>
   <label class="note">Left {_num_input('sz_left', margins.get('left'))}</label>
   <label class="note">Right {_num_input('sz_right', margins.get('right'))}</label><br>
   <label class="note">Top {_num_input('sz_top', margins.get('top'))}</label>
   <label class="note">Bottom {_num_input('sz_bottom', margins.get('bottom'))}</label>
   <p class="note">Safe zone margins in EMU (914400 per inch).</p>
   <label class="note">Edge tolerance EMU
    {_num_input('tol_edge', align.get('edge_tolerance_emu'))}</label>
   <label class="note">Spacing tolerance EMU
    {_num_input('tol_spacing', align.get('spacing_tolerance_emu'))}</label>
  </fieldset>
  <fieldset><legend>{esc(MODULE_LABELS['shape_size'])}</legend>
   <label class="note">Size tolerance EMU
    {_num_input('tol_shape_size', shape.get('size_tolerance_emu'))}</label>
   <label class="note">Near-miss ratio
    {_num_input('tol_near_miss', shape.get('near_miss_ratio', 0.08))}</label>
   <label class="note">Min cohort size
    {_num_input('tol_min_cohort', shape.get('min_cohort_size'))}</label>
  </fieldset>
  <fieldset><legend>{esc(MODULE_LABELS['header_footer'])}</legend>
   <label class="note">Footer text
    <input type="text" name="footer_text"
     value="{esc(footer.get('footer_text') or '')}"
     style="{_INPUT}width:16rem"></label>
   <p class="note">Empty means no footer text is enforced.</p>
   <label class="note"><input type="checkbox" name="slide_number"{slide_no}>
    Require slide numbers</label>
  </fieldset>
  <fieldset style="flex-basis:100%"><legend>Raw JSON</legend>
   <textarea name="raw" rows="16" spellcheck="false"
    style="{_MONO_AREA}">{esc(raw)}</textarea>
   <p class="note">The full config block. If you edit it, it replaces the
    fields above wholesale; leave it untouched to save the fields instead.</p>
  </fieldset>
 </div>
 <div class="actions">
  <button class="btn primary" type="submit">Save profile</button>
  <a class="btn ghost" href="/profiles">Back to profiles</a>
  <p class="note">Saving needs a lead or admin role: pick your name in the
   top right. Viewing is open to everyone.</p>
 </div>
</form>
{_master_card(pid, data, master_note)}"""
    return _shell(f"Edit profile: {pid}", body)
