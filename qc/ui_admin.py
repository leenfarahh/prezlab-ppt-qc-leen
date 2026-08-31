"""The team roster, and the one error page the roster can raise.

Rendering only; route logic lives in qc/web_admin.py. Reuses the Prezlab brand
shell from qc/ui.py so it reads as one product with the two pages that matter.

The profile list and the profile editor used to live here. They are gone:
profiles are created and revised in step 1 of Prepare a deck, from the master
they describe, which is the only place a designer has the file open anyway.
"""

from .ui import _shell, esc

_INPUT = ("border:1px solid var(--line);border-radius:10px;"
          "padding:0.42rem 0.75rem;background:#fff;color:var(--teal);"
          "font-size:0.88rem;")


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
<p><a href="/">Run an audit</a> &middot; <a href="/prep">Prepare a deck</a></p>"""
    return _shell("Team", body)


def render_admin_error(kicker: str, heading: str, message: str,
                       back: str = "/profiles") -> str:
    body = f"""
<span class="kicker">{esc(kicker)}</span>
<h1>{esc(heading)}</h1>
{_warn(message)}
<p><a class="btn ghost" href="{esc(back)}">Back</a></p>"""
    return _shell(heading, body)
