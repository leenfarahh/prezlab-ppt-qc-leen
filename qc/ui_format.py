"""The formatting page: apply a profile's master to every slide of a deck.

Separate from the audit page on purpose. Auditing reads a deck and reports;
this REWRITES it. Those deserve different doors, different wording, and a
result screen that says exactly what happened to each slide rather than a
single "done".

Rendering only; route logic lives in qc/web.py.
"""

from .ui import _shell, esc

_RULE_LABEL = {
    "name": ("matched by name", "The deck's layout and the master's share a "
                                "name, so the designer meant them to correspond."),
    "archetype": ("matched by archetype", "No name match, but both layouts "
                                          "declare the same OOXML archetype."),
    "fallback": ("fell back", "Neither the name nor the archetype matched. "
                              "Content may be orphaned; check these slides."),
    "none": ("no target", "The master defines no layout that could be used."),
}
_RULE_ORDER = ("name", "archetype", "fallback", "none")


def _warn(message: str) -> str:
    return f'<div class="banner warn">{esc(message)}</div>' if message else ""


def render_format_intake(profiles: list[dict], message: str = "",
                         com_ready: bool = True) -> str:
    """profiles: [{id, name, has_master, layouts}] - only ones carrying a
    master can be applied, and the rest are shown disabled WITH the reason,
    because a silently missing option reads as a bug."""
    usable = [p for p in profiles if p["has_master"]]
    unusable = [p for p in profiles if not p["has_master"]]

    # Each card names the master file it would apply and what frame that file
    # states. A profile whose stored master predates the designer's latest one
    # formats decks on the old file, and the only symptom is something missing
    # from the output (design lead, 21/08/2026: the presentation-space
    # rectangle). Saying it here is what turns that into a visible fact.
    def _frame_label(p) -> str:
        frame = (p.get("frame") or "").replace("_", " ")
        stored = p.get("master_stored")
        bits = [f"{p['layouts']} layouts"]
        if stored:
            bits.append(f"stored {stored}")
        if frame:
            bits.append(f"frame from {frame}")
        return " &middot; ".join(bits)

    cards = "".join(
        f"""<label class="radio-card"><input type="radio" name="profile"
        value="{esc(p['id'])}"{' checked' if i == 0 else ''}>
        <span><b>{esc(p['name'])}</b><small>{_frame_label(p)}</small></span></label>"""
        for i, p in enumerate(usable))

    if not usable:
        cards = ('<p class="note">No profile carries a master yet. Read a '
                 'master on the <a href="/master">Read a master</a> page and '
                 'save it as a profile; the master is stored with it and this '
                 'page can then apply it.</p>')

    others = ""
    if unusable:
        names = ", ".join(esc(p["name"]) for p in unusable)
        others = (f'<p class="note">Not listed, because they carry no master '
                  f'file to apply: {names}. A profile only gains one when it '
                  f'is created from a master.</p>')

    blocker = ""
    if not com_ready:
        blocker = _warn(
            "Desktop PowerPoint is not reachable on this machine, and applying "
            "a master needs it: PowerPoint's own placeholder matching is what "
            "moves each slide's content into the new layout. Run this on the "
            "Windows box.")

    disabled = " disabled" if (not usable or not com_ready) else ""

    body = f"""
<h1>Apply a master to a deck.</h1>
<p class="sub">Rebuilds every slide on the chosen master's layouts. Each slide is
copied, the master's layout is applied to the copy, and the original is deleted,
one slide at a time. Your upload is never modified; you get a new file back.</p>
{_warn(message)}{blocker}
<form action="/format" method="post" enctype="multipart/form-data" id="f">
  <div class="drop" id="drop" tabindex="0" role="button"
       aria-label="Drop the deck to format here or press Enter to browse">
    <strong>Drop the deck to format here</strong> or click to browse
    <div class="hint">Processed on this machine. The upload is deleted after
    processing and the rebuilt deck is held in memory for download.</div>
    <div class="file" id="fname" aria-live="polite"></div>
    <input type="file" name="deck" id="deck" accept=".pptx" required hidden>
  </div>
  <div class="config">
    <fieldset><legend>Master to apply</legend>{cards}{others}</fieldset>
  </div>
  <div class="actions">
    <button class="btn primary" id="go" type="submit"{disabled}>Apply the master</button>
  </div>
</form>
<script>
const drop = document.getElementById('drop'), input = document.getElementById('deck'),
      fname = document.getElementById('fname'), go = document.getElementById('go');
const blocked = {str(bool(disabled)).lower()};
function arm() {{
  if (input.files.length) {{ fname.textContent = input.files[0].name; }}
  go.disabled = blocked || !input.files.length;
}}
drop.addEventListener('click', () => input.click());
drop.addEventListener('keydown', e => {{ if (e.key === 'Enter' || e.key === ' ') input.click(); }});
input.addEventListener('change', arm);
document.getElementById('f').addEventListener('submit', () => {{
  go.disabled = true;
  showBusy('Applying the master to ' + (input.files[0] ? input.files[0].name : 'the deck'),
           'Rebuilding one slide at a time through PowerPoint. Expect roughly a second per slide.');
}});
['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => {{
  e.preventDefault(); drop.classList.add('armed'); }}));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => {{
  e.preventDefault(); drop.classList.remove('armed'); }}));
drop.addEventListener('drop', e => {{
  if (e.dataTransfer.files.length) {{ input.files = e.dataTransfer.files; arm(); }} }});
arm();
</script>"""
    return _shell("Apply a master", body)


def _slide_rows(plans: list, errors: dict) -> str:
    rows = []
    for p in plans:
        err = errors.get(p.slide_index)
        if err:
            state = f'<span class="pill error">failed</span>'
            detail = esc(err)
        else:
            label = _RULE_LABEL.get(p.match_rule, (p.match_rule, ""))[0]
            cls = {"name": "ok", "archetype": "ok",
                   "fallback": "warn"}.get(p.match_rule, "err")
            state = f'<span class="pill {cls}">{esc(label)}</span>'
            detail = esc(p.note)
        rows.append(
            f"<tr><td>{p.slide_index + 1}</td>"
            f"<td>{esc(p.source_layout)}"
            f"{f' <span class=note>({esc(p.source_type)})</span>' if p.source_type else ''}</td>"
            f"<td>&rarr; {esc(p.target_layout or '—')}</td>"
            f"<td>{state}</td><td class='note'>{detail}</td></tr>")
    return "".join(rows)


def _removed_block(changes: list, job_id: str, restored: list,
                   restore_error: str | None, notes: dict | None = None) -> str:
    """The removals, each with a tick to put it back.

    Listing the text was only half an answer: a designer who wanted it back had
    to retype it and place it by eye. Each piece is kept with its own XML, so
    ticking it returns the same words in the same box (design lead,
    20/08/2026)."""
    removed = [c for c in changes if getattr(c, "removed_text", None)]
    if not removed:
        return ""
    done = set(restored or [])
    notes = notes or {}
    rows = []
    for c in removed:
        rid = getattr(c, "restore_id", None)
        label = f"Slide {c.slide_index + 1}: <b>{esc(c.removed_text)}</b>"
        if rid and rid in done:
            what = notes.get(rid, "put back in place")
            rows.append(f'<li>{label} &mdash; '
                        f'<span class="pill ok">back in the deck</span> '
                        f'<span class="note">{esc(what)}</span></li>')
        elif rid and getattr(c, "removed_xml", None):
            rows.append(
                f'<li><label><input type="checkbox" name="restore_ids" '
                f'value="{esc(rid)}"> {label}</label></li>')
        else:
            rows.append(f'<li>{label} &mdash; <span class="note">no copy kept, '
                        f'so it has to go back by hand</span></li>')
    pending = [c for c in removed
               if getattr(c, "restore_id", None)
               and getattr(c, "removed_xml", None)
               and c.restore_id not in done]
    action = ""
    if pending:
        action = f"""
  <div class="actions" style="margin-top:0.7rem">
    <button class="btn ghost" type="submit">Put the ticked pieces back</button>
    <span class="note">Each comes back whole and exactly where it was - same
    wording, same box, same formatting, same position. Nothing else on the slide
    moves, so a piece whose old spot the master has taken will print over
    it.</span>
  </div>"""
    err = (f'<p class="note">The restore failed: {esc(restore_error)}</p>'
           if restore_error else "")
    return f"""
<form method="post" action="/format/{esc(job_id)}/restore">
<div class="banner warn">
  <b>&#33; {len(removed)} piece(s) of text were removed.</b> The master defines
  no placeholder for them. Tick anything that still matters and put it back;
  it returns with its own wording, box and formatting.
  <ul style="margin:0.5rem 0 0 1.1rem;list-style:none;padding-left:0">{''.join(rows)}</ul>
  {action}{err}
</div>
</form>"""


def _content_section(changes: list, job_id: str = "", restored: list = (),
                     restore_error: str | None = None,
                     restored_notes: dict | None = None) -> str:
    """What moved, grouped by kind and then listed per slide. Applying the
    layout and migrating the content are separate operations with separate
    failure modes, so they get separate sections rather than one blurred
    'done'."""
    if not changes:
        return ""
    by_action: dict[str, list] = {}
    for c in changes:
        by_action.setdefault(c.action, []).append(c)

    chips = "".join(
        f'<span class="fchip"><b>{len(items)}</b> {esc(action)}</span>'
        for action, items in sorted(by_action.items(),
                                    key=lambda kv: -len(kv[1])))

    def row(c) -> str:
        alert = getattr(c, "severity", "info") == "alert"
        # Not all alerts are removals: a heading left sitting past the margin is
        # an alert too, and labelling every one "content was removed" would send
        # a designer hunting for text that never left the deck.
        why = ("Content was removed" if getattr(c, "removed_text", None)
               else "Needs a designer's decision")
        mark = (f'<span class="pill err" title="{why}">'
                '&#33;</span> ' if alert else "")
        style = (' style="background:rgba(255,124,74,0.07)"' if alert else "")
        return (f"<tr{style}><td>{c.slide_index + 1}</td>"
                f"<td>{mark}<b>{esc(c.action)}</b></td>"
                f"<td class='note'>{esc(c.detail)}</td></tr>")

    # Alerts first, then by slide: removals are the rows a designer has to act
    # on, and burying them in slide order among routine moves is how content
    # goes missing unnoticed.
    rows = "".join(row(c) for c in sorted(
        changes,
        key=lambda c: (getattr(c, "severity", "info") != "alert",
                       c.slide_index, c.action)))

    removed_block = _removed_block(changes, job_id, restored, restore_error,
                                   restored_notes)

    misfits = [c for c in changes if c.action == "content does not fit"]
    warn = _warn(
        f"{len(misfits)} slide(s) have content taller or wider than the "
        f"master's content region. It was moved as far as it fits on the "
        f"canvas and left alone otherwise; nothing was scaled, because "
        f"shrinking a text box does not shrink its type."
    ) if misfits else ""

    return f"""
{removed_block}{warn}
<div class="card">
  <div class="tag">Content</div>
  <h2 style="margin-top:0">What moved into the master</h2>
  <p class="sub">Applying a layout only remaps content already in placeholders.
  Free-floating shapes stay put, which is why an applied master can look like
  nothing happened. These are the moves that followed.</p>
  <div class="kpis">{chips}</div>
  <table class="w3">
    <thead><tr><th>#</th><th>Change</th><th>Detail</th></tr></thead>
    <tbody>{rows}</tbody>
  </table>
</div>"""


def _masters_note(masters: int, stragglers: list, plans: list) -> str:
    """The consequence of a slide that could not be rebuilt: the deck's ORIGINAL
    master stays alive to serve it, so the output carries two masters and
    PowerPoint's master view lists the original FIRST.

    A designer opens that view, sees a master with none of the new guides and no
    presentation-space rectangle, and reads it as "the master was not copied"
    (design lead, 21/08/2026). It was: onto the other one. Saying so here is the
    difference between a five-minute check and an afternoon."""
    if masters <= 1 and not stragglers:
        return ""
    where = ", ".join(str(i + 1) for i in sorted(stragglers)[:12])
    more = f" and {len(stragglers) - 12} more" if len(stragglers) > 12 else ""
    slides = (f" Slide(s) {where}{more} are still on it." if stragglers else "")
    return _warn(
        f"This deck now carries {masters} slide masters, because not every "
        f"slide could be rebuilt on the new one.{slides} PowerPoint's master "
        f"view lists the ORIGINAL master first, so opening it shows a master "
        f"with none of the new guides, furniture or presentation-space "
        f"rectangle: that master is the leftover, not the applied one. Audit "
        f"the deck to see those slides as 'foreign master' findings, which the "
        f"fix engine can move onto the applied master.")


def render_format_result(*, deck_name: str, profile_name: str, job_id: str,
                         plans: list, errors: dict, applied: int,
                         content_changes: list | None = None,
                         restored: list | None = None,
                         restored_notes: dict | None = None,
                         restore_error: str | None = None,
                         masters: int = 1,
                         stragglers: list | None = None,
                         space_notes: list | None = None) -> str:
    counts = {}
    for p in plans:
        key = "failed" if p.slide_index in errors else p.match_rule
        counts[key] = counts.get(key, 0) + 1

    chips = []
    for rule in _RULE_ORDER:
        if counts.get(rule):
            label, why = _RULE_LABEL[rule]
            chips.append(f'<span class="fchip" title="{esc(why)}">'
                         f'<b>{counts[rule]}</b> {esc(label)}</span>')
    if counts.get("failed"):
        chips.append(f'<span class="fchip"><b>{counts["failed"]}</b> failed</span>')

    fallback_note = ""
    if counts.get("fallback"):
        fallback_note = _warn(
            f"{counts['fallback']} slide(s) had no matching layout in the "
            f"master and fell back to a content layout. Those are the slides "
            f"to open first: PowerPoint keeps unmatched content but leaves it "
            f"orphaned in place rather than in a placeholder.")

    restored_note = ""
    if restored:
        notes = list((restored_notes or {}).values())
        over = sum(1 for d in notes if "printing over" in d)
        restored_note = _warn(
            f"{len(restored)} removed piece(s) are back in the deck, each at the "
            f"exact position it was removed from and on top of its slide. "
            + (f"{over} print over something the master has since put in that "
               f"spot; they are named 'RESTORED ...' in PowerPoint's selection "
               f"pane, and each row below says what it covers. " if over else "")
            + "Download again to get the version with them: this puts content "
              "back, it does not lay it out.")

    failed_note = ""
    if errors:
        failed_note = _warn(
            f"{len(errors)} slide(s) could not be rebuilt and were left exactly "
            f"as they were. The deck is still usable; those slides simply did "
            f"not change.")

    # Whether the frame came across, stated rather than left to be discovered
    # in master view: a marker the designer drew on one layout serves only that
    # layout, and "the presentation space was not copied" is how that reads.
    space_note = ""
    if space_notes:
        space_note = ("<div class='card'><div class='tag'>Presentation space</div>"
                      "<ul style='margin:0.4rem 0 0 1.1rem'>"
                      + "".join(f"<li>{esc(n)}</li>" for n in space_notes)
                      + "</ul></div>")

    body = f"""
<h1>{esc(deck_name)}</h1>
<p class="sub">Rebuilt <b>{applied}</b> of <b>{len(plans)}</b> slides on
<b>{esc(profile_name)}</b>. Each slide was copied, given the master's layout,
and the original deleted.</p>
<div class="actions" style="gap:0.6rem">
  <a class="btn primary" href="/format/{esc(job_id)}/review?view=master">Review before / after</a>
  <a class="btn ghost" href="/format/{esc(job_id)}/download">Download the rebuilt deck</a>
  <a class="btn ghost" href="/format">Format another deck</a>
</div>
<p class="note">The review shows the master's layouts before and after, and the
deck slide by slide, with an Undo on every change. Look before you download:
undoing after the file has gone out means doing it twice.</p>
{fallback_note}{failed_note}{_masters_note(masters, stragglers or [], plans)}{restored_note}
<div class="kpis">{''.join(chips)}</div>
{space_note}
{_content_section(content_changes or [], job_id, restored or [], restore_error,
                  restored_notes or {})}
<div class="card">
  <div class="tag">Per slide</div>
  <h3 style="margin:0 0 0.2rem">Layout assignment</h3>
  <h2 style="margin-top:0">What happened to each slide</h2>
  <p class="sub">Layout matching runs by name first, then by archetype, then
  falls back. Every row says which rule chose its target, so a surprising
  result is traceable rather than mysterious.</p>
  <table class="w3">
    <thead><tr><th>#</th><th>Was on</th><th>Now on</th><th>How</th>
    <th>Notes</th></tr></thead>
    <tbody>{_slide_rows(plans, errors)}</tbody>
  </table>
</div>
<p class="note">Open the result in PowerPoint before sending it anywhere. This
step hands each slide to PowerPoint's placeholder matching, which preserves
content but can move it; that is exactly the judgment a designer signs off.</p>"""
    return _shell(f"Applied: {deck_name}", body)
