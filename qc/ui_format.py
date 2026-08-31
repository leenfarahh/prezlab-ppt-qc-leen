"""The blocks that describe applying a master, for the page that does it.

There is no formatting page any more. Applying a master is step 2 of Prepare a
deck, and this module is what that page reads a rebuild through: how each slide
was matched, what the migration moved, and what it took out and can put back.

Rendering only, and fragments only - the shell belongs to qc/ui_prep.py. Route
logic lives in qc/web.py.
"""

from .ui import esc

RULE_LABEL = {
    "name": ("matched by name", "The deck's layout and the master's share a "
                                "name, so the designer meant them to correspond."),
    "archetype": ("matched by archetype", "No name match, but both layouts "
                                          "declare the same OOXML archetype."),
    "fallback": ("fell back", "Neither the name nor the archetype matched. "
                              "Content may be orphaned; check these slides."),
    "reviewed": ("placed by review",
                 "Neither the name nor the archetype matched, so the slide and "
                 "the master's layouts were looked at and this one was chosen "
                 "for the shape of the content. Confirm it in the review "
                 "below."),
    "reviewed (uncertain)": (
        "placed by review, unsure",
        "Chosen by looking, but two layouts would have served about equally "
        "well or the structure had no real counterpart. Open these first."),
    "none": ("no target", "The master defines no layout that could be used."),
}
RULE_ORDER = ("name", "archetype", "reviewed", "reviewed (uncertain)",
               "fallback", "none")


def _coverage_block(coverage) -> str:
    """The deck-level layout gap report, when the route computed one.

    Rendered by qc.ui_check, which owns this block because the re-check page
    is where it is read most. Imported here rather than at module scope: that
    module reads the rule labels out of this one, and a top-level import in both
    directions is a cycle."""
    if coverage is None:
        return ""
    from .ui_check import render_coverage

    return render_coverage(coverage)


def _warn(message: str) -> str:
    return f'<div class="banner warn">{esc(message)}</div>' if message else ""


def _slide_rows(plans: list, errors: dict) -> str:
    rows = []
    for p in plans:
        err = errors.get(p.slide_index)
        if err:
            state = f'<span class="pill error">failed</span>'
            detail = esc(err)
        else:
            label = RULE_LABEL.get(p.match_rule, (p.match_rule, ""))[0]
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


def _proposed_block(changes: list, job_id: str, removed: list,
                    remove_error: str | None = None) -> str:
    """What the pass found that it would take out, each with a tick.

    Nothing here has happened yet, and that is the point (design lead,
    26/08/2026): the migration used to remove these and offer them back, so a
    designer's first sight of the rebuilt deck was already missing things. The
    wording is deliberately in the present tense - "is a second copy", not "was
    removed" - because every one of these is still on the slide.
    """
    proposals = [c for c in changes if getattr(c, "remove_op", None)]
    if not proposals:
        return ""
    done = set(removed or [])
    pending = [c for c in proposals if c.remove_id not in done]

    rows = []
    for c in proposals:
        what = esc(c.removed_text or c.action)
        label = f"Slide {c.slide_index + 1}: <b>{what}</b>"
        if c.remove_id in done:
            rows.append(f'<li>{label} &mdash; '
                        f'<span class="pill ok">taken out</span> '
                        f'<span class="note">undo it on the review page</span>'
                        f'</li>')
        else:
            rows.append(
                f'<li><label><input type="checkbox" name="remove_ids" '
                f'value="{esc(c.remove_id)}"> {label}</label>'
                f'<div class="note" style="margin-left:1.4rem">'
                f'{esc(c.detail)}</div></li>')

    action = ""
    if pending:
        action = """
  <div class="actions" style="margin-top:0.7rem">
    <button class="btn ghost" type="submit"
            data-busy="Taking out the ticked pieces">Remove the ticked
    pieces</button>
    <span class="note">Each comes out with an Undo beside it on the review page,
    so this is reversible. Nothing else on the slide moves.</span>
  </div>"""
    err = (f'<p class="note">The removal failed: {esc(remove_error)}</p>'
           if remove_error else "")

    return f"""
<form method="post" action="/format/{esc(job_id)}/remove">
<div class="banner warn">
  <b>&#33; {len(pending)} piece(s) this pass would take out, and did not.</b>
  Nothing leaves a slide unless you say so, so they are all still in the deck.
  Tick anything the master now supplies or that is a leftover, and it comes out
  with an Undo attached.
  <ul style="margin:0.5rem 0 0 1.1rem;list-style:none;padding-left:0">{''.join(rows)}</ul>
  {action}{err}
</div>
</form>"""


def _removed_block(changes: list, job_id: str, restored: list,
                   restore_error: str | None, notes: dict | None = None) -> str:
    """The removals, each with a tick to put it back.

    Listing the text was only half an answer: a designer who wanted it back had
    to retype it and place it by eye. Each piece is kept with its own XML, so
    ticking it returns the same words in the same box (design lead,
    20/08/2026)."""
    # Only pieces that really have gone. A proposal carries removed_text too -
    # it is what the piece SAYS, and the page shows it either way - so keying
    # off that alone put every proposal in a block headed "were removed".
    removed = [c for c in changes if getattr(c, "removed_text", None)
               and not getattr(c, "remove_op", None)]
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
                     restored_notes: dict | None = None,
                     removed: list = (),
                     remove_error: str | None = None) -> str:
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
        why = ("Left in place; you decide whether it goes"
               if getattr(c, "remove_op", None)
               else "Content was removed" if getattr(c, "removed_text", None)
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

    # Proposals first: they are the only rows that ask the designer for
    # something, and a deck they have not looked at yet is a deck still carrying
    # everything this pass found.
    proposed_block = _proposed_block(changes, job_id, removed, remove_error)
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
{proposed_block}{removed_block}{warn}
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
