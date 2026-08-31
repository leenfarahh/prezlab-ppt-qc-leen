"""Layout coverage: what a master can and cannot build.

The block renders in step 1 of the prepared deck's page, because "eleven slides
fell back" is the sentence a designer needs while they still have the before
and after in front of them. It used to have a pre-flight page of its own; that
was a second door onto the same report and a second deck to upload, and the
report is now part of every prepare run whether or not anyone asked for it.

One thing still has its own answer: the RE-CHECK. A designer who builds the
missing layout needs to know whether the gap closed, and that question is about
a master that does not exist yet, so it takes an upload and gets a page
(render_check_result). It reads the deck the prepare run already holds.

Rendering only; route logic lives in qc/web.py.
"""

from .layoutgap import headline
from .ui import _shell, esc
from .ui_format import RULE_LABEL, RULE_ORDER, _warn


def _slide_list(indices: list[int], limit: int = 24) -> str:
    """Slide numbers, 1-based, because that is what PowerPoint's own numbering
    says and a designer is going to type them into the go-to box."""
    shown = ", ".join(str(i + 1) for i in indices[:limit])
    if len(indices) > limit:
        shown += f", and {len(indices) - limit} more"
    return shown


def _gap_card(gap) -> str:
    # A slide grouped with one that WAS looked at carries the same answer,
    # because the structure and the source layout are the same. Saying so is
    # the difference between a checked claim and an inferred one, and a
    # designer deciding whether to add a layout to a client's master is
    # entitled to know which they are reading.
    def _split(n_total: int, n_asked: int) -> str:
        repeats = n_total - n_asked
        if not repeats:
            return ""
        return (f" ({n_asked} looked at, {repeats} sharing the same structure "
                f"and the same source layout)")

    confirmed = ""
    if gap.refused:
        confirmed = (
            f'<p class="note"><b>{gap.refused}</b> of these were checked '
            f'against every layout in the master and none of them fit'
            f'{_split(gap.refused, min(gap.asked, gap.refused))}. That is the '
            f'strongest evidence here: not a missed name match, a missing '
            f'layout.</p>')
    elif gap.reviewed:
        confirmed = (f'<p class="note">{gap.reviewed} answered'
                     f'{_split(gap.reviewed, gap.asked)}; the rest were not '
                     f'reached.</p>')

    closest = ""
    if gap.closest:
        closest = (f'<p class="note">Closest thing the master has: '
                   f'<b>{esc(gap.closest)}</b>. It {esc(gap.closest_note)}</p>')

    came_from = ""
    if gap.source_layouts:
        names = ", ".join(esc(n) for n in gap.source_layouts[:6])
        came_from = (f'<p class="note">In the submitted deck these were on: '
                     f'{names}'
                     + (", and others" if len(gap.source_layouts) > 6 else "")
                     + ".</p>")

    return f"""<div class="card">
  <div class="tag">{gap.places} slide{'s' if gap.places != 1 else ''}</div>
  <h3 style="margin:0 0 0.3rem">{esc(gap.label)}</h3>
  <p class="sub" style="margin:0 0 0.6rem">Slides {_slide_list(gap.slides)}</p>
  {closest}{came_from}{confirmed}
</div>"""


_SUGGEST_CSS = """
<style>
.sug { display:grid; grid-template-columns:minmax(14rem,1fr) 1.4fr; gap:1.1rem;
  align-items:start; }
@media (max-width: 52rem) { .sug { grid-template-columns:1fr; } }
.sug .boxes { display:flex; flex-wrap:wrap; gap:0.3rem; margin:0.5rem 0 0; }
.sug .boxes span { border:1px solid var(--line); border-radius:999px;
  padding:0.1rem 0.55rem; font-size:0.74rem; color:var(--teal); }
</style>
"""


def _suggestion_card(s, picture: str) -> str:
    """One layout to add: what to call it, what goes in it, and who it serves."""
    boxes = "".join(
        f'<span>{esc(b.get("label") or b["kind"])}'
        + (f' &middot; col {b["column"]}' if b["column"] else " &middot; full width")
        + '</span>'
        for b in s.boxes)

    collision = ""
    if s.collides_with:
        collision = _warn(
            f"The master already has a layout called "
            f"{s.collides_with!r}. That makes this a naming problem rather than "
            f"a missing layout: the slides did not reach it, and matching runs "
            f"on the layout NAME first. Renaming the deck's own layout to match, "
            f"or the master's, may be the whole fix.")

    return f"""<div class="card">
  <div class="tag">{s.places} slide{'s' if s.places != 1 else ''} would use it</div>
  <div class="sug">
    <div>
      <h3 style="margin:0 0 0.2rem">{esc(s.name)}</h3>
      <p class="note" style="margin:0 0 0.4rem">Archetype
      <code>{esc(s.archetype)}</code>, {s.columns} column{'s' if s.columns != 1 else ''}.</p>
      <p class="sub" style="margin:0 0 0.4rem">{esc(s.why)}</p>
      <div class="boxes">{boxes}</div>
      <p class="note" style="margin:0.6rem 0 0">For the slides that want
      {esc(s.gap_label)}: {_slide_list(s.serves)}.</p>
    </div>
    <div>{picture}</div>
  </div>
  {collision}
</div>"""


def render_suggestions(suggestions: list, pictures: dict, asked: int = 0,
                       note: str = "", propose_job: str = "",
                       offer: bool = False) -> str:
    """The window: one card per layout the master should have.

    Below the gaps rather than instead of them, because the gap is the evidence
    and this is the proposal. A designer who disagrees with the proposal still
    has the six slides that could not be placed.
    """
    if not suggestions:
        if not note:
            return ""
        return f'<div class="card"><div class="tag">Suggested layouts</div>' \
               f'<p class="note" style="margin:0.3rem 0 0">{esc(note)}</p></div>'

    cards = "".join(_suggestion_card(s, pictures.get(i, ""))
                    for i, s in enumerate(suggestions))
    return f"""
<div class="card" style="background:none;border:0;padding:0">
  <div class="tag">Suggested layouts</div>
  <h2 style="margin:0 0 0.4rem">What to add to the master</h2>
  <p class="sub" style="margin:0 0 0.8rem">One proposal per group of slides that
  had nowhere to go. The boxes are placed on the master's own frame, so the
  wireframe shows where they would sit against the client's margins. Nothing is
  built: a layout carries type styles, guides and brand furniture that are a
  designer's to add, and this tool does not edit a client's master.</p>
  {cards}
</div>{_SUGGEST_CSS}"""


def _capitalise(text: str) -> str:
    return (text[:1].upper() + text[1:]) if text else text


def _misfit_block(cov) -> str:
    """Slides that were PLACED on a layout their content does not fit.

    A different sentence from a gap, and a designer acts on it differently: the
    master has a layout for these slides and it is the wrong shape. Matching runs
    by layout NAME first and a name match is never questioned, because the
    designer called both layouts the same thing and meant them to correspond -
    which is a claim about intent and says nothing about whether the content fits
    the boxes. This is the check that reads the slides anyway (design lead,
    26/08/2026).
    """
    misfits = list(getattr(cov, "misfits", None) or [])
    if not misfits:
        return ""

    by_layout: dict[str, list] = {}
    for m in misfits:
        by_layout.setdefault(m.layout, []).append(m)

    cards = []
    for layout, members in sorted(by_layout.items(),
                                  key=lambda kv: -len(kv[1])):
        first = members[0]
        cards.append(f"""<div class="card">
  <div class="tag">{len(members)} slide{'s' if len(members) != 1 else ''} on
  {esc(layout)}</div>
  <h3 style="margin:0 0 0.3rem">{esc(first.label)}</h3>
  <p class="sub" style="margin:0 0 0.5rem">Slides
  {_slide_list([m.slide_index for m in members])}, placed by
  {esc(first.rule)}.</p>
  <p class="note">{esc(_capitalise(first.offers))}, and {esc(first.reason)}.
  PowerPoint will remap what it can and leave the rest where it is.</p>
</div>""")

    # No "confirmed by the model" line any more: the vision pass that produced
    # that evidence was replaced by the layout step, where a designer answers
    # this themselves rather than being told a second opinion agreed
    # (31/08/2026). The claim is the same claim it always was - a shape
    # comparison, stated as one.
    head = _warn(
        f"{len(misfits)} slide(s) were placed on a layout their content does not "
        f"fit. The names matched, so the format pass will use them - and the "
        f"content is the wrong shape for the boxes, so part of it will be left "
        f"outside a placeholder. Each of these is offered on the layout step, "
        f"where you can send it somewhere else.")

    return f"""
<div class="card" style="background:none;border:0;padding:0">
  <div class="tag">Placed, but the wrong shape</div>
  <h2 style="margin:0 0 0.4rem">Slides on a layout that does not fit them</h2>
  {head}{''.join(cards)}
</div>"""


def _recheck_form(check_id: str, master_name: str) -> str:
    """The loop: build the layout, come back with the revised master.

    Only the master is asked for. The deck is the one already checked and is held
    in memory for it, because a designer revising a layout three times should
    upload three masters rather than three masters and three decks."""
    if not check_id:
        return ""
    against = (f" It last checked against <b>{esc(master_name)}</b>."
               if master_name else "")
    return f"""
<div class="card">
  <div class="tag">When you have built it</div>
  <h3 style="margin:0 0 0.3rem">Check the revised master</h3>
  <p class="sub" style="margin:0 0 0.6rem">Add the layouts above to the client's
  master in PowerPoint, then bring the file back here. The same deck is checked
  again against it, so you can see whether the gaps actually closed.{against}</p>
  <form action="/check/{esc(check_id)}/again" method="post"
        enctype="multipart/form-data">
    <input type="file" name="master" accept=".pptx" required>
    <div class="actions" style="margin-top:0.7rem">
      <button class="btn primary" type="submit"
              data-busy="Checking the revised master">Check again</button>
      <span class="note">The revised master is read for this check and dropped.
      It does not replace the one the profile carries; that is a separate
      decision when you save the master as a profile.</span>
    </div>
  </form>
</div>"""


def render_coverage(cov, *, standalone: bool = False,
                    suggestions: list | None = None,
                    pictures: dict | None = None,
                    suggest_note: str = "",
                    check_id: str = "", master_name: str = "",
                    propose_job: str = "") -> str:
    """The report itself. `standalone` only changes the wording of the closing
    line: on a re-check nothing has been rewritten, and on the prepared deck's
    page it already has."""
    if cov is None or not cov.slides:
        return ""

    chips = []
    for rule in RULE_ORDER:
        if cov.by_rule.get(rule):
            label, why = RULE_LABEL[rule]
            chips.append(f'<span class="fchip" title="{esc(why)}">'
                         f'<b>{cov.by_rule[rule]}</b> {esc(label)}</span>')

    if not cov.unplaced:
        gaps = ('<div class="banner ok">Every slide in this deck has a layout in '
                'this master.'
                + ('</div>' if getattr(cov, "misfits", None) else
                   ' Nothing is missing.</div>'))
    else:
        gaps = "".join(_gap_card(g) for g in cov.gaps)

    # A slide nobody decided about is a weaker claim than a slide somebody
    # looked at and rejected every layout for, and the two are indistinguishable
    # on this page unless it says which. Only the second is evidence that the
    # master is missing something.
    caveat = ""
    if cov.unplaced and not cov.review_ran:
        caveat = _warn(
            "These are grouped from the file alone: the layouts were matched by "
            "name and by archetype, and nobody has chosen for the rest. A slide "
            "listed here may well have a home in this master under a different "
            "name - the layout step is where you say so.")
    elif cov.not_reviewed:
        caveat = _warn(
            f"{cov.not_reviewed} of these were left on the fallback rather than "
            f"given a layout, so they are grouped from the file alone. They may "
            f"still have a home in this master under a different name.")

    unused = ""
    if cov.unused_layouts:
        names = ", ".join(esc(n) for n in cov.unused_layouts[:14])
        more = (f", and {len(cov.unused_layouts) - 14} more"
                if len(cov.unused_layouts) > 14 else "")
        unused = f"""<div class="card">
  <div class="tag">Unused</div>
  <h3 style="margin:0 0 0.3rem">{len(cov.unused_layouts)}
  layout{'s' if len(cov.unused_layouts) != 1 else ''} in this master that this
  deck never lands on</h3>
  <p class="note">{names}{more}. Not a fault: a master carries a whole
  vocabulary and one deck uses part of it. Worth a look only when a layout here
  is plainly what one of the slides above was asking for, which means the names
  and the archetypes did not line up.</p>
</div>"""

    closing = ("<p class=\"note\">Nothing has been changed. This read the deck "
               "and the master and put them both down again.</p>"
               if standalone else "")

    proposed = render_suggestions(
        suggestions or [], pictures or {}, note=suggest_note,
        propose_job=propose_job,
        offer=bool(cov.unplaced or getattr(cov, "misfit_clusters", None)))
    misfit = _misfit_block(cov)
    recheck = _recheck_form(check_id, master_name) if (
        cov.unplaced or getattr(cov, "misfits", None)) else ""

    return f"""<div class="card" style="background:none;border:0;padding:0">
  <div class="tag">Layout coverage</div>
  <h2 style="margin:0 0 0.4rem">What this master can build</h2>
  <p class="sub" style="margin:0 0 0.7rem">{esc(headline(cov))}</p>
  <div class="kpis" style="margin-bottom:0.9rem">
    <div class="kpi"><div class="n">{cov.matched}</div>
      <div class="l">placed</div></div>
    <div class="kpi{' warn' if cov.unplaced else ''}">
      <div class="n">{cov.unplaced}</div><div class="l">no layout</div></div>
    <div class="kpi info"><div class="n">{len(cov.gaps)}</div>
      <div class="l">distinct gaps</div></div>
  </div>
  <div class="controls" style="margin-bottom:0.6rem">{''.join(chips)}</div>
  {caveat}{gaps}{misfit}{proposed}{recheck}{unused}{closing}
</div>"""


def render_check_result(*, deck_name: str, profile_name: str, profile_id: str,
                        coverage, look_note: str = "",
                        suggestions: list | None = None,
                        pictures: dict | None = None,
                        suggest_note: str = "",
                        check_id: str = "", master_name: str = "") -> str:
    """The revised master, read against the deck already prepared under this id.

    A view of one run, not a door of its own: the only way here is the "check
    again" form on the prepared deck, so the way out is back to that deck."""
    # The re-check runs under the prepare job's own id, so it is the prepared
    # deck this returns to. Falls back to a fresh start only if the id is gone.
    back = f"/prep/{esc(check_id)}" if check_id else "/prep"
    body = f"""
<h1>{esc(deck_name)}</h1>
<p class="sub">The revised master checked against <b>{esc(deck_name)}</b>, the
deck prepared against <b>{esc(profile_name)}</b>. Nothing was changed and no
file was written: this answers whether the gap closed, nothing more.</p>
{_warn(look_note)}
{render_coverage(coverage, standalone=True, suggestions=suggestions,
                pictures=pictures, suggest_note=suggest_note,
                check_id=check_id, master_name=master_name)}
<div class="actions" style="gap:0.6rem">
  <a class="btn primary" href="{back}">Back to the prepared deck</a>
</div>
<p class="note">A layout added after the rebuild does not reach the slides that
already fell back. If the gaps above are closed, save the revised master as the
profile in step 1 of <a href="/prep">Prepare a deck</a> and prepare the deck
again; that is the pass those slides need.</p>"""
    return _shell(f"Coverage: {deck_name}", body)
