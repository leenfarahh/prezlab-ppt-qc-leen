"""The review page for a format run: what the master did, and how to take it back.

Two views of the same run, because a designer asks two different questions and
one page answering both answers neither well:

MASTER - every layout the submitted deck had, next to every layout it has now.
This is the question "is this actually our master?", and it is asked about the
template, not about any one slide. It is also where a leftover master shows
itself: a deck that could not rebuild every slide keeps its original master
alive, so the AFTER column lists layouts from both and says which master each
came from.

DECK - the slides themselves, before and after, with every change this run made
listed underneath and an Undo on each one. This is the question "do I accept
what happened to slide 7?".

Undo is per change and it is exact: each row replays the state the migration
stored before it acted (qc.undo), never a re-derived delta. Rows that changed
nothing - a heading reported as sitting past a margin, a slide that does not fit
- carry no Undo, and say so rather than offering a button that does nothing.
The layout assignment carries none either: PowerPoint performed it
(qc.applymaster), and the honest answer is to say so and point at the original
upload rather than to fake a reversal.

The page WORKS WITHOUT A RENDERER. Rendering needs desktop PowerPoint or
LibreOffice on the host, and neither is guaranteed; the change list and every
Undo button are plain HTML that do not depend on an image existing. A review a
designer cannot use because a render failed is not a review.

Rendering only; route logic lives in qc/web.py.
"""

from pathlib import Path

from .ui import _shell, esc

# Changes that report rather than act, and therefore have nothing to put back.
# Each gets a reason on the row: an empty Undo column reads as a missing button.
_WHY_NOTHING = {
    "content does not fit": "nothing was moved or scaled, so there is nothing "
                            "to put back; the block is where the frame left it",
    "wider than the margins": "reported only - narrowing a text box reflows "
                              "its text, so nothing was resized",
    "heading past the margin": "nothing was moved: whether a heading may break "
                               "the margin is the client's call",
    "body not seated on a stated frame": "a report about the master, not a "
                                         "change to this slide",
    "text overlaps text": "left exactly as the designer arranged it",
    "overlap needs a designer": "nothing was moved; it could not clear the "
                                "placeholder without leaving the canvas",
    "migration skipped": "the slide was not changed at all",
}


def _tabs(job_id: str, view: str) -> str:
    def tab(key, label):
        on = " primary" if view == key else " ghost"
        return (f'<a class="btn{on}" href="/format/{esc(job_id)}/review'
                f'?view={key}">{label}</a>')
    return (f'<div class="actionbar no-print">{tab("master", "Master")}'
            f'{tab("deck", "Deck")}<span class="grow"></span>'
            f'<a class="btn ghost" href="/format/{esc(job_id)}/download">'
            f'Download the deck</a>'
            f'<a class="btn ghost" href="/format">Format another deck</a></div>')


def _shot(job_id: str, key: str, alt: str, available: bool) -> str:
    """One picture, or an honest gap where it would be. `key` is the render's
    own key (qc.web._url_keys), used verbatim in the URL."""
    if not available:
        return ('<div class="shot"><p class="note" style="padding:1.2rem">'
                'Not rendered.</p></div>')
    return (f'<div class="shot"><img src="/review-img/{esc(job_id)}/'
            f'{esc(key)}.png" alt="{esc(alt)}" loading="lazy"></div>')


def _render_note(error: str | None) -> str:
    """Says what is missing and, just as important, what is not.

    Everything on this page except the pictures is read out of the deck by
    python-pptx: the layout names, the change list, every Undo. A banner that
    only reports the failure invites the reader to distrust the rest of it."""
    if not error:
        return ""
    wedged = "System call failed" in error or "Presentations.Open" in error
    how = (" This one usually means a leftover PowerPoint from an interrupted "
           "run: it is started with /AUTOMATION and has no window, so closing "
           "PowerPoint does not touch it. End POWERPNT.EXE in Task Manager and "
           "reload." if wedged else "")
    return (f'<div class="banner warn"><b>The pictures are missing.</b> '
            f'{esc(error)}{how} Everything else on this page is read from the '
            f'deck itself, not from the renders: the layouts below, the change '
            f'list, and every Undo.</div>')


# --------------------------------------------------------------- master view


def _layout_cards(job_id: str, side: str, entries: list, used: dict,
                  images: set) -> str:
    if not entries:
        return '<p class="note">No layouts to show.</p>'
    cards = []
    for e in entries:
        key = f"layout_{side}_{e['index']}"
        n = used.get(e["layout"], 0)
        badge = (f'<span class="pill ok">{n} slide{"s" if n != 1 else ""}</span>'
                 if n else '<span class="pill">unused</span>')
        master = (f'<span class="note">master {e["master"] + 1}</span>'
                  if e.get("master") else "")
        name = e["layout"]
        cards.append(
            f'<div class="pane"><div class="tag">{esc(name)}</div>'
            f'{_shot(job_id, key, f"Layout {name}", key in images)}'
            f'<div class="difflabels">{badge} {master}</div></div>')
    return f'<div class="lgrid">{"".join(cards)}</div>'


def _master_view(*, job_id: str, previews: dict, used_after: dict,
                 used_before: dict, masters: int, truncated: bool) -> str:
    images = set((previews or {}).get("images", {}))
    before = (previews or {}).get("before") or []
    after = (previews or {}).get("after") or []

    extra = ""
    if masters > 1:
        extra = ('<div class="banner warn">The rebuilt deck carries more than '
                 'one slide master, because not every slide could be rebuilt on '
                 'the new one. Layouts below are labelled with the master they '
                 'belong to; anything marked <b>master 2</b> or higher is a '
                 'leftover from the original design, not part of the applied '
                 'one.</div>')
    cut = ('<p class="note">Only the first 24 layouts of each side are shown; '
           'a master with more is truncated here, not in the deck.</p>'
           if truncated else "")

    return f"""
<p class="sub">One empty slide per layout, so what you see is the layout's own
furniture, placeholders and background with none of the deck's content on top.
This is the template comparison; the Deck tab is the slide-by-slide one.</p>
{extra}
<div class="card">
  <div class="tag">Before</div>
  <h2 style="margin-top:0">Layouts the deck arrived with</h2>
  {_layout_cards(job_id, "before", before, used_before, images)}
</div>
<div class="card">
  <div class="tag">After</div>
  <h2 style="margin-top:0">Layouts the deck has now</h2>
  {_layout_cards(job_id, "after", after, used_after, images)}
</div>{cut}"""


# ----------------------------------------------------------------- deck view


def _change_row(job_id: str, change, undone: set, notes: dict,
                brings: int = 0) -> str:
    cid = getattr(change, "change_id", None)
    alert = getattr(change, "severity", "info") == "alert"
    mark = '<span class="pill err" title="Needs a decision">&#33;</span> ' \
        if alert else ""
    action = f'{mark}<b>{esc(change.action)}</b>'

    if cid and cid in undone:
        control = ('<span class="pill ok">undone</span> '
                   f'<span class="note">{esc(notes.get(cid, ""))}</span>')
    elif getattr(change, "undo", None) and cid:
        # What the button will actually do, on the button. Each step on a slide
        # was computed on the state the steps before it left, so this one cannot
        # come back on its own; saying so beforehand is the difference between a
        # tool that surprises a designer and one that does not.
        with_it = (f'<div class="note">also takes back the {brings} later '
                   f'change{"s" if brings != 1 else ""} on this slide</div>'
                   if brings else "")
        control = (f'<button class="btn ghost" type="submit" name="change_ids" '
                   f'value="{esc(cid)}">Undo</button>{with_it}')
    else:
        why = _WHY_NOTHING.get(change.action, "this pass recorded it but did "
                                              "not change the slide")
        control = f'<span class="note">{esc(why)}</span>'

    style = ' style="background:rgba(255,124,74,0.07)"' if alert else ""
    return (f"<tr{style}><td>{action}</td>"
            f"<td class='note'>{esc(change.detail)}</td>"
            f"<td style='white-space:nowrap'>{control}</td></tr>")


def _plan_line(plan, error: str | None) -> str:
    if plan is None:
        return ""
    if error:
        return (f'<div class="difflabels"><span class="pill error">failed</span> '
                f'{esc(error)}</div>')
    return (f'<div class="difflabels">Was on <b>{esc(plan.source_layout)}</b> '
            f'&rarr; now on <b>{esc(plan.target_layout or "—")}</b> '
            f'<span class="note">({esc(plan.match_rule)})</span>'
            f'<span class="note"> &middot; the layout assignment is '
            f'PowerPoint\'s own work and cannot be undone one slide at a time; '
            f'the original upload is the way back</span></div>')


def _pager(job_id: str, page: int, page_size: int, reviewable: int) -> str:
    """Where you are and how to get to the rest.

    A review that stops at slide 20 of a 26-slide deck and says nothing reads as
    six slides the tool declined to show (design lead, 23/08/2026). The count is
    stated whether or not there is a next page."""
    pages = max(1, -(-reviewable // page_size))
    first = page * page_size + 1
    last = min(reviewable, (page + 1) * page_size)
    where = (f'<span class="note">Slides {first}&ndash;{last} of '
             f'{reviewable} with changes'
             + (f", page {page + 1} of {pages}" if pages > 1 else "")
             + '</span>')
    if pages <= 1:
        return f'<div class="difflabels">{where}</div>'

    def link(target, label, on):
        if not on:
            return f'<span class="btn ghost" aria-disabled="true">{label}</span>'
        return (f'<a class="btn ghost" href="/format/{esc(job_id)}/review'
                f'?view=deck&amp;page={target}">{label}</a>')
    return (f'<div class="actionbar no-print">{where}<span class="grow"></span>'
            f'{link(page - 1, "&larr; Previous", page > 0)}'
            f'{link(page + 1, "Next &rarr;", page + 1 < pages)}</div>')


def _deck_view(*, job_id: str, previews: dict, changes: list, plans: list,
               errors: dict, undone: set, notes: dict, shown: list,
               total_slides: int, reviewable: int = 0, page: int = 0,
               page_size: int = 20) -> str:
    images = set((previews or {}).get("images", {}))
    by_slide: dict[int, list] = {}
    for c in changes:
        by_slide.setdefault(c.slide_index, []).append(c)
    by_index = {p.slide_index: p for p in plans}

    blocks = []
    for idx in shown:
        on_slide = by_slide.get(idx, [])
        # How many undoable changes sit AFTER each one, on this slide. Counted
        # here rather than asked per row so the arithmetic is done once and the
        # page and qc.undo.followers cannot disagree about it.
        later = {}
        seen = 0
        for c in reversed(on_slide):
            later[id(c)] = seen
            if getattr(c, "undo", None):
                seen += 1
        rows = "".join(_change_row(job_id, c, undone, notes, later[id(c)])
                       for c in on_slide)
        if not rows:
            rows = ('<tr><td colspan="3" class="note">Nothing was changed on '
                    'this slide beyond the layout assignment.</td></tr>')
        blocks.append(f"""
<div class="diffslide">
 <h3>Slide {idx + 1}</h3>
 {_plan_line(by_index.get(idx), errors.get(idx))}
 <div class="panes">
  <div class="pane"><div class="tag">Before</div>
   {_shot(job_id, f"slide_before_{idx}", f"Slide {idx + 1} as uploaded",
          f"slide_before_{idx}" in images)}</div>
  <div class="pane"><div class="tag">After</div>
   {_shot(job_id, f"slide_after_{idx}", f"Slide {idx + 1} rebuilt",
          f"slide_after_{idx}" in images)}</div>
 </div>
 <table class="w3">
  <thead><tr><th>Change</th><th>Detail</th><th>Put it back</th></tr></thead>
  <tbody>{rows}</tbody>
 </table>
</div>""")

    # What is NOT on this page, and why, in both directions: the slides paged
    # past and the slides this run never touched. The deck always has all of
    # them; only the review is paginated.
    untouched = total_slides - reviewable
    cut = (f'<p class="note">The deck has all {total_slides} slides. '
           f'{reviewable} of them were changed by this run and are reviewable '
           f'here; the other {untouched} came through the content pass '
           f'unchanged, so there is nothing to compare on them.</p>'
           if untouched > 0 else
           f'<p class="note">The deck has all {total_slides} slides and this '
           f'run changed every one of them.</p>')
    if not blocks:
        blocks = ['<p class="note">This run changed no slide content, so there '
                  'is nothing to review here. The Master tab still shows what '
                  'the template did.</p>']

    pager = _pager(job_id, page, page_size, reviewable) if reviewable else ""
    return f"""
<p class="sub">Each slide as it arrived and as it stands now, with every change
this run made underneath. Undo puts one change back exactly as it was - same
box, same wording, same coordinates - and leaves the rest of the slide alone.</p>
{cut}{pager}
<form method="post" action="/format/{esc(job_id)}/undo">
<input type="hidden" name="page" value="{page}">
{''.join(blocks)}
</form>{pager}"""


# --------------------------------------------------------------------- shell


def render_review(*, deck_name: str, profile_name: str, job_id: str,
                  view: str = "master", previews: dict | None = None,
                  changes: list | None = None, plans: list | None = None,
                  errors: dict | None = None, undone: set | None = None,
                  notes: dict | None = None, shown: list | None = None,
                  reviewable: int = 0, page: int = 0, page_size: int = 20,
                  total_slides: int = 0, masters: int = 1,
                  used_before: dict | None = None,
                  used_after: dict | None = None,
                  truncated: bool = False,
                  render_error: str | None = None,
                  undo_error: str | None = None) -> str:
    view = "deck" if view == "deck" else "master"
    head = f"""
<span class="kicker">Review &middot; {esc(profile_name)}</span>
<h1 class="file">{esc(Path(deck_name).name)}</h1>
{_tabs(job_id, view)}{_render_note(render_error)}"""
    if undo_error:
        head += f'<div class="banner warn">The undo failed: {esc(undo_error)}</div>'

    if view == "master":
        body = _master_view(job_id=job_id, previews=previews or {},
                            used_after=used_after or {},
                            used_before=used_before or {},
                            masters=masters, truncated=truncated)
    else:
        body = _deck_view(job_id=job_id, previews=previews or {},
                          changes=changes or [], plans=plans or [],
                          errors=errors or {}, undone=undone or set(),
                          notes=notes or {}, shown=shown or [],
                          total_slides=total_slides,
                          reviewable=reviewable or len(shown or []),
                          page=page, page_size=page_size)
    style = """
<style>
.lgrid{display:grid;gap:0.9rem;grid-template-columns:repeat(auto-fill,minmax(230px,1fr))}
.lgrid .pane{margin:0}
</style>"""
    return _shell(f"Review: {deck_name}", head + body + style)
