"""Step 2: the layouts, approved before anything is rebuilt.

The one page between dropping a deck and getting a rebuilt one. It exists
because applying a master used to be a single press that guessed at the slides
it could not place, and a designer only found out what it had guessed by opening
the result - by which point the guess was already in the file.

IT SHOWS EVERY SLIDE, AND IT LEADS ON THE UNCERTAIN ONES. Those are two
different things and the page used to conflate them: a slide that matched by
name onto a layout its content fits is not a question, so it was left off
entirely and replaced by a line saying how many were "not listed because there
is nothing to decide about them". A designer approving a rebuild before it
happens wants to see the deck (design lead, 02/09/2026).

So every slide has a card, in deck order, and the ones that are questions are
what the count, the ordering and the colour lead on. A matched slide sits
folded, pre-selected onto the layout it is already going to, and changing it is
one click - which is the thing that was impossible before, because the slide was
not on the page to change. Leaving it alone records nothing: a designer who
never opened a card has not decided anything about it, and the coverage report
reads that difference (qc.layoutpick.apply_picks).

Every option is one of the master's OWN layouts, shown as the master draws it,
picked by radio. There is no free-text box and nothing invents a layout: the
list is closed, and the last option on it is always "leave it". That one is not
an escape hatch, it is the most informative answer on the page - a designer who
reads the master against this slide and says none of these fit has established
that the master is missing a layout, which is exactly what the coverage report
is trying to find out (qc.layoutpick.apply_picks).

Rendering only. The ranking is qc.layoutpick's and the rebuild is qc.prep's.
"""

from .layoutpick import LEAVE
from .ui import _shell, esc

_CSS = """
<style>
.lbar { position: sticky; top: 0; z-index: 10; background: #fff;
  border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 0.7rem 1.1rem; margin: 0.4rem 0 1.2rem; display: flex; gap: 1rem;
  align-items: center; flex-wrap: wrap;
  box-shadow: 0 2px 12px rgba(0, 37, 40, 0.05); }
.lbar .grow { flex: 1; }
.lbar .count { font-weight: 700; font-variant-numeric: tabular-nums; }

.slidecard { background: #fff; border: 1px solid var(--line-soft);
  border-radius: 12px; padding: 1.1rem 1.2rem; margin: 0 0 1rem; }
.slidehead { display: flex; gap: 0.7rem; align-items: baseline;
  flex-wrap: wrap; margin-bottom: 0.15rem; }
.slidehead .n { font-weight: 800; font-size: 1.05rem; }
.slidehead .tag { font-size: 11px; font-weight: 700; letter-spacing: 0.1em;
  text-transform: uppercase; border-radius: 999px; padding: 0.15rem 0.6rem; }
.tag.unplaced { background: var(--sand); color: var(--burgundy); }
.tag.misfit { background: var(--lavender); color: var(--burgundy); }
.tag.matched { background: var(--teal-light); color: var(--teal); }

/* A matched slide is on the page to be SEEN and, if anyone disagrees, changed.
   Folded by default so the questions are what a designer scrolls through, and
   quieter so the two kinds are never mistaken for each other at a glance. */
.slidecard.settled { border-color: var(--line-soft); background: var(--hover); }
.slidecard.settled > summary { list-style: none; cursor: pointer;
  display: flex; gap: 0.7rem; align-items: baseline; flex-wrap: wrap; }
.slidecard.settled > summary::-webkit-details-marker { display: none; }
.slidecard.settled > summary::after { content: "Change layout";
  color: var(--teal); font-size: 0.82rem; font-weight: 700; margin-left: auto; }
.slidecard.settled[open] > summary::after { content: "Close"; }
.slidecard.settled > summary:hover { color: var(--teal); }
.slidecard.settled .goes { color: var(--slate-text); font-size: 0.9rem; }
.slidecard.settled .goes b { color: var(--teal); }
.slidecard.settled .pickrow { margin-top: 0.9rem; }
.why { color: var(--slate-text); font-size: 0.9rem; margin: 0 0 0.9rem; }
.why b { color: var(--teal); }

.pickrow { display: flex; gap: 1.1rem; align-items: flex-start;
  flex-wrap: wrap; }
.thisslide { flex: 0 0 15rem; }
.thisslide img { width: 100%; border: 1px solid var(--line); border-radius: 8px;
  display: block; background: #fff; }
.thisslide .cap { color: var(--slate-text); font-size: 0.8rem;
  margin-top: 0.35rem; }
.noshot { border: 1px dashed var(--line); border-radius: 8px; padding: 1.6rem 1rem;
  text-align: center; color: var(--slate-text); font-size: 0.85rem; }

.options { flex: 1 1 22rem; display: grid; gap: 0.5rem;
  grid-template-columns: repeat(auto-fill, minmax(11rem, 1fr)); }
.opt { position: relative; display: block; cursor: pointer; }
.opt input { position: absolute; opacity: 0; pointer-events: none; }
.opt .box { border: 1px solid var(--line); border-radius: 10px;
  padding: 0.5rem; height: 100%; background: #fff; }
.opt:hover .box { background: var(--hover); }
.opt input:checked + .box { border-color: var(--teal); border-width: 2px;
  padding: calc(0.5rem - 1px); background: var(--hover); }
.opt input:focus-visible + .box { outline: 2px solid var(--orange);
  outline-offset: 2px; }
.opt img { width: 100%; border: 1px solid var(--line-soft); border-radius: 6px;
  display: block; margin-bottom: 0.4rem; background: #fff; }
.opt .nm { font-weight: 700; font-size: 0.88rem; overflow-wrap: anywhere; }
.opt .of { color: var(--slate-text); font-size: 0.78rem; margin-top: 0.15rem; }
.opt .bad { color: var(--burgundy); font-size: 0.78rem; margin-top: 0.15rem; }
.opt .rec { display: inline-block; font-size: 10px; font-weight: 800;
  letter-spacing: 0.08em; text-transform: uppercase; color: var(--teal);
  background: var(--teal-light); border-radius: 999px;
  padding: 0.1rem 0.45rem; margin-bottom: 0.25rem; }

.more { margin-top: 0.7rem; font-size: 0.88rem; color: var(--slate-text); }
.more select { font: inherit; padding: 0.35rem 0.5rem; border-radius: 8px;
  border: 1px solid var(--line); background: #fff; color: var(--teal);
  max-width: 100%; }
.lbar .restcount { color: var(--slate-text); font-size: 0.9rem; }
</style>
"""


def _shot(src: str | None, alt: str, missing: str) -> str:
    if not src:
        return f'<div class="noshot">{esc(missing)}</div>'
    return f'<img src="{esc(src)}" alt="{esc(alt)}" loading="lazy">'


def _option(choice, cand, index: int, thumbs: dict, checked: bool) -> str:
    """One layout, as the master draws it. `cand` is a layoutpick.Candidate."""
    shot = thumbs.get(cand.name)
    detail = (f'<div class="of">{esc(cand.offers)}</div>' if cand.fits
              else f'<div class="bad">{esc(cand.why or cand.offers)}</div>')
    # On a settled slide the first option is where the slide ALREADY IS, not a
    # ranking's preference, and badging that "Suggested" would read as the tool
    # proposing a move it is not proposing.
    if choice.settled:
        rec = ('<span class="rec">Where it is now</span>'
               if cand.name == choice.current else "")
    else:
        rec = '<span class="rec">Suggested</span>' if index == 0 else ""
    return f"""
<label class="opt"><input type="radio" name="pick_{choice.slide_index}"
  value="{esc(cand.name)}"{' checked' if checked else ''}>
 <span class="box">
  {_shot(shot, f"Layout {cand.name}", "No preview")}
  {rec}<div class="nm">{esc(cand.name)}</div>{detail}
 </span></label>"""


def _leave_option(choice, checked: bool) -> str:
    """Always last, always available, on every card including a matched one.

    A designer who does not want to decide about this slide yet has to be able
    to say so; the alternative is that they pick something to get past the page.

    It stays on a MATCHED card because that is where it is hardest to say and
    most worth hearing: the file thinks this slide is placed, and a designer who
    opens the card to disagree needs somewhere to say "and none of these fit
    either". Selecting it is what records the refusal - scrolling past the card
    does not, because the pre-selected radio is the layout the slide is already
    on and the route drops a pick that has not moved.
    """
    current = choice.current or "no layout"
    head = ("None of these fit" if choice.settled else "Leave it")
    body = (f"Says this master has no home for this slide even though "
            f"{esc(current)} matched it. It still rebuilds on {esc(current)}, "
            f"and the gap is reported against the master."
            if choice.settled else
            "Says this master has no home for this slide. It rebuilds on "
            "what the file picked, and the gap is reported against the "
            "master.")
    return f"""
<label class="opt"><input type="radio" name="pick_{choice.slide_index}"
  value="{LEAVE}"{' checked' if checked else ''}>
 <span class="box">
  <div class="noshot">{head}</div>
  <div class="nm">Leave on {esc(current)}</div>
  <div class="of">{body}</div>
 </span></label>"""


def _all_layouts(choice, layout_names: list, shortlisted: set) -> str:
    """Everything the shortlist left out, in a select. The shortlist is the
    five worth looking at; the master may have forty, and any of them is a
    legitimate answer."""
    rest = [n for n in layout_names if n not in shortlisted]
    if not rest:
        return ""
    opts = "".join(f'<option value="{esc(n)}">{esc(n)}</option>' for n in rest)
    return f"""
<div class="more"><label>Or another layout from this master:
 <select name="other_{choice.slide_index}">
  <option value="">&mdash;</option>{opts}</select></label></div>"""


def _pickrow(choice, thumbs: dict, slide_shots: dict, options: str) -> str:
    return f"""
<div class="pickrow">
 <div class="thisslide">
  {_shot(slide_shots.get(choice.slide_index),
         f"Slide {choice.slide_index + 1}",
         "The slide could not be rendered on this machine. The layout names "
         "and what each one holds are below.")}
  <div class="cap">The slide as it arrived</div>
 </div>
 <div class="options">{options}</div>
</div>"""


def _card(choice, thumbs: dict, slide_shots: dict, layout_names: list) -> str:
    if choice.settled:
        return _settled_card(choice, thumbs, slide_shots, layout_names)

    unplaced = choice.rule in ("fallback", "none")
    tag = ('<span class="tag unplaced">No match</span>' if unplaced
           else '<span class="tag misfit">Does not fit</span>')
    shortlisted = {c.name for c in choice.candidates}
    # When nothing on the shortlist actually FITS, the honest default is "leave
    # it": that is the true state of this master against this slide, and it is
    # the answer that gets the gap reported. Pre-selecting the best of a bad lot
    # would bury a missing layout under a pick nobody made
    # (qc.layoutpick.apply_picks records the refusal).
    nothing_fits = not any(c.fits for c in choice.candidates)
    picked_default = None if nothing_fits else choice.suggested
    options = "".join(
        _option(choice, cand, i, thumbs, cand.name == picked_default)
        for i, cand in enumerate(choice.candidates))
    options += _leave_option(choice, nothing_fits)

    return f"""
<div class="slidecard">
 <div class="slidehead"><span class="n">Slide {choice.slide_index + 1}</span>
  {tag}</div>
 <p class="why">{esc(_sentence(choice.reason))} This slide holds
  <b>{esc(choice.wants)}</b>.</p>
 {_pickrow(choice, thumbs, slide_shots, options)}
 {_all_layouts(choice, layout_names, shortlisted)}
</div>"""


def _settled_card(choice, thumbs: dict, slide_shots: dict,
                  layout_names: list) -> str:
    """A slide the file placed on a layout its content fits.

    A <details>, so the summary line carries the whole answer - which slide,
    where it is going, what is on it - and the options are one click away for
    the designer who disagrees. Folded rather than omitted: the summary is what
    they asked to be able to read, and the radios behind it are what makes
    reading it useful rather than merely informative.

    The pre-selected radio is the layout it is ALREADY on (Choice.suggested),
    and the route ignores a pick that has not moved, so scrolling past a card
    changes nothing about the run (qc.web.prep_apply_layouts).
    """
    shortlisted = {c.name for c in choice.candidates}
    options = "".join(
        _option(choice, cand, i, thumbs, cand.name == choice.current)
        for i, cand in enumerate(choice.candidates))
    options += _leave_option(choice, False)
    where = esc(choice.current or "no layout")

    return f"""
<details class="slidecard settled">
 <summary><span class="n">Slide {choice.slide_index + 1}</span>
  <span class="tag matched">Matched</span>
  <span class="goes">Going onto <b>{where}</b>. This slide holds
   {esc(choice.wants)}.</span></summary>
 {_pickrow(choice, thumbs, slide_shots, options)}
 {_all_layouts(choice, layout_names, shortlisted)}
</details>"""


def _sentence(text: str) -> str:
    """A reason, as a sentence. Capital at the front and a stop at the end.

    The reasons are written as clauses ("the file names no layout this master
    has", "matched by name to 'X', but the slide sits in 2 columns") because
    qc.layoutpick composes them, and running one straight into the sentence
    after it produced "...and the layout offers 1 column This slide holds"."""
    text = (text or "").strip()
    if not text:
        return ""
    text = text[:1].upper() + text[1:]
    return text if text[-1] in ".!?" else text + "."


def render_layouts(*, deck_name: str, profile_name: str, plan_id: str,
                   choices: list, layout_names: list,
                   slide_shots: dict, layout_thumbs: dict,
                   render_note: str = "", message: str = "") -> str:
    """`choices` are qc.layoutpick.Choice, one per slide of the deck, in deck
    order; `slide_shots` maps slide index to an image URL and `layout_thumbs`
    maps a layout NAME to one. Both may be empty: a host with no renderer still
    gets a working page, in words.

    The matched/asking split is read off `choices` rather than passed in. It
    used to be a `matched` argument computed by the caller, which was a second
    place that had to agree with the list - and the whole defect being fixed
    here was a count and a list disagreeing about the deck."""
    banner = f'<div class="banner warn">{esc(message)}</div>' if message else ""
    note = f'<p class="note">{esc(render_note)}</p>' if render_note else ""
    total = len(choices)
    asking = sum(1 for c in choices if not c.settled)
    settled = total - asking

    cards = "".join(_card(c, layout_thumbs, slide_shots, layout_names)
                    for c in choices)

    # The bar leads on the questions and states the rest, because those are the
    # two facts a designer needs before scrolling: how much of this needs me,
    # and is my whole deck here. It used to say "{n} slides to place" where n
    # was the questions and the deck was not on the page at all, so the number
    # and the list agreed with each other and neither agreed with the deck.
    if asking:
        count = (f'{asking} slide{"s" if asking != 1 else ""} to place')
        rest = (f'<span class="restcount">{settled} already matched, '
                f'listed below and changeable</span>' if settled else "")
    else:
        count = "Every slide matched"
        rest = ('<span class="restcount">All ' + str(total)
                + ' are listed below; open any one to move it.</span>')

    body = f"""
<h1 class="file">{esc(deck_name)}</h1>
<p class="sub">Step 2 of 3 &middot; choosing layouts against
 <b>{esc(profile_name)}</b>. Nothing has been rebuilt yet and the deck you
 uploaded is untouched.</p>
{banner}{note}
<form method="post" action="/prep/{esc(plan_id)}/layouts" id="lf">
 <div class="lbar">
  <span class="count">{count}</span>
  {rest}
  <span class="grow"></span>
  <button class="btn primary" type="submit"
   data-busy="Applying the master"
   data-busysub="Every slide is copied onto its layout through PowerPoint, then
    the content is migrated and the rebuilt deck is audited. This takes a
    minute or two on a long deck.">Apply master to all {total}
   slide{'s' if total != 1 else ''}</button>
 </div>
 {cards}
</form>
"""
    return _shell(f"Layouts: {deck_name}", _CSS + body)


def render_nothing_to_choose(*, deck_name: str, profile_name: str,
                             plan_id: str, slides: int) -> str:
    """The fallback page, for a run that could not work out the choices at all.

    This used to be the "every slide matched" page. It is not any more: a deck
    where everything matched now gets the ordinary page with every slide on it,
    folded (render_layouts). What reaches this is a plan whose choices could not
    be computed - qc.prep.plan degrades to an empty list rather than raising,
    because a deck that rebuilds on the targets plan_assignments picked is still
    a rebuilt deck.

    Still a page and still a press. Skipping straight to the rebuild would mean
    the one run with nothing to show is the one run a designer never got to
    confirm, so it states what it is about to do and asks."""
    body = f"""
<h1 class="file">{esc(deck_name)}</h1>
<p class="sub">Step 2 of 3 &middot; choosing layouts against
 <b>{esc(profile_name)}</b>.</p>
<div class="card">
 <p><b>The per-slide layout review could not be built for this deck.</b> Every
  slide still has a target layout, worked out from its own layout name and
  archetype, and the rebuild will use those. What is missing is the page that
  would let you change them.</p>
 <form method="post" action="/prep/{esc(plan_id)}/layouts"
  data-busy="Applying the master"
  data-busysub="Every slide is copied onto its layout through PowerPoint, then
   the content is migrated and the rebuilt deck is audited.">
  <button class="btn primary" type="submit">Apply master to
   {slides} slide{'s' if slides != 1 else ''}</button>
 </form>
</div>"""
    return _shell(f"Layouts: {deck_name}", _CSS + body)
