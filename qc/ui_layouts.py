"""Step 2: the layouts, approved before anything is rebuilt.

The one page between dropping a deck and getting a rebuilt one. It exists
because applying a master used to be a single press that guessed at the slides
it could not place, and a designer only found out what it had guessed by opening
the result - by which point the guess was already in the file.

WHAT IT SHOWS IS ONLY WHAT IS UNCERTAIN. A slide whose layout name matches a
layout in the master, on a layout its content fits, is not a question, and a
page that asks forty questions to surface four is a page people press Apply on
without reading. The count of what matched is stated, in a line, and that is all
it needs.

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
.matched { color: var(--slate-text); font-size: 0.9rem; margin: 1.4rem 0 0; }
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
    rec = '<span class="rec">Suggested</span>' if index == 0 else ""
    return f"""
<label class="opt"><input type="radio" name="pick_{choice.slide_index}"
  value="{esc(cand.name)}"{' checked' if checked else ''}>
 <span class="box">
  {_shot(shot, f"Layout {cand.name}", "No preview")}
  {rec}<div class="nm">{esc(cand.name)}</div>{detail}
 </span></label>"""


def _leave_option(choice, checked: bool) -> str:
    """Always last, always available. A designer who does not want to decide
    about this slide yet has to be able to say so; the alternative is that they
    pick something to get past the page."""
    current = choice.current or "no layout"
    return f"""
<label class="opt"><input type="radio" name="pick_{choice.slide_index}"
  value="{LEAVE}"{' checked' if checked else ''}>
 <span class="box">
  <div class="noshot">Leave it</div>
  <div class="nm">Leave on {esc(current)}</div>
  <div class="of">Says this master has no home for this slide. It rebuilds on
   what the file picked, and the gap is reported against the master.</div>
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


def _card(choice, thumbs: dict, slide_shots: dict, layout_names: list) -> str:
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
 <p class="why">{esc(_capitalise(choice.reason))} This slide holds
  <b>{esc(choice.wants)}</b>.</p>
 <div class="pickrow">
  <div class="thisslide">
   {_shot(slide_shots.get(choice.slide_index),
          f"Slide {choice.slide_index + 1}",
          "The slide could not be rendered on this machine. The layout names "
          "and what each one holds are below.")}
   <div class="cap">The slide as it arrived</div>
  </div>
  <div class="options">{options}</div>
 </div>
 {_all_layouts(choice, layout_names, shortlisted)}
</div>"""


def _capitalise(text: str) -> str:
    text = (text or "").strip()
    return text[:1].upper() + text[1:] if text else ""


def render_layouts(*, deck_name: str, profile_name: str, plan_id: str,
                   choices: list, layout_names: list, matched: int,
                   slide_shots: dict, layout_thumbs: dict,
                   render_note: str = "", message: str = "") -> str:
    """`choices` are qc.layoutpick.Choice; `slide_shots` maps slide index to an
    image URL and `layout_thumbs` maps a layout NAME to one. Both may be empty:
    a host with no renderer still gets a working page, in words."""
    banner = f'<div class="banner warn">{esc(message)}</div>' if message else ""
    note = f'<p class="note">{esc(render_note)}</p>' if render_note else ""
    n = len(choices)

    cards = "".join(_card(c, layout_thumbs, slide_shots, layout_names)
                    for c in choices)
    matched_line = ""
    if matched:
        matched_line = (f'<p class="matched">The other {matched} slide'
                        f'{"s" if matched != 1 else ""} matched a layout in '
                        f'this master by name or archetype, on a layout their '
                        f'content fits. They are not listed because there is '
                        f'nothing to decide about them.</p>')

    body = f"""
<h1 class="file">{esc(deck_name)}</h1>
<p class="sub">Step 2 of 3 &middot; choosing layouts against
 <b>{esc(profile_name)}</b>. Nothing has been rebuilt yet and the deck you
 uploaded is untouched.</p>
{banner}{note}
<form method="post" action="/prep/{esc(plan_id)}/layouts" id="lf">
 <div class="lbar">
  <span class="count">{n} slide{'s' if n != 1 else ''} to place</span>
  <span class="grow"></span>
  <button class="btn primary" type="submit"
   data-busy="Applying the master"
   data-busysub="Every slide is copied onto its layout through PowerPoint, then
    the content is migrated and the rebuilt deck is audited. This takes a
    minute or two on a long deck.">Apply master to all {matched + n}
   slides</button>
 </div>
 {cards}
</form>
{matched_line}
"""
    return _shell(f"Layouts: {deck_name}", _CSS + body)


def render_nothing_to_choose(*, deck_name: str, profile_name: str,
                             plan_id: str, slides: int) -> str:
    """Every slide matched. Still a page, and still a press.

    Skipping straight to the rebuild would be faster and would also mean the
    one run where the tool was certain is the one run a designer never got to
    confirm. The page states what it is about to do and asks."""
    body = f"""
<h1 class="file">{esc(deck_name)}</h1>
<p class="sub">Step 2 of 3 &middot; choosing layouts against
 <b>{esc(profile_name)}</b>.</p>
<div class="card">
 <p><b>Every slide matched a layout in this master.</b> All {slides} of them
  matched by name or archetype, onto layouts their content fits, so there is
  nothing to choose.</p>
 <form method="post" action="/prep/{esc(plan_id)}/layouts"
  data-busy="Applying the master"
  data-busysub="Every slide is copied onto its layout through PowerPoint, then
   the content is migrated and the rebuilt deck is audited.">
  <button class="btn primary" type="submit">Apply master to
   {slides} slide{'s' if slides != 1 else ''}</button>
 </form>
</div>"""
    return _shell(f"Layouts: {deck_name}", _CSS + body)
