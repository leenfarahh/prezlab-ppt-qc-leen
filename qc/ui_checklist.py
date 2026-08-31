"""The colour and type checklist: what this deck is actually made of.

A designer's first question about a client deck is not a finding. It is "what am
I working with" - which navy is the navy, how many of them are there, what is the
body font, and how much of this was typed over the brand by hand. The audit
answers none of that, because none of it is a defect. Three near-identical navies
IS a defect and gets a finding; the palette itself is just a fact, and a designer
needs the fact before they can judge the finding.

So this page states it and stops. Nothing here is a finding, nothing here has a
tick, and nothing here changes the deck (design lead, 26/08/2026: "colors, fonts
checklist written down for the designer to stay in touch"). It is the page a
designer keeps open beside PowerPoint.

WHAT IT SHOWS THAT POWERPOINT WILL NOT. For every colour and every typeface, the
LEVEL it comes from: an explicit hex somebody typed, a reference to a theme slot,
or an inherited value from the master. On screen those are identical. In the file
they are the difference between a deck that survives a rebrand and one that has
to be hunted through by hand, and it is the single most useful thing this tool
knows that a designer cannot see (qc.extract).

Rendering only; route logic lives in qc/web.py.
"""

from .ui import _shell, esc

_CSS = """
<style>
.swat { display:inline-block; width:1.15rem; height:1.15rem; border-radius:4px;
  border:1px solid var(--line); vertical-align:-0.25rem; margin-right:0.4rem; }
.swat.big { width:2.6rem; height:2.6rem; border-radius:6px; margin:0 0.5rem 0 0; }
table.chk { table-layout:auto; }
table.chk td, table.chk th { vertical-align:top; }
table.chk code { font-size:0.85em; }
.hand { background:var(--sand); color:var(--burgundy); border-radius:999px;
  padding:0.1rem 0.5rem; font-size:0.72rem; font-weight:700;
  white-space:nowrap; }
.inh { color:var(--slate-text); font-size:0.78rem; }
.slots { display:flex; flex-wrap:wrap; gap:0.9rem; margin:0.7rem 0 0; }
.slot { text-align:center; font-size:0.72rem; color:var(--slate-text); }
.slot .swat.big { display:block; margin:0 auto 0.25rem; }
.slot code { display:block; color:var(--teal); font-weight:700; }
</style>
"""

# The theme's twelve slots in the order the standard states them, so a designer
# reading two decks side by side is reading the same order twice.
_SLOT_ORDER = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
               "accent4", "accent5", "accent6", "hlink", "folHlink")

_SLOT_LABEL = {"dk1": "text 1", "lt1": "background 1", "dk2": "text 2",
               "lt2": "background 2", "hlink": "link",
               "folHlink": "visited link"}


def _swatch(hexval: str, big: bool = False) -> str:
    safe = "".join(c for c in (hexval or "") if c in "0123456789abcdefABCDEF")
    if len(safe) != 6:
        return ""
    return (f'<span class="swat{" big" if big else ""}" '
            f'style="background:#{safe}" aria-hidden="true"></span>')


def _slides(indices, limit: int = 10) -> str:
    shown = ", ".join(str(i) for i in indices[:limit])
    if len(indices) > limit:
        shown += f" +{len(indices) - limit}"
    return shown or "&mdash;"


def _theme_slots(slots: dict) -> str:
    if not slots:
        return ('<p class="note">This deck states no theme colours, which means '
                'every colour in it is a literal value. There is nothing for a '
                'rebrand to change centrally.</p>')
    cells = []
    for slot in _SLOT_ORDER:
        hexval = slots.get(slot)
        if not hexval:
            continue
        label = _SLOT_LABEL.get(slot, slot)
        cells.append(f'<div class="slot">{_swatch(hexval, big=True)}'
                     f'<code>{esc(slot)}</code>#{esc(hexval)}<br>{esc(label)}'
                     f'</div>')
    return f'<div class="slots">{"".join(cells)}</div>'


def _colour_rows(colours: list) -> str:
    rows = []
    for c in colours:
        written = c.get("written") or {}
        hand = written.get("explicit_rgb") or 0
        themed = sum(n for k, n in written.items() if k != "explicit_rgb")
        how = []
        if hand:
            how.append(f'<span class="hand">{hand} typed by hand</span>')
        if themed:
            slots = ", ".join(c.get("theme_slots") or []) or "the theme"
            how.append(f'<span class="inh">{themed} through {esc(slots)}</span>')
        roles = ", ".join(sorted(c.get("roles") or {}))
        rows.append(
            f'<tr><td>{_swatch(c["hex"])}<code>#{esc(c["hex"])}</code></td>'
            f'<td>{c.get("uses", 0)}</td>'
            f'<td class="note">{esc(roles)}</td>'
            f'<td>{" ".join(how)}</td>'
            f'<td class="note">{_slides(c.get("slides") or [])}</td></tr>')
    return "".join(rows)


def _font_rows(fonts: list) -> str:
    rows = []
    for f in fonts:
        sizes = ", ".join(f"{s['pt']:g}pt" for s in (f.get("sizes") or [])[:6])
        where = (f'<span class="hand">set on the run</span>'
                 if f.get("set_by_hand")
                 else f'<span class="inh">{esc(f.get("from") or "")}</span>')
        cs = f.get("complex_script")
        rows.append(
            f'<tr><td><b>{esc(f["family"])}</b>'
            + (f'<div class="note">Arabic and other complex scripts: '
               f'{esc(cs)}</div>' if cs else "")
            + f'</td><td>{f.get("uses", 0)}</td>'
              f'<td class="note">{esc(sizes) or "&mdash;"}</td>'
              f'<td>{where}</td>'
              f'<td class="note">{_slides(f.get("slides") or [])}</td></tr>')
    return "".join(rows)


def _backgrounds(colours: list) -> str:
    """Which colours act as grounds, and at which level.

    Its own section because it is its own question. "Which colours are explicit
    for backgrounds" was asked separately from "which colours exist" for a
    reason: a hex typed onto one slide's background is a local override that
    beats the master, and it is the reason a deck adopts a master and still
    comes out the wrong colour."""
    grounds = [c for c in colours
               if any(r.startswith("background") for r in (c.get("roles") or {}))]
    if not grounds:
        return ('<p class="note">No slide in this deck states a background of '
                'its own, so every slide takes the master\'s. That is the state '
                'to want.</p>')
    rows = []
    for c in grounds:
        levels = [r for r in sorted(c.get("roles") or {})
                  if r.startswith("background")]
        local = any("(slide)" in r for r in levels)
        rows.append(
            f'<tr><td>{_swatch(c["hex"])}<code>#{esc(c["hex"])}</code></td>'
            f'<td class="note">{esc(", ".join(levels))}</td>'
            f'<td>{"<span class=hand>overrides the master</span>" if local else "<span class=inh>inherited</span>"}</td>'
            f'<td class="note">{_slides(c.get("slides") or [])}</td></tr>')
    return f"""<table class="w3 chk">
  <thead><tr><th>Colour</th><th>Stated by</th><th></th><th>Slides</th></tr></thead>
  <tbody>{"".join(rows)}</tbody></table>"""


def render_checklist(*, deck_name: str, job_id: str, back: str,
                     palette: dict, fonts: dict, tabs: str = "") -> str:
    slots = palette.get("theme_slots") or {}
    colours = palette.get("colours") or []
    distinct = palette.get("distinct_count") or len(colours)
    explicit = palette.get("explicit_count") or 0
    omitted = max(0, distinct - len(colours))

    hand_note = ""
    if explicit:
        hand_note = (
            f'<div class="banner warn"><b>{explicit} of the {distinct} colours '
            f'in this deck were typed in by hand</b> rather than taken from the '
            f'theme. Each one is a value a rebrand will not reach: changing the '
            f'client\'s palette changes the theme, and a typed hex stays exactly '
            f'as it is. That is a fact rather than a fault - it may be entirely '
            f'deliberate - and it is the thing to know before quoting a '
            f'recolour.</div>')

    font_hand = fonts.get("set_by_hand") or 0
    font_note = ""
    if font_hand:
        font_note = _warn_type(font_hand, fonts.get("distinct_families") or 0)

    body = f"""
<span class="kicker">Checklist &middot; nothing here changes the deck</span>
<h1>{esc(deck_name)}</h1>
<p class="sub">What this deck is made of: every colour and every typeface, with
the LEVEL each one comes from. None of this is a finding and none of it has a
tick. It is here to be read beside PowerPoint while you work.</p>
{tabs or f'<div class="actions" style="gap:0.6rem"><a class="btn ghost" href="{esc(back)}">Back to the findings</a></div>'}
{hand_note}
<div class="card">
  <div class="tag">Colour</div>
  <h2 style="margin-top:0">The theme's own palette</h2>
  <p class="sub" style="margin-bottom:0">These twelve slots are what a colour
  can REFER to. A shape pointing at accent1 follows the client's brand through a
  rebrand; a shape holding accent1's hex does not.</p>
  {_theme_slots(slots)}
</div>
<div class="card">
  <div class="tag">Colour</div>
  <h2 style="margin-top:0">Every colour this deck uses</h2>
  <p class="sub">Most-used first. "Typed by hand" means the file states the hex
  itself; the alternative is a reference to a theme slot, which is what survives
  a palette change.</p>
  <table class="w3 chk">
    <thead><tr><th>Colour</th><th>Uses</th><th>Where it appears</th>
    <th>How it is written</th><th>Slides</th></tr></thead>
    <tbody>{_colour_rows(colours)}</tbody>
  </table>
  {f'<p class="note">{omitted} less-used colour(s) are not listed.</p>'
   if omitted else ''}
</div>
<div class="card">
  <div class="tag">Colour</div>
  <h2 style="margin-top:0">What the slides sit on</h2>
  <p class="sub">A slide's own background beats the layout's, which beats the
  master's. A colour listed as overriding the master is the reason a slide can
  adopt every other part of a master and still come out on the wrong ground.</p>
  {_backgrounds(colours)}
</div>
{font_note}
<div class="card">
  <div class="tag">Type</div>
  <h2 style="margin-top:0">Every typeface this deck uses</h2>
  <p class="sub">The theme states {_theme_font_line(fonts)}. Anything else below
  came from somewhere more local, and "set on the run" means somebody typed over
  the brand on that shape.</p>
  <table class="w3 chk">
    <thead><tr><th>Typeface</th><th>Runs</th><th>Sizes</th>
    <th>Where it comes from</th><th>Slides</th></tr></thead>
    <tbody>{_font_rows(fonts.get("fonts") or [])}</tbody>
  </table>
</div>
<p class="note">Read out of the file, not from a render: every value here is
what PowerPoint will resolve when it draws the slide, including the ones the
file inherits rather than states.</p>
{_CSS}"""
    return _shell(f"Checklist: {deck_name}", body)


def _theme_font_line(fonts: dict) -> str:
    theme = fonts.get("theme_fonts") or {}
    major = ((theme.get("major") or {}).get("latin")) or "nothing"
    minor = ((theme.get("minor") or {}).get("latin")) or "nothing"
    return f"<b>{esc(major)}</b> for headings and <b>{esc(minor)}</b> for body"


def _warn_type(hand: int, families: int) -> str:
    return (f'<div class="banner warn"><b>{hand} of the {families} typeface '
            f'entries below are set on the runs themselves.</b> Type set on a '
            f'run ignores the master: restyling the deck onto a new master '
            f'moves the boxes and leaves that type exactly as it was. It is the '
            f'usual reason a rebuilt deck still looks like the old one.</div>')
