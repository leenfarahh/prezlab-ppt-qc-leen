"""The design QC page: one slide at a time, everything wrong with it, and the
ways out.

The audit report is a list of occurrences grouped by slide, collapsed, with no
picture. That is the right shape for "does this deck match the profile" and the
wrong shape for "is this slide any good", because the second question cannot be
answered without looking at the slide. So this page is built the other way
round: ONE SLIDE PER PAGE, rendered, with every finding against it beside the
render and a box drawn round each one.

Flipped through, not scrolled through (design lead, 23/08/2026). A designer
reviewing a deck works slide by slide and holds one slide in their head at a
time; a single 26-slide scroll is a page you lose your place in. Previous and
Next move one slide, the strip jumps anywhere, and the slide you were on
survives applying a fix.

BOTH SETS OF FINDINGS ARE HERE, AND BOTH ARE ACTIONABLE. The design pass owns
colour, contrast, fit, overlap and the frame; qc.modules owns fonts, margins,
alignment, page furniture. A designer looking at slide 7 wants slide 7's
problems, not the ones that happen to belong to whichever pass found them - so
the audit's own records for the slide are listed underneath WITH THEIR TICK
BOXES. They were read-only for one release, on the argument that the report
already had that control; the objection to that was the right one, because it is
not two controls, it is one control put where the problem is visible (design
lead, 24/08/2026). The engine underneath is still qc.fixer, ticked exactly as
the report ticks it.

Every card is a choice and none has a default. A pre-selected remedy is the tool
deciding and asking for a rubber stamp, and the reason this page exists rather
than another entry in FIXABLE_ISSUES is that these calls are not the tool's to
make: recoloring the text and recoloring the panel are both correct fixes for
grey-on-grey, and only the person looking at the slide knows which the slide
wants. "Leave it" is on every card, is a real answer, and is recorded as one.

AND A DESIGNER MAY HAND THE WHOLE THING OVER. "Let the tool decide" is not that
default coming back in through the side door. A default is the tool answering a
question nobody asked and hoping for a rubber stamp; this is one deliberate
action, taken once, by the person who chose to take it, and every decision it
makes lands in the same list as a hand-picked one with the same Undo beside it.
It declines the calls the checks have already said are not the tool's
(qc.design.UNDECIDABLE_KINDS) and says how many it left.

Rendering only; route logic lives in qc/web.py.
"""

import re
from pathlib import Path

from .ui import MODULE_LABELS, esc, _shell, issue_label
from .ui_chat import chat_css, chat_panel

_KIND_LABEL = {
    "contrast": "Contrast",
    "fit": "Fit",
    "overlap": "Overlap",
    "palette": "Palette",
    "frame": "Outside the frame",
    "error": "Check failed",
}

_KIND_BLURB = {
    "palette": "Colors typed in by hand that disagree with the palette, or with "
               "each other. Theme references are never listed: they move when "
               "the theme moves, which is the point of them.",
    "contrast": "Text measured against what is actually behind it, to WCAG AA "
                "(4.5:1, or 3:1 for 18pt and above). Text over a photograph or "
                "a gradient is not judged: there is no single color to measure "
                "against, and one number for it would be invented.",
    "overlap": "Shapes printing on top of each other. Two graphics overlapping "
               "is composition and is not listed; text another shape hides is "
               "an error, because it survives every proofread.",
    "fit": "Whether the words fit their box, and whether the box fits the card "
           "it sits in. Heights are estimated from the resolved type sizes, not "
           "from a render.",
    "frame": "Text outside the presentation space the master states, in no "
             "placeholder. Listed, never removed on its own: a page number and "
             "a stranded eyebrow look identical from the file.",
}

_ORDER = ("contrast", "fit", "overlap", "palette", "frame", "error")

_STYLE = """
<style>
.swatch { display:inline-block; width:0.85rem; height:0.85rem; border-radius:3px;
  border:1px solid rgba(0,37,40,0.25); vertical-align:-2px; margin-right:0.3rem }
.ev { display:flex; flex-wrap:wrap; gap:0.4rem 1.1rem; margin:0.55rem 0 0.2rem;
  font-size:0.82rem; color:var(--slate-text) }
.ev code { font-family:ui-monospace,Menlo,Consolas,monospace; font-size:0.82rem }
.dsplit { display:grid; grid-template-columns:minmax(0,1.15fr) minmax(0,1fr);
  gap:1.1rem; align-items:start }
@media (max-width:1100px) { .dsplit { grid-template-columns:1fr } }
/* ONE sticky bar on this page, and the slide parked below it.
   .actionbar is sticky at top:0 for the report, and this page has three of
   them - the tabs, the pager above the split and the pager below it. All three
   pinned to the same 3.4rem, so they painted over each other AND over the top
   of the sticky slide, which is the strip of the render that went missing as
   soon as you scrolled (design lead, 24/08/2026). The tabs and the closing
   pager scroll away like ordinary content; the pager that carries Previous and
   Next stays, because that is the one a designer reaches for mid-slide.
   --dbar-h is its measured height (see _STICKY_JS); the fallback is only for
   the instant before that runs, and for no-JS. */
.dtabs, .dfoot { position:static }
.dbar { z-index:12 }
/* The pager pins its border box 0.4rem down (its own margin), so clearing
   --dbar-h + 1rem leaves about 0.6rem of air between the bar and the render
   instead of the overlap it had. */
.dshot { position:sticky; top:calc(var(--dbar-h, 4.4rem) + 1rem);
  /* A render taller than what is left of the window scrolls inside its own box
     rather than running off the bottom of it. Scroll chaining is left alone:
     when the slide does fit, this box never scrolls, and trapping the wheel
     over the picture would be worse than the problem. */
  max-height:calc(100vh - var(--dbar-h, 4.4rem) - 2rem); overflow:auto }
@media (max-width:1100px) {
  /* one column: the slide is above its findings, not beside them, so pinning
     it would cover them */
  .dshot { position:static; max-height:none; overflow:visible }
}
.dframe { position:relative; border:1px solid var(--line-soft); border-radius:10px;
  overflow:hidden; background:#fff }
.dframe img { display:block; width:100%; height:auto }
.dframe .hit { position:absolute; border:2px solid var(--orange); border-radius:3px;
  box-shadow:0 0 0 9999px rgba(0,0,0,0); pointer-events:none }
.dframe .hit.error { border-color:var(--burgundy) }
.dframe .hit.info { border-color:var(--slate) }
.dframe .hit b { position:absolute; top:-1px; left:-1px; background:var(--orange);
  color:var(--teal); font-size:10px; line-height:1; padding:2px 4px;
  border-radius:2px 0 3px 0 }
.dframe .hit.error b { background:var(--burgundy); color:var(--offwhite) }
.dframe .hit.info b { background:var(--slate); color:var(--offwhite) }
.dcard { margin-bottom:0.8rem }
.dcard h3 { margin:0.25rem 0 0.3rem; font-size:1rem }
.dcard fieldset { border:1px solid var(--line-soft); border-radius:10px;
  padding:0.4rem 0.85rem 0.7rem; margin:0.7rem 0 0 }
/* what the option would produce, beside the option. The contrast chip carries
   letters because a pair of swatches does not answer "can this be read". */
.rprev { flex:0 0 auto; align-self:center; display:inline-flex;
  align-items:center; justify-content:center; min-width:2.2rem; height:1.6rem;
  padding:0 0.35rem; border-radius:4px; border:1px solid rgba(0,37,40,0.25);
  font-size:0.82rem; font-weight:700; line-height:1 }
.rchips { flex:0 0 auto; align-self:center; display:inline-flex;
  align-items:center; gap:0.25rem }
.rchips i { display:block; width:1.15rem; height:1.15rem; border-radius:3px;
  border:1px solid rgba(0,37,40,0.25) }
.rchips u { font-size:0.7rem; color:var(--slate); text-decoration:none }
.dpin { display:inline-block; min-width:1.25rem; text-align:center;
  border-radius:4px; font-size:11px; font-weight:700; padding:1px 4px;
  background:var(--orange); color:var(--teal); margin-right:0.35rem }
.dpin.error { background:var(--burgundy); color:var(--offwhite) }
.dpin.info { background:var(--slate); color:var(--offwhite) }
.strip { display:flex; flex-wrap:wrap; gap:0.28rem; margin:0.5rem 0 0.9rem }
.strip a, .strip span.here { display:inline-flex; flex-direction:column;
  align-items:center; gap:2px; min-width:1.9rem; padding:0.24rem 0.3rem;
  border:1px solid var(--line-soft); border-radius:7px; font-size:0.76rem;
  font-weight:600; text-decoration:none; color:var(--teal); background:#fff }
.strip span.here { border-color:var(--teal); background:var(--teal);
  color:var(--offwhite) }
.strip .sd { display:flex; gap:2px; height:5px }
.strip .sd i { width:5px; height:5px; border-radius:50%; display:block }
.done { display:flex; gap:0.7rem; align-items:baseline; padding:0.55rem 0;
  border-top:1px solid var(--line-soft) }
.done:first-of-type { border-top:0 }
.done .grow { flex:1 }
.auditrow { display:flex; gap:0.6rem; align-items:baseline; padding:0.4rem 0;
  border-top:1px solid var(--line-soft); font-size:0.86rem }
.auditrow:has(input:checked) { background:var(--hover) }
.auditrow .afix { flex:0 0 6.5rem; text-align:left }
.auditrow .nofixwhy { color: var(--slate-text); font-size: 0.82rem;
  margin-top: 0.2rem; font-style: italic; }
.nofix { font-size:0.74rem; opacity:0.7 }
.autocard { border-color:var(--teal) }
.autohold { display:flex; gap:0.5rem; align-items:baseline; margin:0.6rem 0 0;
  font-size:0.84rem; color:var(--slate-text) }
/* answered on the card, settled on the render: the box for a finding that has
   been decided comes off the picture at once, so the slide stops advertising
   problems the designer has already dealt with. */
.dcard.answered { border-color:var(--teal); background:var(--hover) }
.dframe .hit { transition:opacity 0.18s ease, border-color 0.18s ease }
.dframe .hit.settled { opacity:0.12 }
@media (prefers-reduced-motion: reduce) { .dframe .hit { transition:none } }
.hidden { display:none !important }
</style>"""

_APPLY_JS = """
<script>
/* Answering a card changes the picture, and the round trip is not instant.
   Re-auditing the deck and re-rendering the slide takes seconds (PowerPoint
   startup alone is most of it), and for that whole time the page used to sit
   there showing the boxes for the very problems being fixed - so a designer
   could not tell whether the click had registered (design lead, 24/08/2026).
   Two things happen the moment a remedy is picked, both of them free:
     - the box comes off the render, because that finding is being dealt with;
     - the card says so.
   "Leave it" takes the box off too. The finding is answered either way; what
   differs is what happens to the deck, not whether it is still an open
   question on this slide. */
(function () {
  const frame = document.querySelector('.dframe');
  const box = id => frame && id
    ? frame.querySelector('.hit[data-finding="' + id + '"]') : null;

  function mark(card, on) {
    const id = card.getAttribute('data-finding');
    card.classList.toggle('answered', on);
    const hit = box(id);
    if (hit) hit.classList.toggle('settled', on);
  }

  document.querySelectorAll('.dcard[data-finding]').forEach(card => {
    card.querySelectorAll('input[type="radio"]').forEach(radio => {
      radio.addEventListener('change', () => mark(card, radio.checked));
    });
  });

  /* the wait itself, said out loud */
  const form = document.getElementById('dform');
  if (form) form.addEventListener('submit', () => {
    const n = form.querySelectorAll('input[type="radio"]:checked').length;
    if (window.showBusy) showBusy(
      'Applying ' + n + ' choice' + (n === 1 ? '' : 's'),
      'Changing the deck, re-auditing it and re-rendering this slide. The '
      + 'boxes that have gone are the ones being dealt with.');
  });
  document.querySelectorAll('form[action$="/auto"]').forEach(f =>
    f.addEventListener('submit', ev => {
      const scope = (ev.submitter && ev.submitter.value) === 'deck'
        ? 'the whole deck' : 'this slide';
      if (window.showBusy) showBusy('Deciding ' + scope,
        'Applying the audit fixes first, then answering each design card. '
        + 'Everything it decides gets an Undo.');
    }));
  document.querySelectorAll('form[action$="/fix"]').forEach(f =>
    f.addEventListener('submit', () => {
      const n = f.querySelectorAll('input[name="record_ids"]:checked').length;
      if (window.showBusy) showBusy(
        'Applying ' + n + (n === 1 ? ' fix' : ' fixes'),
        'Then re-auditing the deck to check what the fix actually did.');
    }));
  document.querySelectorAll('form[action$="/undo"]').forEach(f =>
    f.addEventListener('submit', () => {
      if (window.showBusy) showBusy('Putting it back',
        'Replaying the state stored before the change, then re-rendering.');
    }));
})();
</script>"""

_STICKY_JS = """
<script>
/* How far down the sticky slide has to start: the pager's real height, not a
   guess at it. The bar wraps on a narrow window and grows a line, and a
   hard-coded offset then cuts the top off the render again - which is the bug
   this is here for. Measured on load and on every resize of the bar itself. */
(function () {
  const bar = document.querySelector('.actionbar.dbar');
  if (!bar) return;
  const set = () => document.documentElement.style.setProperty(
    '--dbar-h', bar.getBoundingClientRect().height.toFixed(1) + 'px');
  set();
  if (window.ResizeObserver) new ResizeObserver(set).observe(bar);
  window.addEventListener('resize', set);
})();
</script>"""

_FILTER_JS = """
<script>
(function () {
  const chips = document.querySelectorAll('.fchip[data-f]');
  const cards = document.querySelectorAll('[data-sev]');
  function apply(key) {
    cards.forEach(el => {
      const sev = el.getAttribute('data-sev');
      const kind = el.getAttribute('data-kind') || '';
      let show = true;
      if (key === 'error' || key === 'warning' || key === 'info') show = sev === key;
      else if (key !== 'all') show = kind === key;
      el.classList.toggle('hidden', !show);
    });
    document.querySelectorAll('.dframe .hit').forEach(box => {
      const sev = box.getAttribute('data-sev');
      const kind = box.getAttribute('data-kind') || '';
      let show = true;
      if (key === 'error' || key === 'warning' || key === 'info') show = sev === key;
      else if (key !== 'all') show = kind === key;
      box.classList.toggle('hidden', !show);
    });
  }
  chips.forEach(c => c.addEventListener('click', () => {
    chips.forEach(x => x.setAttribute('aria-pressed', 'false'));
    c.setAttribute('aria-pressed', 'true');
    apply(c.getAttribute('data-f'));
  }));
})();
</script>"""


# ------------------------------------------------------------------ pieces


def _sev_pill(severity: str) -> str:
    return f'<span class="pill {esc(severity)}">{esc(severity)}</span>'


def _swatch(hexval: str) -> str:
    return (f'<span class="swatch" style="background:#{esc(hexval)}"></span>'
            f'<code>#{esc(hexval)}</code>')


def _slides_note(slides: list) -> str:
    if not slides:
        return "the whole deck"
    shown = [str(i + 1) for i in slides[:8]]
    more = f" and {len(slides) - 8} more" if len(slides) > 8 else ""
    if len(slides) == 1:
        return f"slide {shown[0]}"
    return f"slides {', '.join(shown)}{more}"


def _evidence(finding) -> str:
    """The numbers behind the headline, because a designer overrules a check by
    looking at the measurement rather than at the adjective."""
    ev = finding.evidence or {}
    bits = []
    if finding.kind == "palette":
        if ev.get("hex"):
            bits.append(f'<span>Used: {_swatch(ev["hex"])}</span>')
        if ev.get("anchor"):
            name = ev.get("anchor_name") or "the deck's other spelling"
            bits.append(f'<span>{esc(name)}: {_swatch(ev["anchor"])}</span>')
        if ev.get("delta_e") is not None:
            bits.append(f'<span>Difference: {ev["delta_e"]} deltaE</span>')
        if ev.get("surfaces"):
            bits.append(f'<span>On: {esc(", ".join(ev["surfaces"]))}</span>')
    elif finding.kind == "contrast":
        bits.append(f'<span>Text {_swatch(ev.get("text", "000000"))}</span>')
        bits.append(f'<span>On {_swatch(ev.get("ground", "FFFFFF"))}</span>')
        bits.append(f'<span>Measured: <b>{ev.get("ratio")}:1</b> against '
                    f'{ev.get("need")}:1</span>')
        if ev.get("size_pt"):
            bits.append(f'<span>Type: {ev["size_pt"]:.0f}pt</span>')
    elif finding.kind == "fit":
        if ev.get("over_in") is not None:
            bits.append(f'<span>Needs {ev["needs_in"]}in in a '
                        f'{ev["box_in"]}in box</span>')
            bits.append(f'<span>Outside: <b>{ev["over_in"]}in</b></span>')
        if ev.get("escape_in") is not None:
            bits.append(f'<span>Crosses the edge by '
                        f'<b>{ev["escape_in"]}in</b> ({esc(ev["side"])})</span>')
    elif finding.kind == "overlap":
        if ev.get("cover") is not None:
            bits.append(f'<span>Covered: {ev["cover"] * 100:.0f}%</span>')
        if ev.get("share") is not None:
            bits.append(f'<span>Shared area: {ev["share"] * 100:.0f}%</span>')
    elif finding.kind == "frame":
        bits.append(f'<span>At {ev.get("left_in")}in, {ev.get("top_in")}in</span>')
        bits.append(f'<span>Appears {ev.get("places", 0)}&times;</span>')
    if not bits:
        return ""
    return f'<div class="ev">{"".join(bits)}</div>'


_HEX6 = re.compile(r"^[0-9A-Fa-f]{6}$")


def _remedy_preview(finding, option) -> str:
    """What the slide gets if this option is picked, shown as the colour rather
    than spelled as a hex.

    "#464646" is not a colour to the person reading it, it is six characters
    they have to imagine, and this whole page exists because the decision is
    made by LOOKING. The evidence row above the options already shows what is
    there now; without this, the row proposing the change was the only one on
    the card with no colour on it (design lead, 24/08/2026).

    A contrast remedy shows the PAIR, not a swatch. "Is this readable" is never
    a question about one colour: the remedy moves either the text or the ground
    and leaves the other where it is, so the honest preview is the two of them
    together with letters in it. A palette remedy shows the swap, old beside
    new, because what is being judged there is how far the colour travels.
    """
    params = option.params or {}
    ev = finding.evidence or {}
    new = params.get("hex")
    if option.op == "set_theme_color":
        # A theme reference has no hex of its own; what it resolves to today is
        # the anchor the finding matched, and "same colour on screen today" is
        # exactly what the note promises.
        new = ev.get("anchor")
    if not new or not _HEX6.match(str(new)):
        return ""
    new = str(new)

    if finding.kind == "contrast":
        ground, text = ev.get("ground", "FFFFFF"), ev.get("text", "000000")
        if option.remedy_id == "ground":
            ground = new
        else:
            text = new
        if not (_HEX6.match(str(ground)) and _HEX6.match(str(text))):
            return ""
        return (f'<span class="rprev" aria-hidden="true" '
                f'style="background:#{esc(ground)};color:#{esc(text)}">Aa</span>')

    was = str(ev.get("hex") or "")
    if _HEX6.match(was) and was.upper() != new.upper():
        return (f'<span class="rchips" aria-hidden="true">'
                f'<i style="background:#{esc(was)}"></i><u>&rarr;</u>'
                f'<i style="background:#{esc(new)}"></i></span>')
    return (f'<span class="rchips" aria-hidden="true">'
            f'<i style="background:#{esc(new)}"></i></span>')


def _options(finding) -> str:
    """The ways out, each with the colour it would produce.

    The preview is aria-hidden: the label beside it already carries the palette
    name and the hex, so a screen reader gets the answer in words and does not
    also get an unlabelled box read out as nothing."""
    if not finding.options:
        return ""
    rows = "".join(
        f'<label class="radio-card"><input type="radio" '
        f'name="pick_{esc(finding.finding_id)}" value="{esc(o.remedy_id)}">'
        f'{_remedy_preview(finding, o)}'
        f'<span><b>{esc(o.label)}</b><small>{esc(o.note)}</small></span></label>'
        for o in finding.options)
    return f'<fieldset><legend>Pick one</legend>{rows}</fieldset>'


def _finding_card(finding, pin: int | None = None, scope: str = "") -> str:
    places = (finding.evidence or {}).get("places") or 0
    spread = (f' &middot; {places} places' if places > 1 else "")
    badge = (f'<span class="dpin {esc(finding.severity)}">{pin}</span>'
             if pin else "")
    return f"""
<div class="card dcard" data-sev="{esc(finding.severity)}"
     data-kind="{esc(finding.kind)}"
     data-finding="{esc(finding.finding_id)}">
  <div class="difflabels">{_sev_pill(finding.severity)}
    <span class="note">{esc(_KIND_LABEL.get(finding.kind, finding.kind))}
    &middot; {esc(scope or _slides_note(finding.slides))}{spread}</span></div>
  <h3>{badge}{esc(finding.headline)}</h3>
  <p class="note">{esc(finding.detail)}</p>
  {_evidence(finding)}
  {_options(finding)}
</div>"""


def _audit_rows(job_id: str, records: list, current: int = 0,
                can_fix: bool = False, promoted: set | None = None) -> str:
    """The audit's own findings for this slide, each with its fix beside it.

    The SAME tick as the audit report, on purpose and in every detail: the same
    qc.fixer.is_fixable decides whether a row gets one, the same
    qc.fixer.tick_reason decides whether it may be pre-ticked, and the same
    engine applies it. Two pages showing one record in two different states is
    the failure this shares its code to avoid.

    Rows with no tick say why in the last column, because "no box here" and
    "nothing wrong here" look identical otherwise."""
    from .fixer import is_fixable, no_fix_reason, tick_reason

    def _cap(text: str) -> str:
        return text[:1].upper() + text[1:] if text else text

    if not records:
        return ""
    promoted = promoted or set()
    rows, fixable, preticked = [], 0, 0
    # Vision first here too, for the reason the audit report does it: what the
    # model noticed about this slide is the answer to the question the designer
    # opened this page to ask, and a font family is not. Severity orders within
    # each half (qc.records.FindingRecord.source).
    for r in sorted(records, key=lambda r: (r.get("source") != "vision",
                                            r["severity"])):
        module = MODULE_LABELS.get(r["module"], r["module"])
        if r["action"] == "changed":
            fix_cell = '<span class="pill changed">fixed</span>'
        elif not can_fix:
            # The deck is gone from memory, not the fix. Saying "no automatic
            # fix" here would be a lie the banner above then contradicts.
            fix_cell = '<span class="note nofix">deck not in memory</span>'
        elif is_fixable(r):
            fixable += 1
            # Pre-ticked = the tool is confident it is wrong AND has a safe fix:
            # deterministic changes, triage-promoted types, and errors (by
            # taxonomy, error means confidently wrong). Arabic font
            # substitutions and whole-slide moves are never pre-selected - the
            # tick is the designer's approval, and tick_reason says so.
            hold = tick_reason(r)
            pre = (hold is None
                   and (r["confidence"] == "deterministic"
                        or r["issue_type"] in promoted
                        or r["severity"] == "error"))
            preticked += 1 if pre else 0
            why = (hold if hold
                   else "deterministic fix" if r["confidence"] == "deterministic"
                   else "confidently wrong (error) with a safe fix"
                   if r["severity"] == "error"
                   else "validated by designer triage" if pre
                   else "suggestion: ticking it is your approval")
            fix_cell = (f'<input type="checkbox" name="record_ids" '
                        f'value="{esc(r["record_id"])}"'
                        f'{" checked" if pre else ""} title="{esc(why)}" '
                        f'aria-label="Apply this fix">')
        elif r["arabic_flag"]:
            fix_cell = '<span class="note nofix">Arabic, by hand</span>'
        else:
            fix_cell = '<span class="note nofix">no automatic fix</span>'

        # WHY there is no tick, under the finding. "No automatic fix" on its own
        # reads as an unfinished tool, and on a slide where three rows say it
        # that reading is unavoidable - where "the breach is measured, the
        # correction is not" is a designer telling another designer something
        # true (qc.fixer.no_fix_reason, 31/08/2026).
        why_none = "" if can_fix and is_fixable(r) else no_fix_reason(r)
        no_fix = (f'<div class="note nofixwhy">{esc(_cap(why_none))}.'
                  f'</div>' if why_none else "")
        rows.append(
            f'<label class="auditrow" data-sev="{esc(r["severity"])}" '
            f'data-kind="{esc(r["module"])}">'
            f'<span class="afix">{fix_cell}</span>{_sev_pill(r["severity"])}'
            f'<span class="grow" style="flex:1"><b>{esc(issue_label(r["issue_type"]))}</b>'
            f'<div class="note">{esc(r["message"])}</div>{no_fix}</span>'
            f'<span class="note" style="white-space:nowrap">{esc(module)}</span>'
            f'</label>')

    if fixable:
        foot = (f'<div class="actions"><button class="btn primary" '
                f'type="submit">Apply the ticked fixes</button>'
                f'<span class="note" style="margin-left:0.7rem">'
                f'{fixable} of these can be fixed here, {preticked} ticked for '
                f'you. The deck is re-audited afterwards and you come back to '
                f'this slide.</span></div>')
    elif not can_fix:
        foot = ('<p class="note">Nothing can be applied from here while the deck '
                'is out of memory. Re-upload it and the ticks come back.</p>')
    else:
        foot = ('<p class="note">None of these has an automatic fix: they are '
                'either Arabic content, which is never re-typed without a '
                'designer, or checks that flag without computing a target.</p>')

    body = f"""
<div class="card">
  <div class="tag">From the audit</div>
  <h3 style="margin:0.2rem 0 0.1rem">Also on this slide</h3>
  <p class="note">Fonts, margins, alignment and page furniture, found by the
  audit's own checks. Listed here so one slide's problems are in one place, and
  fixable here for the same reason: the tick is the audit report's own, applied
  by the same engine, so a row ticked here and a row ticked there do the same
  thing.</p>
  {''.join(rows)}
  {foot}
</div>"""
    if not fixable:
        return body
    # `n` carries the slide back, so applying a fix ticked on slide 7 answers on
    # slide 7 rather than on slide 1.
    return (f'<form method="post" action="/design/{esc(job_id)}/fix">'
            f'<input type="hidden" name="n" value="{int(current)}">'
            f'{body}</form>')


def _auto_card(job_id: str, current: int, view: str, plan: dict,
               can_fix: bool = True) -> str:
    """Hand the decisions to the tool, in one deliberate action.

    Every number on it comes from the same function the route uses to select the
    work (qc.web._auto_targets), so the count on the button is the count that
    happens - a button that says 9 and does 14 is the end of a designer trusting
    this page.

    Two scopes and no third. "This slide" is the designer who has looked at the
    slide and agrees with what the tool is proposing there; "the whole deck" is
    the designer who wants a first pass to correct rather than a blank one to
    build. A middle option (this section, these five slides) would be a fourth
    way of saying the same thing and a fourth thing to explain.
    """
    if not plan or not can_fix:
        return ""
    slide, deck = plan.get("slide") or {}, plan.get("deck") or {}
    here = deck if view == "deck" else slide
    total_deck = deck.get("fixes", 0) + deck.get("picks", 0)
    total_slide = slide.get("fixes", 0) + slide.get("picks", 0)
    if not (total_deck or deck.get("held") or deck.get("left")):
        return ""

    def button(key, label, n, primary):
        cls = "primary" if primary else "ghost"
        dis = " disabled" if not n else ""
        return (f'<button class="btn {cls}" type="submit" name="scope" '
                f'value="{key}"{dis}>{label} ({n})</button>')

    buttons = ""
    if view != "deck":
        buttons += button("slide", f"Decide slide {current + 1}",
                          total_slide, True)
    buttons += button("deck", "Decide the whole deck", total_deck,
                      view == "deck")

    held = here.get("held", 0)
    hold_box = ""
    if held:
        hold_box = (
            f'<label class="autohold"><input type="checkbox" '
            f'name="include_holds" value="1"> Include the {held} fix'
            f'{"es" if held != 1 else ""} that ask for your explicit approval '
            f'&mdash; Arabic font substitution, which changes how the script '
            f'shapes, and moves that shift every element on a slide together. '
            f'Left out unless you say so.</label>')

    left = here.get("left", 0)
    left_note = ""
    if left:
        # The reasons come from the checks themselves (qc.design.auto_skip_reason)
        # rather than being restated here, so a check that changes its mind about
        # what it can answer does not leave this paragraph lying.
        why = here.get("reasons") or []
        left_note = (f'<p class="note">{left} finding'
                     f'{"s" if left != 1 else ""} here will be left for you '
                     f'either way'
                     + (": " + esc("; and ".join(why)) if why else "") + '.</p>')

    return f"""
<form method="post" action="/design/{esc(job_id)}/auto" class="card autocard">
<input type="hidden" name="n" value="{int(current)}">
  <div class="tag">Hand it over</div>
  <h3 style="margin:0.2rem 0 0.1rem">Let the tool decide</h3>
  <p class="note">Applies the audit's own fixes, then takes the first option on
  each design card &mdash; the one each check already puts first as its
  recommendation, and explains in the note beside it. The audit fixes go first
  so the design calls are made about the deck as they leave it, not about the
  deck as it stands now.</p>
  <p class="note">Nothing here is one-way. Every decision it makes is listed
  under &ldquo;what you have decided&rdquo; with an Undo beside it, exactly as
  if you had picked it yourself, so this is a first pass you correct rather than
  a result you have to accept.</p>
  {left_note}
  {hold_box}
  <div class="actions">{buttons}</div>
</form>"""


def _filters(counts: dict, kinds: list) -> str:
    def chip(key, label, n, pressed=False):
        dis = " disabled" if not n else ""
        return (f'<button type="button" class="fchip" data-f="{esc(key)}" '
                f'aria-pressed="{"true" if pressed else "false"}"{dis}>'
                f'{esc(label)} {n}</button>')

    total = sum(counts.get(k, 0) for k in ("error", "warning", "info"))
    bits = [chip("all", "Everything", total, True),
            chip("error", "Errors", counts.get("error", 0)),
            chip("warning", "Warnings", counts.get("warning", 0)),
            chip("info", "Info", counts.get("info", 0))]
    for key, n in kinds:
        bits.append(chip(key, _KIND_LABEL.get(key, MODULE_LABELS.get(key, key)), n))
    return f'<div class="chips no-print">{"".join(bits)}</div>'


def _strip(job_id: str, current: int, per_slide: dict, total: int) -> str:
    """Every slide as a chip, with dots for what is on it, so a designer can jump
    to the slide they remember rather than paging to it."""
    if total <= 1:
        return ""
    window = 24
    start = max(0, min(current - window // 2, total - window))
    end = min(total, start + window)
    bits = []
    if start > 0:
        bits.append(f'<a href="/design/{esc(job_id)}?n=0" title="Slide 1">1 &hellip;</a>')
    for i in range(start, end):
        sev = per_slide.get(i) or {}
        dots = "".join(
            f'<i style="background:var({c})"></i>'
            for key, c in (("error", "--burgundy"), ("warning", "--orange"),
                           ("info", "--slate")) if sev.get(key))
        inner = f'{i + 1}<span class="sd">{dots}</span>'
        if i == current:
            bits.append(f'<span class="here">{inner}</span>')
        else:
            bits.append(f'<a href="/design/{esc(job_id)}?n={i}">{inner}</a>')
    if end < total:
        bits.append(f'<a href="/design/{esc(job_id)}?n={total - 1}" '
                    f'title="Slide {total}">&hellip; {total}</a>')
    return f'<div class="strip no-print">{"".join(bits)}</div>'


def _pager(job_id: str, current: int, total: int, counted: int,
           sticky: bool = True, also_here: int = 0) -> str:
    """Where you are and how to move. The one above the split stays put while
    the slide is read; the one below it is the end of the list and scrolls away
    (two pinned copies of the same bar is two answers to "where am I")."""
    def link(target, label, on):
        if not on:
            return f'<span class="btn ghost" aria-disabled="true">{label}</span>'
        return (f'<a class="btn ghost" href="/design/{esc(job_id)}?n={target}">'
                f'{label}</a>')

    # "nothing to decide here" is only true when nothing DECK-WIDE covers this
    # slide either. It read as a clean bill of health on slides carrying six
    # deck-wide decisions, which is the same overclaim _nothing_here fixes.
    if counted:
        tail = f" &middot; {counted} to decide here"
    elif also_here:
        tail = (f" &middot; nothing on its own, {also_here} deck-wide "
                f"decision{'s' if also_here != 1 else ''} cover it")
    else:
        tail = " &middot; nothing to decide here"
    where = (f'<span class="note"><b>Slide {current + 1}</b> of {total}'
             + tail + '</span>')
    return (f'<div class="actionbar no-print {"dbar" if sticky else "dfoot"}">'
            f'{where}<span class="grow"></span>'
            f'{link(current - 1, "&larr; Previous", current > 0)}'
            f'{link(current + 1, "Next &rarr;", current + 1 < total)}</div>')


def _shot(job_id: str, index: int, rects: list, error: str | None,
          tag: str = "") -> str:
    if error:
        return (f'<div class="dframe" style="padding:1.1rem">'
                f'<p class="note">No render: {esc(error)} Everything on the '
                f'right is read from the deck itself, not from the picture.</p>'
                f'</div>')
    # data-finding ties a box to its card, so picking a remedy can take the box
    # off the picture the moment it is picked rather than after the round trip.
    boxes = "".join(
        f'<div class="hit {esc(r["severity"])}" data-sev="{esc(r["severity"])}" '
        f'data-kind="{esc(r["kind"])}" '
        f'data-finding="{esc(r.get("finding_id") or "")}" '
        f'title="{esc(r["label"])}" '
        f'style="left:{r["x"] * 100:.2f}%;top:{r["y"] * 100:.2f}%;'
        f'width:{r["w"] * 100:.2f}%;height:{r["h"] * 100:.2f}%">'
        + (f'<b>{r["pin"]}</b>' if r.get("pin") else "") + '</div>'
        for r in rects)
    # ?v= is the deck's own digest (qc.web._render_tag). Without it the browser
    # answers this URL from its own cache and the picture beside a row marked
    # "applied" is the slide as it was before the decision.
    return (f'<div class="dframe"><img src="/design-img/{esc(job_id)}/{index}.png'
            f'{f"?v={esc(tag)}" if tag else ""}"'
            f' alt="Slide {index + 1}" loading="lazy">{boxes}</div>')


# ------------------------------------------------------------------ the views


def _nothing_here(job_id: str, current: int, deck_findings: list) -> str:
    """The empty state, and it must not overclaim.

    "Nothing to decide on this slide. No color conflict, no unreadable text,
    nothing overflowing its box, nothing hidden" was printed whenever this
    slide had no finding OF ITS OWN - and a finding that spans slides is
    deliberately shown on the deck-wide tab instead (render_design splits on
    `f.slides == [current]`). So a slide covered by six deck-wide decisions was
    told it was clean, while the strip above it drew dots for those same six
    (_design_severity_map counts every finding on every slide it touches). The
    page contradicted itself, and the half a designer reads is the sentence
    (design lead, 27/08/2026: "how is there nothing to decide, the contents are
    all over the place").

    The split itself is right and stays: a colour spelled two ways across forty
    shapes is one decision about the deck, and pretending it can be taken on
    slide 2 alone would apply it to the other thirty-nine without saying so.
    What was wrong was calling that "nothing".
    """
    here = [f for f in deck_findings if current in (f.slides or [])]
    if here:
        kinds: dict[str, int] = {}
        for f in here:
            label = _KIND_LABEL.get(f.kind, f.kind)
            kinds[label] = kinds.get(label, 0) + 1
        what = ", ".join(f"{n} {label.lower()}"
                         for label, n in sorted(kinds.items(),
                                                key=lambda kv: -kv[1]))
        return (
            f'<div class="card"><div class="tag">Decided deck-wide</div>'
            f'<h3 style="margin:0.4rem 0">Nothing on this slide <em>alone</em>'
            f'</h3><p class="note">This slide is covered by '
            f'<b>{len(here)}</b> decision{"s" if len(here) != 1 else ""} that '
            f'span several slides ({what}). They are taken once, on the '
            f'deck-wide tab, because taking one here would silently apply it '
            f'to every other slide it touches.</p>'
            f'<div class="actions" style="margin-top:0.6rem">'
            f'<a class="btn primary" href="/design/{esc(job_id)}?view=deck">'
            f'See the {len(here)} that cover this slide</a></div></div>')
    # Genuinely nothing from this pass. The scope of the claim is spelled out
    # rather than left implied: text set over a picture is read as composition
    # and never flagged (qc.design._overlap_findings - "text drawn over a
    # graphic is the normal case"), so a visibly busy slide can be quiet here
    # and a designer should know that is a rule and not an oversight.
    return ('<div class="card clean"><div class="mark">&#10003;</div>'
            '<h3 style="margin:0.4rem 0">Nothing to decide on this slide'
            '</h3><p class="note">No colour conflict, no unreadable text, '
            'nothing overflowing its box, nothing hidden. Text set over a '
            'picture is read as composition and is not flagged here, so a busy '
            'slide can still come out quiet.</p></div>')


def _tabs(job_id: str, view: str, deck_n: int, current: int,
          back: tuple | None = None) -> str:
    def tab(key, label, href):
        on = " primary" if view == key else " ghost"
        return f'<a class="btn{on}" href="{href}">{label}</a>'

    # Where "back" goes depends on how the designer got here. A deck that was
    # prepared came from one page carrying the coverage, the gaps and these
    # findings together, and sending them to the audit report instead would
    # strand the half of the answer that is about the master.
    back_href, back_label = back or (f"/audit/{job_id}", "Back to the audit")

    # Unpinned on the slide view, where the pager below it is the bar that
    # stays; the deck view has no pager, so there it keeps .actionbar's own
    # stickiness rather than losing the only fixed thing on the page.
    cls = "actionbar no-print" + ("" if view == "deck" else " dtabs")
    return (f'<div class="{cls}">'
            + tab("slide", "Slide by slide", f"/design/{esc(job_id)}?n={current}")
            + tab("deck", f"Deck-wide{f' ({deck_n})' if deck_n else ''}",
                  f"/design/{esc(job_id)}?view=deck")
            + '<span class="grow"></span>'
            + f'<a class="btn ghost" href="/download/{esc(job_id)}" download>'
              f'Download the deck</a>'
            + f'<a class="btn ghost" href="{esc(back_href)}">'
              f'{esc(back_label)}</a></div>')


def _applied_block(job_id: str, applied: list, only_slide: int | None = None) -> str:
    """The decisions already made, each with a way back.

    Shown on the slide it belongs to as well as on the deck view, filtered by
    `only_slide`. A designer who fixes slide 3 and immediately wants it back
    should not have to find another tab to say so; the button belongs where the
    decision was made.

    A "leave it" is listed with the rest. It cost nothing to perform and it is
    still a decision, and a designer who changes their mind needs the finding
    back rather than a page that has quietly forgotten it was raised."""
    if only_slide is not None:
        applied = [a for a in applied if only_slide in (a.slides or [])]
    if not applied:
        return ""
    rows = []
    for entry in applied:
        acted = bool(entry.undo)
        if not entry.done:
            state = '<span class="pill error">failed</span>'
        elif acted:
            state = '<span class="pill changed">applied</span>'
        else:
            state = '<span class="pill info">left alone</span>'
        rows.append(
            f'<div class="done">{state}'
            f'<span class="grow"><b>{esc(entry.headline)}</b>'
            f'<div class="note">{esc(entry.label)} &mdash; '
            f'{esc(entry.detail)}</div></span>'
            f'<button class="btn ghost" type="submit" name="finding_ids" '
            f'value="{esc(entry.finding_id)}">'
            f'{"Undo" if acted else "Ask me again"}</button></div>')
    # The slide to come back to, so pressing Undo on slide 7 does not answer by
    # showing slide 1. Absent on the deck view, which has no slide to return to.
    back = (f'<input type="hidden" name="n" value="{only_slide}">'
            if only_slide is not None else "")
    where = ("What you have decided about this slide" if only_slide is not None
             else "What you have already decided")
    return f"""
<form method="post" action="/design/{esc(job_id)}/undo">
{back}
<div class="card">
  <div class="tag">Decisions</div>
  <h2 style="margin-top:0">{esc(where)}</h2>
  <p class="sub">Undo replays the state stored before the change - the same box,
  the same color, the same place in the drawing order - and leaves every other
  decision alone. A change that touched a shape another decision also touched
  comes back with it, because putting one of them back on its own would erase
  the other without saying so.</p>
  {''.join(rows)}
</div>
</form>"""


def _apply_form(job_id: str, current: int, body: str, any_pick: bool) -> str:
    if not any_pick:
        return body
    return f"""
<form method="post" action="/design/{esc(job_id)}/apply" id="dform">
<input type="hidden" name="n" value="{current}">
{body}
<div class="actions"><button class="btn primary" type="submit">Apply the
choices I made</button><span class="note" style="margin-left:0.7rem">Only the
cards you picked an answer on are touched, and you come back to this slide.
</span></div>
</form>"""


def render_design(*, deck_name: str, profile_name: str, job_id: str,
                  view: str = "slide", current: int = 0, total_slides: int = 0,
                  findings: list | None = None, deck_findings: list | None = None,
                  applied: list | None = None, audit_records: list | None = None,
                  rects: list | None = None, per_slide: dict | None = None,
                  banner: str = "", error: str | None = None,
                  render_error: str | None = None,
                  has_deck: bool = True, can_fix: bool = False,
                  promoted: set | None = None,
                  shot_tag: str = "",
                  auto: dict | None = None,
                  chat: bool = False, chat_note: str = "",
                  back: tuple | None = None,
                  job_tabs: str = "") -> str:
    findings = findings or []
    deck_findings = deck_findings or []
    applied = applied or []
    audit_records = audit_records or []
    rects = rects or []
    auto = auto or {}

    head = f"""
<span class="kicker">Design QC &middot; {esc(profile_name)}</span>
<h1 class="file">{esc(Path(deck_name).name)}</h1>
{job_tabs}"""

    notes = ""
    if banner:
        notes += (f'<div class="banner ok" role="status">&#10003;&nbsp; '
                  f'{esc(banner)}</div>')
    if error:
        notes += f'<div class="banner warn">{esc(error)}</div>'
    if not has_deck:
        notes += ('<div class="banner warn"><b>The deck is no longer held in '
                  'memory.</b> Newer audits replaced it, so nothing can be '
                  'changed from here. The findings below are still the ones '
                  'this deck had; re-upload it to act on them.</div>')

    tabs = _tabs(job_id, view, len(deck_findings), current, back)
    # Under the tabs and above the cards: a designer reads the question they
    # have before they read forty rows, and the answer usually tells them which
    # row to read (qc.ui_chat).
    ask = chat_panel(job_id, "audit", chat, chat_note)
    # Beside the ask box rather than in the nav: it is about THIS deck, and a
    # nav entry that needs a job id is a link that breaks when nothing is open.
    ask += (f'<p class="note" style="margin:-0.6rem 0 1.2rem">'
            f'What is this deck made of? '
            f'<a href="/checklist/{esc(job_id)}">Colour and type checklist</a>'
            f' &mdash; every colour and typeface, and which level each one comes '
            f'from. Nothing there changes the deck.</p>')

    if view == "deck":
        body = (_auto_card(job_id, current, "deck", auto, can_fix)
                + _applied_block(job_id, applied))
        if deck_findings:
            inner = ('<p class="sub">Decisions that are not about one slide. A '
                     'color spelled two ways across forty shapes is one '
                     'question, and a badge on every slide is one question; '
                     'asking either of them per slide is how a designer learns '
                     'to skip this page.</p>')
            for kind in _ORDER:
                group = [f for f in deck_findings if f.kind == kind]
                if not group:
                    continue
                inner += (f'<h2 style="margin:1.4rem 0 0.2rem">'
                          f'{esc(_KIND_LABEL.get(kind, kind))} '
                          f'<span class="note">({len(group)})</span></h2>'
                          f'<p class="sub" style="margin-top:0.2rem">'
                          f'{esc(_KIND_BLURB.get(kind, ""))}</p>')
                inner += "".join(_finding_card(f) for f in group)
            body += _apply_form(job_id, current, inner, True)
        elif not applied:
            body += ('<div class="card clean"><div class="mark">&#10003;</div>'
                     '<h2>Nothing deck-wide</h2><p>Every open decision belongs '
                     'to a single slide. Use the slide view.</p></div>')
        return _shell(f"Design QC: {deck_name}",
                      head + tabs + notes + ask + body + _STYLE + chat_css()
                      + _APPLY_JS)

    # ---- slide view
    counts = {"error": 0, "warning": 0, "info": 0}
    kinds: dict[str, int] = {}
    for f in findings:
        counts[f.severity] = counts.get(f.severity, 0) + 1
        kinds[f.kind] = kinds.get(f.kind, 0) + 1
    for r in audit_records:
        counts[r["severity"]] = counts.get(r["severity"], 0) + 1
        kinds[r["module"]] = kinds.get(r["module"], 0) + 1

    # One pin per finding, numbered down the list, and the SAME number on the
    # box over the render. Two numbering rules for one slide is worse than none.
    pins = {f.finding_id: i + 1 for i, f in enumerate(findings)}
    for r in rects:
        r["pin"] = pins.get(r.get("finding_id"))

    cards = ""
    for kind in _ORDER:
        group = [f for f in findings if f.kind == kind]
        if not group:
            continue
        cards += (f'<h3 style="margin:1.1rem 0 0.15rem" data-sev="all">'
                  f'{esc(_KIND_LABEL.get(kind, kind))}</h3>')
        cards += "".join(_finding_card(f, pins.get(f.finding_id),
                                       scope=f"slide {current + 1}")
                         for f in group)
    if not findings:
        cards = _nothing_here(job_id, current, deck_findings)

    right = (_filters(counts, sorted(kinds.items(), key=lambda kv: -kv[1]))
             + _auto_card(job_id, current, "slide", auto, can_fix)
             + _apply_form(job_id, current, cards, bool(findings))
             + _applied_block(job_id, applied, current)
             + _audit_rows(job_id, audit_records, current, can_fix, promoted))

    covering = sum(1 for f in deck_findings if current in (f.slides or []))
    body = (_pager(job_id, current, total_slides, len(findings),
                   also_here=covering)
            + _strip(job_id, current, per_slide or {}, total_slides)
            + f'<div class="dsplit"><div class="dshot">'
            + _shot(job_id, current, rects, render_error, shot_tag)
            + f'</div><div class="dlist">{right}</div></div>'
            + _pager(job_id, current, total_slides, len(findings),
                     sticky=False, also_here=covering))

    return _shell(f"Design QC: {deck_name} slide {current + 1}",
                  head + tabs + notes + ask + body + _STYLE + chat_css()
                  + _STICKY_JS + _APPLY_JS + _FILTER_JS)
