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

BOTH SETS OF FINDINGS ARE HERE. The design pass owns colour, contrast, fit,
overlap and the frame; qc.modules owns fonts, margins, alignment, page
furniture. A designer looking at slide 7 wants slide 7's problems, not the ones
that happen to belong to whichever pass found them - so the audit's own records
for the slide are listed underneath, read-only, and say where their fixes live.

Every card is a choice and none has a default. A pre-selected remedy is the tool
deciding and asking for a rubber stamp, and the reason this page exists rather
than another entry in FIXABLE_ISSUES is that these calls are not the tool's to
make: recoloring the text and recoloring the panel are both correct fixes for
grey-on-grey, and only the person looking at the slide knows which the slide
wants. "Leave it" is on every card, is a real answer, and is recorded as one.

Rendering only; route logic lives in qc/web.py.
"""

from pathlib import Path

from .ui import MODULE_LABELS, esc, _shell, issue_label

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
.dshot { position:sticky; top:0.6rem }
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
.hidden { display:none !important }
</style>"""

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


def _options(finding) -> str:
    if not finding.options:
        return ""
    rows = "".join(
        f'<label class="radio-card"><input type="radio" '
        f'name="pick_{esc(finding.finding_id)}" value="{esc(o.remedy_id)}">'
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
     data-kind="{esc(finding.kind)}">
  <div class="difflabels">{_sev_pill(finding.severity)}
    <span class="note">{esc(_KIND_LABEL.get(finding.kind, finding.kind))}
    &middot; {esc(scope or _slides_note(finding.slides))}{spread}</span></div>
  <h3>{badge}{esc(finding.headline)}</h3>
  <p class="note">{esc(finding.detail)}</p>
  {_evidence(finding)}
  {_options(finding)}
</div>"""


def _audit_rows(records: list) -> str:
    """The audit's own findings for this slide, read-only.

    Read-only and saying so. Their fixes are ticked on the audit report, which
    applies them as a batch and re-audits; duplicating that control here would
    give a designer two buttons for one action with different consequences."""
    if not records:
        return ""
    rows = []
    for r in sorted(records, key=lambda r: r["severity"]):
        module = MODULE_LABELS.get(r["module"], r["module"])
        rows.append(
            f'<div class="auditrow" data-sev="{esc(r["severity"])}" '
            f'data-kind="{esc(r["module"])}">{_sev_pill(r["severity"])}'
            f'<span class="grow" style="flex:1"><b>{esc(issue_label(r["issue_type"]))}</b>'
            f'<div class="note">{esc(r["message"])}</div></span>'
            f'<span class="note" style="white-space:nowrap">{esc(module)}</span>'
            f'</div>')
    return f"""
<div class="card">
  <div class="tag">From the audit</div>
  <h3 style="margin:0.2rem 0 0.1rem">Also on this slide</h3>
  <p class="note">Fonts, margins, alignment and page furniture, found by the
  audit's own checks. Listed here so one slide's problems are in one place;
  these are fixed by ticking them on the audit report, which applies them
  together and re-audits.</p>
  {''.join(rows)}
</div>"""


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


def _pager(job_id: str, current: int, total: int, counted: int) -> str:
    def link(target, label, on):
        if not on:
            return f'<span class="btn ghost" aria-disabled="true">{label}</span>'
        return (f'<a class="btn ghost" href="/design/{esc(job_id)}?n={target}">'
                f'{label}</a>')

    where = (f'<span class="note"><b>Slide {current + 1}</b> of {total}'
             + (f" &middot; {counted} to decide here" if counted else
                " &middot; nothing to decide here") + '</span>')
    return (f'<div class="actionbar no-print">{where}<span class="grow"></span>'
            f'{link(current - 1, "&larr; Previous", current > 0)}'
            f'{link(current + 1, "Next &rarr;", current + 1 < total)}</div>')


def _shot(job_id: str, index: int, rects: list, error: str | None) -> str:
    if error:
        return (f'<div class="dframe" style="padding:1.1rem">'
                f'<p class="note">No render: {esc(error)} Everything on the '
                f'right is read from the deck itself, not from the picture.</p>'
                f'</div>')
    boxes = "".join(
        f'<div class="hit {esc(r["severity"])}" data-sev="{esc(r["severity"])}" '
        f'data-kind="{esc(r["kind"])}" title="{esc(r["label"])}" '
        f'style="left:{r["x"] * 100:.2f}%;top:{r["y"] * 100:.2f}%;'
        f'width:{r["w"] * 100:.2f}%;height:{r["h"] * 100:.2f}%">'
        + (f'<b>{r["pin"]}</b>' if r.get("pin") else "") + '</div>'
        for r in rects)
    return (f'<div class="dframe"><img src="/design-img/{esc(job_id)}/{index}.png"'
            f' alt="Slide {index + 1}" loading="lazy">{boxes}</div>')


# ------------------------------------------------------------------ the views


def _tabs(job_id: str, view: str, deck_n: int, current: int) -> str:
    def tab(key, label, href):
        on = " primary" if view == key else " ghost"
        return f'<a class="btn{on}" href="{href}">{label}</a>'

    return (f'<div class="actionbar no-print">'
            + tab("slide", "Slide by slide", f"/design/{esc(job_id)}?n={current}")
            + tab("deck", f"Deck-wide{f' ({deck_n})' if deck_n else ''}",
                  f"/design/{esc(job_id)}?view=deck")
            + '<span class="grow"></span>'
            + f'<a class="btn ghost" href="/download/{esc(job_id)}" download>'
              f'Download the deck</a>'
            + f'<a class="btn ghost" href="/audit/{esc(job_id)}">Back to the '
              f'audit</a></div>')


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
                  has_deck: bool = True) -> str:
    findings = findings or []
    deck_findings = deck_findings or []
    applied = applied or []
    audit_records = audit_records or []
    rects = rects or []

    head = f"""
<span class="kicker">Design QC &middot; {esc(profile_name)}</span>
<h1 class="file">{esc(Path(deck_name).name)}</h1>"""

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

    tabs = _tabs(job_id, view, len(deck_findings), current)

    if view == "deck":
        body = _applied_block(job_id, applied)
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
                      head + tabs + notes + body + _STYLE)

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
        cards = ('<div class="card clean"><div class="mark">&#10003;</div>'
                 '<h3 style="margin:0.4rem 0">Nothing to decide on this slide'
                 '</h3><p class="note">No color conflict, no unreadable text, '
                 'nothing overflowing its box, nothing hidden.</p></div>')

    right = (_filters(counts, sorted(kinds.items(), key=lambda kv: -kv[1]))
             + _apply_form(job_id, current, cards, bool(findings))
             + _applied_block(job_id, applied, current)
             + _audit_rows(audit_records))

    body = (_pager(job_id, current, total_slides, len(findings))
            + _strip(job_id, current, per_slide or {}, total_slides)
            + f'<div class="dsplit"><div class="dshot">'
            + _shot(job_id, current, rects, render_error)
            + f'</div><div class="dlist">{right}</div></div>'
            + _pager(job_id, current, total_slides, len(findings)))

    return _shell(f"Design QC: {deck_name} slide {current + 1}",
                  head + tabs + notes + body + _STYLE + _FILTER_JS)
