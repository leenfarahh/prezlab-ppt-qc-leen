"""One page for the whole job: the master applied, and the deck audited after.

A designer used to do this in three visits. Check the coverage, go back and
apply the master, upload the rebuilt file to the audit, and hold the first two
answers in their head while reading the third. The three answers are about the
same afternoon's work and two of them are only true in that order, so the page
is one page (design lead, 27/08/2026).

WHAT IT LEADS WITH IS WHAT IS LEFT TO DO, not what was done. "Rebuilt 34 of 40"
is a fact about the past. The three things below it are the future:

    the master is missing this      a layout to build, and no amount of fixing
                                    slides produces one
    the tool can do this for you    one press, counted before it is pressed
    and this is left for you        with the slides it is on

Those are three different kinds of work for three different moments, and mixing
them into one list of findings is how a designer ends up reading two thousand
rows to find the four that matter.

TWO LISTS, NEVER MERGED. The gaps are about the MASTER and the findings are
about the SLIDES. They look similar on a page and they are not the same thing:
a gap is a conversation with whoever owns the master, a finding is a tick.

Rendering only; route logic lives in qc/web.py, and the pipeline in qc/prep.py.
"""

from .ui import _shell, esc
from .ui_chat import chat_css, chat_panel
from .ui_check import render_coverage
from .ui_format import (RULE_LABEL, RULE_ORDER, _content_section,
                        _masters_note, _slide_rows, _warn)
from .ui_master import spec_review

_CSS = """
<style>
.leadrow { display:flex; flex-wrap:wrap; gap:0.9rem; align-items:stretch;
  margin:1rem 0 0.4rem; }
.lead { flex:1 1 15rem; background:#fff; border:1px solid var(--line-soft);
  border-radius:12px; padding:0.95rem 1.05rem; }
.lead .n { font-size:1.8rem; font-weight:700; color:var(--teal);
  line-height:1.1; }
.lead .l { font-size:0.78rem; color:var(--slate-text); margin-top:0.15rem; }
.lead .w { font-size:0.82rem; margin-top:0.5rem; color:var(--teal); }
.lead.quiet .n { color:var(--slate-text); }
.handover { background:var(--sand); border:1px solid var(--line);
  border-radius:12px; padding:1.1rem 1.2rem; margin:1.1rem 0; }
.worst { display:flex; flex-wrap:wrap; gap:0.4rem; margin:0.6rem 0 0; }
.worst a { text-decoration:none; }
.stepnum { display:inline-block; min-width:1.4rem; height:1.4rem;
  line-height:1.4rem; text-align:center; border-radius:999px;
  background:var(--teal); color:#fff; font-size:0.78rem; font-weight:700;
  margin-right:0.45rem; }
</style>
"""


def _plural(n: int, one: str, many: str | None = None) -> str:
    return f"{n} {one if n == 1 else (many or one + 's')}"


def _looked(note: str, ok: bool) -> str:
    """What the visual model was asked and what it said.

    Said out loud whether or not it found anything, and said even when it could
    not run. "Asked for" and "possible" are different things: a deck judged from
    its geometry alone looks identical on this page to one a designer's eye went
    over, and it is a materially weaker answer.
    """
    if not note:
        return ""
    if not ok:
        return _warn(note)
    return (f'<p class="note" style="margin:0.6rem 0 0"><b>Looked at:</b> '
            f'{esc(note)}</p>')


def _step(number: int, title: str, blurb: str) -> str:
    return (f'<h2 style="margin:1.6rem 0 0.2rem"><span class="stepnum">'
            f'{number}</span>{esc(title)}</h2>'
            f'<p class="sub" style="margin:0 0 0.7rem">{esc(blurb)}</p>')


# --- the three numbers at the top -----------------------------------------


def _lead(prep, design_open: int | None, audit_note: str) -> str:
    """Rebuilt, missing, left. Three cards because they are three different
    jobs, and a single "issues" number would hide the only one a designer
    cannot fix by ticking something."""
    cov = prep.coverage
    unplaced = getattr(cov, "unplaced", 0) if cov is not None else 0
    n_gaps = len(getattr(cov, "gaps", None) or []) if cov is not None else 0

    rebuilt = (f'<div class="lead"><div class="n">{prep.applied}'
               f'<span style="font-size:1rem;color:var(--slate-text)">'
               f' / {prep.slides}</span></div>'
               f'<div class="l">slides rebuilt on the master</div>'
               + (f'<div class="w">{_plural(len(prep.errors), "slide")} could '
                  f'not be rebuilt and {"were" if len(prep.errors) != 1 else "was"} '
                  f'left exactly as {"they were" if len(prep.errors) != 1 else "it was"}.'
                  f'</div>' if prep.errors else "")
               + '</div>')

    missing = (f'<div class="lead{"" if unplaced else " quiet"}">'
               f'<div class="n">{n_gaps}</div>'
               f'<div class="l">layout{"s" if n_gaps != 1 else ""} the master '
               f'does not have</div>'
               + (f'<div class="w">{_plural(unplaced, "slide")} had nowhere to '
                  f'go and kept a fallback. That is a change to the master, not '
                  f'to these slides.</div>' if unplaced else
                  '<div class="w">Every slide in this deck has a home in this '
                  'master.</div>')
               + '</div>')

    if prep.manifest is None:
        left = ('<div class="lead quiet"><div class="n">&mdash;</div>'
                '<div class="l">findings on the rebuilt deck</div>'
                f'<div class="w">{esc(audit_note or "The audit did not run.")}'
                '</div></div>')
    else:
        total = prep.findings + (design_open or 0)
        left = (f'<div class="lead{"" if total else " quiet"}">'
                f'<div class="n">{total}</div>'
                f'<div class="l">things open on the rebuilt deck</div>'
                f'<div class="w">'
                + (f'{_plural(prep.findings, "audit finding")}'
                   + (f' and {_plural(design_open, "design decision")}'
                      if design_open else "")
                   + '. The card below says how many of those the tool can '
                     'take by itself.'
                   if total else
                   'Nothing left against the profile, and no design decision '
                   'open.')
                + '</div></div>')

    return f'<div class="leadrow">{rebuilt}{missing}{left}</div>'


# --- step two: what is left on the slides ---------------------------------


def _handover(job_id: str, auto: dict, design_open: int | None) -> str:
    """The one press. Counted before it is pressed, from the same selection the
    route will act on, so the number a designer reads is the number that
    happens (qc.web._auto_targets counts and selects in one function for
    exactly this reason)."""
    deck = (auto or {}).get("deck") or {}
    fixes, picks = deck.get("fixes") or 0, deck.get("picks") or 0
    left, held = deck.get("left") or 0, deck.get("held") or 0
    if not (fixes or picks or held or left):
        return ('<div class="banner ok">Nothing on this deck is waiting on a '
                'decision. Download it and open it in PowerPoint.</div>')

    if not (fixes or picks):
        return (f'<div class="handover"><div class="tag">Nothing to hand over'
                f'</div><p class="note" style="margin:0.3rem 0 0">'
                f'{_plural(left + held, "decision")} on this deck '
                f'{"are" if left + held != 1 else "is"} yours to make, and none '
                f'of them are the tool\'s to decide. Open the slides and they '
                f'are one tick each.</p></div>')

    bits = []
    if fixes:
        bits.append(_plural(fixes, "audit fix", "audit fixes"))
    if picks:
        bits.append(_plural(picks, "design decision"))

    reasons = ""
    if left:
        why = "; and ".join(esc(r) for r in (deck.get("reasons") or [])[:3])
        reasons = (f'<p class="note" style="margin:0.5rem 0 0">'
                   f'{_plural(left, "decision")} would be left for you'
                   + (f", because {why}." if why else ".") + '</p>')
    holds = ""
    if held:
        holds = (f'<p class="note" style="margin:0.35rem 0 0">'
                 f'{_plural(held, "fix", "fixes")} '
                 f'{"ask" if held != 1 else "asks"} for your explicit approval '
                 f'- Arabic font substitutions and whole-slide body moves - and '
                 f'{"are" if held != 1 else "is"} never included here. They are '
                 f'a tick each on the slide.</p>')

    return f"""
<div class="handover">
  <div class="tag">Hand it over</div>
  <h3 style="margin:0.2rem 0 0.2rem">Let the tool do {" and ".join(bits)}</h3>
  <p class="note" style="margin:0">The audit fixes land first and the deck is
  re-audited between the two passes, so the design judgments are made about the
  deck as the fixes leave it. Every one of them lands in the same list as a
  hand-picked decision, with the same Undo: this is a starting point you
  correct, not a commitment.</p>
  <form action="/design/{esc(job_id)}/auto" method="post"
        style="margin-top:0.75rem">
    <input type="hidden" name="scope" value="deck">
    <input type="hidden" name="n" value="0">
    <button class="btn primary" type="submit"
            data-busy="Deciding the whole deck"
            data-busysub="Applying the fixes, re-auditing, then taking the design decisions. This can take a minute.">
      Decide the whole deck
    </button>
    <a class="btn ghost" href="/design/{esc(job_id)}">Go slide by slide instead</a>
  </form>
  {reasons}{holds}
</div>"""


_SEVERITY_ORDER = ("error", "warning", "info")


def _worst_slides(per_slide: dict, job_id: str, limit: int = 8) -> str:
    """The slides to open first, as buttons that open them.

    A roll-up by slide rather than by rule, because a designer works a slide at
    a time: eleven findings spread over eleven slides and eleven on one slide
    are the same number and completely different afternoons."""
    if not per_slide:
        return ""
    ranked = sorted(
        per_slide.items(),
        key=lambda kv: (-(kv[1].get("error", 0) * 100 + kv[1].get("warning", 0)
                          * 10 + kv[1].get("info", 0)), kv[0]))[:limit]
    ranked = [(i, counts) for i, counts in ranked if sum(counts.values())]
    if not ranked:
        return ""
    chips = "".join(
        f'<a class="btn ghost" href="/design/{esc(job_id)}?n={i}">'
        f'Slide {i + 1} &middot; '
        + ", ".join(f"{counts[s]} {s}" for s in _SEVERITY_ORDER
                    if counts.get(s))
        + '</a>'
        for i, counts in ranked)
    more = len([1 for c in per_slide.values() if sum(c.values())]) - len(ranked)
    tail = (f'<p class="note" style="margin:0.5rem 0 0">and '
            f'{_plural(more, "more slide")} with something on '
            f'{"them" if more != 1 else "it"}.</p>' if more > 0 else "")
    return f"""
<div class="card">
  <div class="tag">Open these first</div>
  <h3 style="margin:0 0 0.2rem">The slides carrying the most</h3>
  <p class="note" style="margin:0">Both passes counted together, because which
  check found it is not the question a designer standing on a slide is
  asking.</p>
  <div class="worst">{chips}</div>
  {tail}
</div>"""


def _by_issue(manifest: dict, job_id: str, limit: int = 12) -> str:
    """What the audit found, grouped by rule, so the tail of a long audit does
    not have to be read to know what is in it."""
    records = [r for r in (manifest.get("records") or [])
               if r.get("module") != "preflight"]
    if not records:
        return ""
    counts: dict[tuple, int] = {}
    for record in records:
        key = (record.get("issue_type") or "", record.get("severity") or "info")
        counts[key] = counts.get(key, 0) + 1
    ranked = sorted(counts.items(), key=lambda kv: -kv[1])
    rows = "".join(
        f'<tr><td><code>{esc(issue)}</code></td>'
        f'<td><span class="pill {esc(severity)}">{esc(severity)}</span></td>'
        f'<td>{n}</td></tr>'
        for (issue, severity), n in ranked[:limit])
    more = (f'<p class="note" style="margin:0.5rem 0 0">and '
            f'{_plural(len(ranked) - limit, "more rule")} with fewer '
            f'occurrences each. Every row is on the '
            f'<a href="/audit/{esc(job_id)}">full audit report</a>.</p>'
            if len(ranked) > limit else "")
    return f"""
<div class="card">
  <div class="tag">By rule</div>
  <h3 style="margin:0 0 0.4rem">What the audit found on the rebuilt deck</h3>
  <table class="w3"><thead><tr><th>Rule</th><th>Severity</th><th>Times</th>
  </tr></thead><tbody>{rows}</tbody></table>
  {more}
</div>"""


# --- the page -------------------------------------------------------------


def render_prep_intake(profiles: list[dict], message: str = "",
                       com_ready: bool = True, look: bool = True,
                       look_note: str = "", saved: str = "",
                       can_save: bool = True, spec: dict | None = None,
                       spec_id: str = "", spec_message: str = "",
                       replaced: bool = False) -> str:
    """profiles: [{id, name, has_master, layouts, frame, master_stored}]. Only
    ones carrying a master can be applied, and the rest are named WITH the
    reason - a silently missing option reads as a bug.

    `spec` is the master just read, if one was. It renders UNDER step 1 on this
    same page rather than on a page of its own, so reading a master, saving it
    as a profile, and applying it to a deck are one destination."""
    usable = [p for p in profiles if p["has_master"]]
    unusable = [p for p in profiles if not p["has_master"]]
    saved_name = next((p["name"] for p in profiles if p["id"] == saved), saved)

    def _frame_label(p) -> str:
        frame = (p.get("frame") or "").replace("_", " ")
        bits = [f"{p['layouts']} layouts"]
        if p.get("master_stored"):
            bits.append(f"stored {p['master_stored']}")
        if frame:
            bits.append(f"frame from {frame}")
        return " &middot; ".join(bits)

    picked = saved if any(p["id"] == saved for p in usable) else (
        usable[0]["id"] if usable else "")
    cards = "".join(
        f"""<label class="radio-card"><input type="radio" name="profile"
        value="{esc(p['id'])}"{' checked' if p['id'] == picked else ''}>
        <span><b>{esc(p['name'])}</b><small>{_frame_label(p)}</small></span></label>"""
        for p in usable)
    if not usable:
        cards = ('<p class="note">No profile carries a master yet. Save one '
                 'from a master file in step 1; the master is stored with the '
                 'profile JSON and this step can then apply it.</p>')

    others = ""
    if unusable:
        names = ", ".join(esc(p["name"]) for p in unusable)
        others = (f'<p class="note">Not listed, because they carry no master '
                  f'file to apply: {names}.</p>')

    blocker = ""
    if not com_ready:
        blocker = _warn(
            "Desktop PowerPoint is not reachable on this machine, and applying "
            "a master needs it: PowerPoint's own placeholder matching is what "
            "moves each slide's content into the new layout. Run this on the "
            "Windows box.")
    disabled = " disabled" if (not usable or not com_ready) else ""

    look_box = ""
    if look:
        look_box = """
    <fieldset><legend>Look at the rebuilt slides</legend>
      <label class="radio-card"><input type="checkbox" name="look" value="1"
        checked><span><b>Ask the visual model what a designer would adjust</b>
        <small>The question measuring cannot answer: which edges were MEANT to
        line up. Proximity alone has to guess at intent, so it stays quiet on
        the ones that are furthest out, and it cannot tell a card from the three
        shapes it is made of. This runs after the rebuild, on the slides you are
        actually going to send. Every judgment is re-measured here and arrives
        as an ordinary tickable finding, never pre-selected. Slide images leave
        this machine. Which layout each slide lands on is not asked: you choose
        that yourself on the next page.</small></span></label>
    </fieldset>"""
    elif look_note:
        look_box = f'<p class="note">{esc(look_note)}</p>'

    saved_banner = ""
    if saved and replaced:
        saved_banner = (
            f'<div class="banner ok">Master replaced on <b>{esc(saved_name)}</b>. '
            f'Its frame, reserved bands, grid and layout names were re-read from '
            f'this file; its fonts, palette and tolerances were left alone. Every '
            f'deck prepared against it is now rebuilt on this master.</div>')
    elif saved:
        saved_banner = (
            f'<div class="banner ok">Saved <b>{esc(saved_name)}</b> as a '
            f'profile. It is selected below; drop the messy deck and apply '
            f'it.</div>')

    save_note = (
        '<p class="note">Sign in as a lead or admin to save the profile JSON. '
        'You can still read the master and review what it declares.</p>'
        if not can_save else
        '<p class="note">The master file is stored with the profile so a later '
        'deck can be rebuilt onto it, not only audited against the rules.</p>')

    # The read master, in place. The anchor is what the page scrolls to: a spec
    # rendered below the fold looks like nothing happened, and the designer
    # presses Read the master again.
    read_block = ""
    if spec is not None:
        read_block = (f'<div id="spec">'
                      f'{spec_review(spec, spec_id, can_save=can_save, message=spec_message, profiles=profiles)}'
                      f'</div>')

    body = f"""
<h1>Prepare a deck.</h1>
<p class="sub">Save a finished master as a profile, then apply that profile to a
messy client deck. The rebuild happens first and the audit reads what it left,
in that order: an audit of the raw file reports margins the master is about to
reset and fonts it is about to replace.</p>
{_warn(message)}{saved_banner}{blocker}

{_step(1, "Save a master as a profile",
       "Upload the finished master. This reads the slide master, its layouts, "
       "and the theme, then writes a JSON profile you can apply to another deck.")}
<form action="/master" method="post" enctype="multipart/form-data" id="mf">
  <div class="drop" id="mdrop" tabindex="0" role="button"
       aria-label="Drop a .pptx master here or press Enter to browse">
    <strong>Drop the master .pptx here</strong> or click to browse
    <div class="hint">Read locally and deleted straight after. Slide content is
    never read, so a master with no slides is exactly what this expects.</div>
    <div class="file" id="mname" aria-live="polite"></div>
    <input type="file" name="master" id="masterfile" accept=".pptx" required hidden>
  </div>
  <div class="actions">
    <button class="btn primary" id="mgo" type="submit" disabled>Read the master</button>
  </div>
</form>
{save_note}
{read_block}

{_step(2, "Apply it to a messy deck",
       "Reads the deck against the stored master and shows you which slides "
       "need a layout choosing. Nothing is rebuilt until you approve them; "
       "then every slide is copied onto its layout and the result is audited.")}
<form action="/prep" method="post" enctype="multipart/form-data" id="f">
  <div class="drop" id="drop" tabindex="0" role="button"
       aria-label="Drop the deck to prepare here or press Enter to browse">
    <strong>Drop the client deck here</strong> or click to browse
    <div class="hint">Processed on this machine. The upload is deleted after
    processing and the rebuilt deck is held in memory for download.</div>
    <div class="file" id="fname" aria-live="polite"></div>
    <input type="file" name="deck" id="deck" accept=".pptx" required hidden>
  </div>
  <div class="config">
    <fieldset><legend>Profile to apply</legend>{cards}{others}</fieldset>
    {look_box}
  </div>
  <div class="actions">
    <button class="btn primary" id="go" type="submit"{disabled}>Choose the layouts</button>
  </div>
</form>
<script>
function wireDrop(drop, input, nameEl, go, blocked) {{
  function arm() {{
    if (input.files.length) {{ nameEl.textContent = input.files[0].name; }}
    if (go) go.disabled = blocked || !input.files.length;
  }}
  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', e => {{
    if (e.key === 'Enter' || e.key === ' ') input.click(); }});
  input.addEventListener('change', arm);
  ['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => {{
    e.preventDefault(); drop.classList.add('armed'); }}));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => {{
    e.preventDefault(); drop.classList.remove('armed'); }}));
  drop.addEventListener('drop', e => {{
    if (e.dataTransfer.files.length) {{ input.files = e.dataTransfer.files; arm(); }} }});
  arm();
}}
const mdrop = document.getElementById('mdrop'),
      minput = document.getElementById('masterfile'),
      mname = document.getElementById('mname'),
      mgo = document.getElementById('mgo');
wireDrop(mdrop, minput, mname, mgo, false);
document.getElementById('mf').addEventListener('submit', () => {{
  mgo.disabled = true;
  showBusy('Reading ' + (minput.files[0] ? minput.files[0].name : 'master'),
           'Walking the slide master, every layout, and the theme part.');
}});
const drop = document.getElementById('drop'), input = document.getElementById('deck'),
      fname = document.getElementById('fname'), go = document.getElementById('go');
const blocked = {str(bool(disabled)).lower()};
wireDrop(drop, input, fname, go, blocked);
document.getElementById('f').addEventListener('submit', () => {{
  go.disabled = true;
  showBusy('Preparing ' + (input.files[0] ? input.files[0].name : 'the deck'),
           'Rebuilding one slide at a time through PowerPoint, then auditing the result. Expect roughly a second per slide.');
}});
// The spec lands below the form that asked for it, so the answer to "did that
// work" is off screen unless the page goes to it.
const specBlock = document.getElementById('spec');
if (specBlock) specBlock.scrollIntoView({{behavior: 'smooth', block: 'start'}});
</script>"""
    return _shell("Prepare a deck", body)


def _component_review(job_id: str) -> str:
    """The visual pass, offered where the work is.

    It answers the two questions the geometry rules cannot, and it was reachable
    only from the audit report, two clicks away from the page a designer is
    actually on. Opt-in and never pre-selected, because it costs a vision call
    per slide and it sends slide IMAGES off this machine; that is why it is a
    button with the cost written on it rather than part of the run."""
    return f"""
<div class="card">
  <div class="tag">Ask the visual model</div>
  <h3 style="margin:0 0 0.3rem">Component review</h3>
  <p class="sub" style="margin:0 0 0.7rem">Answers what measuring cannot:
  which shapes are <b>one thing</b> (a card with its icon and its label), and
  which line they were meant to share, the master's frame or another component.
  It catches the misalignments the geometry rules skip on purpose, a label
  sitting off the block it heads being the common one. The model never supplies
  a coordinate: it names the intended line and the geometry is measured here,
  so every suggestion arrives tickable with a real target.</p>
  <form method="post" action="/components/{esc(job_id)}"
    onsubmit="showBusy('Working out what the components are',
     'Naming the things on each slide and the line they belong on, then measuring the geometry here. A full deck takes a minute or two.')">
   <button class="btn ghost" type="submit">Review the components</button>
  </form>
  <p class="note" style="margin:0.6rem 0 0">Slide <b>images</b> leave this
  machine for this one. Use it on decks approved for cloud processing.
  Suggestions are never pre-selected.</p>
</div>"""


def render_prep_result(*, prep, job_id: str, profile_name: str,
                       headline: str, auto: dict | None = None,
                       design_open: int | None = None,
                       design_error: str | None = None,
                       per_slide: dict | None = None,
                       banner: str = "", error: str | None = None,
                       chat: bool = False, chat_note: str = "",
                       restored: list | None = None,
                       restored_notes: dict | None = None,
                       restore_error: str | None = None,
                       removed: list | None = None,
                       remove_error: str | None = None,
                       layout_note: str = "", layout_ok: bool = True,
                       tabs: str = "") -> str:
    """The whole run on one page: what was rebuilt, what the master is missing,
    and what is left on the slides."""
    counts: dict[str, int] = {}
    for plan in prep.plans:
        key = "failed" if plan.slide_index in prep.errors else plan.match_rule
        counts[key] = counts.get(key, 0) + 1
    chips = []
    for rule in RULE_ORDER:
        if counts.get(rule):
            label, why = RULE_LABEL[rule]
            chips.append(f'<span class="fchip" title="{esc(why)}">'
                         f'<b>{counts[rule]}</b> {esc(label)}</span>')
    if counts.get("failed"):
        chips.append(f'<span class="fchip"><b>{counts["failed"]}</b> failed</span>')

    space_note = ""
    if prep.space_notes:
        space_note = ("<div class='card'><div class='tag'>Presentation space</div>"
                      "<ul style='margin:0.4rem 0 0 1.1rem'>"
                      + "".join(f"<li>{esc(n)}</li>" for n in prep.space_notes)
                      + "</ul></div>")

    # Every link that reads the manifest is drawn only when there is one. A run
    # whose audit failed still produced a rebuilt deck, and offering it a
    # Design QC button that 404s would turn one honest failure into two.
    audited = prep.manifest is not None
    if audited:
        slides_block = (_handover(job_id, auto or {}, design_open)
                        + _worst_slides(per_slide or {}, job_id)
                        + _by_issue(prep.manifest, job_id))
        slide_links = (
            f'<a class="btn primary" href="/design/{esc(job_id)}">Work through '
            f'the slides</a>')
        report_link = (f'<a class="btn ghost" href="/audit/{esc(job_id)}">Full '
                       f'audit report</a>')
    else:
        slides_block = _warn(prep.audit_note or
                             "The audit did not run on the rebuilt deck.")
        slide_links = report_link = ""

    body = f"""
<h1>{esc(prep.filename)}</h1>
<p class="sub">{esc(headline)} Against <b>{esc(profile_name)}</b>.</p>
{tabs}
{_warn(banner)}{_warn(error or "")}
{_lead(prep, design_open, prep.audit_note)}
<div class="actions" style="gap:0.6rem">
  {slide_links}
  <a class="btn ghost" href="/format/{esc(job_id)}/download">Download the rebuilt deck</a>
  <a class="btn ghost" href="/format/{esc(job_id)}/review?view=master">Review before / after</a>
  <a class="btn ghost" href="/checklist/{esc(job_id)}">Colours and type</a>
  {report_link}
</div>
<p class="note">Look before you download: undoing after the file has gone out
means doing it twice.</p>
{chat_panel(job_id, "prep", chat, chat_note)}
{_looked(layout_note, layout_ok)}
{_warn(prep.match_note)}{_warn(design_error or "")}
{_masters_note(prep.masters, prep.stragglers or [], prep.plans)}

{_step(1, "What the master could not build",
       "A change to the master, not to these slides. No amount of ticking here "
       "produces a layout that does not exist, which is why it is first: it is "
       "the half that needs somebody else.")}
{render_coverage(prep.coverage, suggestions=prep.suggestions,
                 pictures=prep.pictures, suggest_note=prep.suggest_note,
                 check_id=job_id, master_name=profile_name,
                 propose_job=job_id)}

{_step(2, "What is left on the slides",
       "The rebuilt deck read against the profile, plus the judgments that only "
       "have answers once a master is on the deck. Start with the hand-over: it "
       "is counted before it is pressed.")}
{slides_block}
{_component_review(job_id) if audited else ""}

{_step(3, "What the rebuild itself did",
       "Every slide's layout, and every piece of content the migration moved. "
       "Read this when a slide looks wrong and you want to know why.")}
<div class="kpis">{''.join(chips)}</div>
{space_note}
{_content_section(prep.changes or [], job_id, restored or [], restore_error,
                  restored_notes or {}, removed or [], remove_error)}
<div class="card">
  <div class="tag">Per slide</div>
  <h3 style="margin:0 0 0.2rem">Layout assignment</h3>
  <p class="sub">Matching runs by name first, then by archetype, then by looking
  at the slide, then falls back. Every row says which rule chose its target, so
  a surprising result is traceable rather than mysterious.</p>
  <table class="w3">
    <thead><tr><th>#</th><th>Was on</th><th>Now on</th><th>How</th>
    <th>Notes</th></tr></thead>
    <tbody>{_slide_rows(prep.plans, prep.errors)}</tbody>
  </table>
</div>
<p class="note">Open the result in PowerPoint before sending it anywhere. This
hands each slide to PowerPoint's own placeholder matching, which preserves
content but can move it; that is exactly the judgment a designer signs off.</p>
{_CSS}{chat_css()}"""
    return _shell(f"Prepared: {prep.filename}", body)
