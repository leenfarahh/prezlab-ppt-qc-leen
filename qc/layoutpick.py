"""Which of the master's layouts does THIS slide belong on, when the file
cannot say - answered by ranking the master's own layouts and asking the
designer.

THIS REPLACED A VISION CALL (design lead, 31/08/2026). qc.layoutmatch rendered
the unplaced slides, rendered a contact sheet of the master's layouts, and asked
a model to pick one by name. It worked. It was also the one place where applying
a master depended on a model being reachable, funded and in a good mood, and a
run that lost it silently fell back to a preference list - so the same deck
prepared twice could be rebuilt two different ways with nothing on the page
saying which had happened.

The judgment was never the hard part. A designer with both files open answers
this in two seconds, because they can SEE that the slide is a two-column
comparison and that the master has a layout built for exactly that. What they
lacked was somewhere to say so: the choice was made mid-run, between an upload
and a rebuild, with no page in between. So the model came out and a step went
in, and applying a master is now end-to-end deterministic - the same deck, the
same master and the same picks produce the same file every time.

WHAT THIS DOES INSTEAD. For every slide the file could not place, it scores all
of the master's layouts against what the slide actually holds - columns, content
blocks, whether there is a heading - using the same structural reading the
coverage report is built on (qc.layoutgap.signature and layout_signature). Like
for like, both sides read the same way. The ranking is a suggestion and nothing
more: the designer picks, including picking something this file ranked last,
and their pick is recorded as theirs.

NOTHING HERE CALLS ANYTHING. No model, no renderer, no PowerPoint. It is
arithmetic over two files already in memory, so it cannot fail slowly, cannot
fail expensively, and cannot fail differently on a Tuesday. Every cap qc.
layoutmatch needed - twenty slides rendered, twenty calls, sixteen layouts on a
contact sheet - existed to bound a per-call cost that no longer exists, and none
of them survived the move. A hundred-slide deck gets a hundred answers.
"""

from dataclasses import dataclass, field

from .layoutgap import (CONTENT_BUCKET_CAP, MIN_BOXES_TO_JUDGE, describe,
                        fits, layout_signature, signature)

# --- calibration ----------------------------------------------------------
#
# The weights that order the dropdown. Only the ORDER is load-bearing: nothing
# is applied off the back of a score, the designer picks, and a layout ranked
# fourth is one click from a layout ranked first. Tuned so that the layout a
# designer would have reached for is at or near the top, not so that the top one
# is always right.

# Columns first, and by a distance. A two-column comparison put on a one-column
# layout is the defect that sends a slide back; a block count that is off by one
# is a text box a designer moves.
W_COLUMNS = 3.0
# Content boxes second: too few and content is orphaned, which is worse than too
# many, so an undershoot is penalised harder than an overshoot.
W_BOXES_SHORT = 2.0
W_BOXES_SPARE = 0.5
# A heading with nowhere to go is visible immediately; a title box left empty on
# a slide that has no heading is not.
W_TITLE_MISSING = 2.0
W_TITLE_SPARE = 0.4
# The master's own archetype token agreeing with the source slide's is real
# evidence - it is the file stating what the layout is for - so it earns a
# discount rather than a place in the ordering of its own.
BONUS_ARCHETYPE = 1.5
# A shared word in the two layout names ("Section divider" / "Section header")
# is weaker evidence than either, and it breaks ties rather than deciding them.
BONUS_NAME_TOKEN = 0.75

# Words that appear in half the layout names in any master and say nothing about
# what a layout is for. Matching on these would rank every layout equally.
_STOPWORDS = frozenset((
    "slide", "layout", "page", "and", "the", "with", "of", "or", "a", "an",
    "content", "text", "title", "new", "master", "template", "default",
))

# How many layouts the picker offers per slide. Every layout stays reachable -
# the dropdown carries all of them - but the ones above the fold are the ones
# worth reading, and a master with forty layouts would otherwise present forty
# indistinguishable rows.
SHORTLIST = 5

# What the "none of these fit" radio posts. Named here rather than spelled out
# in the page and again in the route: it is a protocol between three files, and
# a typo in one of them silently turns a refusal into a slide nobody looked at.
LEAVE = "__leave__"


@dataclass
class Candidate:
    """One of the master's layouts, scored against one slide."""
    name: str
    score: float
    fits: bool
    offers: str          # what it holds, as a sentence a designer reads
    why: str = ""        # why it does not fit, when it does not


@dataclass
class Choice:
    """One slide, the layout it is going onto, and what else it could go on."""
    slide_index: int
    current: str | None       # what plan_assignments picked unaided
    rule: str                 # how it picked: "fallback", "name", "archetype"
    reason: str               # why the designer is being asked about this one
    wants: str                # what the slide holds, as a sentence
    candidates: list = field(default_factory=list)
    settled: bool = False

    @property
    def suggested(self) -> str | None:
        """What the form pre-selects.

        A settled slide pre-selects WHERE IT ALREADY IS, not what the ranking
        would have chosen. The ranking is arithmetic and the file's own name
        match is a designer's stated intent; moving a matched slide because a
        score preferred something else would rewrite decisions nobody asked to
        revisit, and a page that does that on load is a page whose defaults
        cannot be trusted.
        """
        if self.settled:
            return self.current
        return self.candidates[0].name if self.candidates else self.current


def _tokens(name: str) -> set:
    out = set()
    for raw in (name or "").replace("_", " ").replace("-", " ").split():
        word = "".join(c for c in raw if c.isalnum()).casefold()
        if len(word) > 2 and word not in _STOPWORDS:
            out.add(word)
    return out


def _score(sig: dict, lay: dict, entry: dict, source_type: str | None,
           source_name: str) -> float:
    """How far this layout is from what the slide is asking for. Lower is
    closer; the number is never shown and never applied, it only sorts."""
    want_cols = max(1, int(sig.get("columns") or 1))
    got_cols = max(1, int(lay.get("columns") or 1))
    want_blocks = max(1, min(int(sig.get("blocks") or 1), CONTENT_BUCKET_CAP))
    got_boxes = int(lay.get("bodies") or 0)

    score = W_COLUMNS * abs(got_cols - want_cols)
    if got_boxes < want_blocks:
        score += W_BOXES_SHORT * (want_blocks - got_boxes)
    else:
        score += W_BOXES_SPARE * (got_boxes - want_blocks)
    if sig.get("title") and not lay.get("title"):
        score += W_TITLE_MISSING
    elif lay.get("title") and not sig.get("title"):
        score += W_TITLE_SPARE

    if source_type and entry.get("type") == source_type:
        score -= BONUS_ARCHETYPE
    if _tokens(source_name) & _tokens(entry.get("name") or ""):
        score -= BONUS_NAME_TOKEN
    return score


def _holds_content(sig: dict, lay: dict, ok: bool, why: str) -> tuple[bool, str]:
    """The one place where "cannot be judged" must not read as "fits".

    qc.layoutgap.fits answers a DIFFERENT question - is this placed slide a
    misfit worth reporting - and it answers "no" for a layout with no content
    boxes on purpose: a slide sitting on Blank is the migration pass's business,
    not the coverage report's (MIN_BOXES_TO_JUDGE).

    Reused here unguarded, that abstention became the top of the dropdown. The
    ranking sorts fitting layouts above non-fitting ones whatever the score, so
    a Cover - a title layout with zero body placeholders, which every master
    has - was the only "fitting" candidate for any slide with content, carrying
    the worst score in the list. A deck of diagrams was offered a cover, twenty
    six times (design lead, 02/09/2026).

    A layout that offers nowhere to put content cannot hold a slide that has
    some. It stays in the dropdown, ranked on its score like everything else;
    it just stops claiming to fit.
    """
    if not ok:
        return ok, why
    if int(lay.get("bodies") or 0) >= MIN_BOXES_TO_JUDGE:
        return ok, why
    blocks = int(sig.get("blocks") or 0)
    if not blocks:
        return ok, why          # an empty slide and an empty layout do agree
    return False, (f"the layout offers no content boxes and the slide has "
                   f"{_plural_blocks(blocks)}")


def _plural_blocks(n: int) -> str:
    from .layoutgap import _plural

    return _plural(n, "content block")


def rank(slide, layouts: list[dict], slide_w: int, slide_h: int,
         source_type: str | None = None,
         source_name: str = "") -> tuple[list, dict]:
    """(candidates best first, the slide's signature).

    Layouts that FIT sort ahead of layouts that do not, whatever their score,
    because "this one can hold your content" outranks "this one is nearly the
    right shape". Within each half the score orders them, and the layout name
    breaks ties so the list is stable between runs on the same files.
    """
    sig = signature(slide, slide_w, slide_h)
    out = []
    for entry in layouts:
        name = entry.get("name")
        if not name:
            continue
        lay = layout_signature(entry, slide_w)
        ok, why = fits(sig, lay)
        ok, why = _holds_content(sig, lay, ok, why)
        out.append(Candidate(
            name=name, score=_score(sig, lay, entry, source_type, source_name),
            fits=ok, offers=_offers_sentence(lay), why="" if ok else why))
    out.sort(key=lambda c: (not c.fits, c.score, c.name))
    return out, sig


def _offers_sentence(lay: dict) -> str:
    """What a layout holds, phrased the way the coverage report phrases it, so
    the two pages read as one tool."""
    from .layoutgap import _plural

    boxes = _plural(int(lay.get("bodies") or 0), "content box", "content boxes")
    cols = _plural(max(1, int(lay.get("columns") or 1)), "column")
    head = "a title over " if lay.get("title") else "no title, "
    return f"{head}{boxes} in {cols}"


def choices(deck_prs, layouts: list[dict], plans: list) -> list:
    """EVERY slide, in deck order, with the master's layouts ranked against it
    and a flag saying whether it is a question.

    Three kinds, and the flag is the difference between them.

    A FALLBACK is a slide the file could not place at all: nothing is known and
    the pick is the whole answer. A MISFIT was placed - by name, usually - onto
    a layout whose boxes its content does not fit, which is a claim about intent
    that only a person can settle: a designer who named two layouts the same
    thing may well have meant them to correspond even though the content has
    outgrown one of them. Both are questions, and `settled` is False on them.

    A MATCHED slide is one the file placed onto a layout its content fits.
    `settled` is True: it is not a question, it is not counted as one, and
    leaving it alone records nothing.

    IT USED TO BE OMITTED ENTIRELY, replaced by a line at the foot of the page
    saying how many were "not listed because there is nothing to decide about
    them" - and a designer reviewing a rebuild before it happens wants to SEE
    the deck, not be told the parts of it they are not allowed to look at
    (design lead, 02/09/2026). The reasoning behind the omission still holds
    and is now expressed by the flag rather than by the absence: a page that
    asks forty questions to surface four gets pressed through unread, so the
    four are what the count and the styling lead on. What changed is that the
    other thirty six are visible, checkable against the layout they are going
    onto, and changeable by anyone who disagrees.
    """
    from .layoutgap import misfits as find_misfits

    slide_w = int(deck_prs.slide_width or 0) or 1
    slide_h = int(deck_prs.slide_height or 0) or 1
    slides = list(deck_prs.slides)

    try:
        misfit_by_index = {m.slide_index: m
                           for m in find_misfits(deck_prs, layouts, plans)}
    except Exception:
        misfit_by_index = {}

    out = []
    for plan in plans:
        idx = plan.slide_index
        if idx >= len(slides):
            continue
        misfit = misfit_by_index.get(idx)
        settled = False
        if plan.match_rule in ("fallback", "none"):
            reason = (plan.note or "the file names no layout this master has")
        elif misfit is not None:
            reason = (f"matched by {plan.match_rule} to "
                      f"'{plan.target_layout}', but {misfit.reason}")
        else:
            settled = True
            reason = (f"matched by {plan.match_rule} to "
                      f"'{plan.target_layout}', which has room for what is on "
                      f"it")

        try:
            candidates, sig = rank(slides[idx], layouts, slide_w, slide_h,
                                   source_type=plan.source_type,
                                   source_name=plan.source_layout)
        except Exception:
            continue

        shortlist = _shortlist(candidates, plan.target_layout, settled)
        out.append(Choice(
            slide_index=idx, current=plan.target_layout,
            rule=plan.match_rule, reason=reason, wants=describe(sig),
            candidates=shortlist, settled=settled))
    return out


def _shortlist(candidates: list, current: str | None, settled: bool) -> list:
    """The five worth reading, with the layout the slide is ALREADY GOING ONTO
    guaranteed to be among them for a settled slide.

    Without this the page can pre-select something it is not showing: the
    ranking is by score, a name match is not, and a slide matched by name to a
    layout that scores seventh has that layout nowhere in its five radios. The
    designer then sees a card whose options are all wrong and no indication of
    where the slide is actually going. It leads because it is the answer.
    """
    if not settled or not current:
        return candidates[:SHORTLIST]
    here = next((c for c in candidates if c.name == current), None)
    if here is None:
        return candidates[:SHORTLIST]
    rest = [c for c in candidates if c.name != current]
    return [here] + rest[:SHORTLIST - 1]


def undecided(choices_list: list) -> int:
    """How many of those are actually questions. The number the page leads on
    and the number the run's note is written against - `len(choices)` counts
    the whole deck now and would report every slide as a decision nobody
    made."""
    return sum(1 for c in choices_list if not c.settled)


def apply_picks(plans: list, picks: dict, layouts: list[dict]) -> int:
    """Write the designer's picks onto the plans. Returns how many moved.

    A pick naming a layout the master does not have is DROPPED, not applied and
    not raised: the only way to send one is to edit the form, and a slide that
    keeps its computed target is a slide that still rebuilds. Same for a pick
    that names the layout the slide already had - it is recorded as a decision
    rather than a move, because a designer confirming the tool's suggestion and
    a designer never having looked are different facts about that slide, and the
    coverage report reads `match_rule` to tell them apart.

    LEAVE IS AN ANSWER, AND IT IS THE MOST INFORMATIVE ONE. A designer who reads
    the master's layouts against this slide and says none of them fit has told
    the tool something no measurement can: the master is missing a layout. That
    is exactly what the deck-level report is trying to establish
    (qc.layoutgap.Gap.refused), so it is recorded as a refusal rather than
    dropped on the floor next to the slides nobody opened.
    """
    known = {(l.get("name") or "").strip().casefold(): l.get("name")
             for l in layouts if l.get("name")}
    by_index = {p.slide_index: p for p in plans}
    moved = 0
    for idx, wanted in (picks or {}).items():
        plan = by_index.get(int(idx))
        if plan is None:
            continue
        if wanted == LEAVE:
            plan.review = "no fit"
            continue
        real = known.get((wanted or "").strip().casefold())
        if real is None:
            continue
        if real != plan.target_layout:
            moved += 1
            plan.target_layout = real
            plan.note = ""
        plan.match_rule = "chosen"
        plan.review = "chosen by the designer"
    return moved


def note(choices_offered: int, picked: int, moved: int,
         refused: int = 0, overridden: int = 0) -> str:
    """What happened at the layout step, for the run's own record.

    Reads off counts rather than off the plans, so the sentence on the results
    page and the plans in the file cannot become different claims.

    `refused` is said separately from "left on the fallback" because they are
    different facts: one is a designer saying this master has no home for the
    slide, the other is a slide nobody got to.

    `overridden` is separate for the same reason and is new with the page
    showing every slide (02/09/2026): a designer moving a slide the FILE
    matched is not answering one of the questions the page asked, it is
    disagreeing with a match nobody flagged. Counted apart so a run where the
    tool was never in doubt and a designer moved three slides anyway does not
    read as a run with three open questions. It is included in `moved`, which
    is the count of slides that actually changed layout.
    """
    if not choices_offered and not overridden:
        return ("Every slide matched a layout in this master by name or "
                "archetype, so there was nothing to choose.")
    bits = []
    if choices_offered:
        bits.append(f"{choices_offered} slide"
                    f"{'s' if choices_offered != 1 else ''} needed a layout "
                    f"decision")
    if moved:
        bits.append(f"{moved} moved to a layout you picked")
    # An override is BOTH a pick and a move, so it cancels out of this one:
    # subtracting it here as well would report a slide that was answered and
    # not moved as neither.
    kept = picked - moved
    if kept > 0:
        bits.append(f"{kept} kept the suggestion")
    if refused:
        bits.append(f"{refused} had no layout in this master that fits, which "
                    f"is a gap in the master rather than a fault on the slide")
    if overridden:
        bits.append(f"{overridden} of the moves {'were' if overridden != 1 else 'was'} "
                    f"a slide the file had already matched, which you changed")
    untouched = choices_offered - (picked - overridden) - refused
    if untouched > 0:
        bits.append(f"{untouched} were left on the fallback the file gave them")
    return ", ".join(bits) + "."
