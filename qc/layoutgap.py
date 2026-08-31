"""What this deck needs that the master does not have.

qc.layoutpick answers a per-slide question: which of the master's layouts does
THIS slide belong on. It ranks them and a designer picks, and when the answer is
"none of these" the slide keeps its fallback and the report says so, one row at
a time.

That leaves the deck-level question unasked, and it is the one a designer acts
on. Six slides that all fell back are not six problems. They are usually ONE
problem stated six times - the deck is full of two-column comparisons and the
master has no two-column layout - and the fix is not to open six slides. It is
to add a layout to the master and format again.

So this pass reads the plans the format pass already produced and rolls them up:

    coverage   how each slide got its target - by name, by archetype, by
               looking, or by falling back
    gaps       the fallbacks CLUSTERED by the shape of their content, each with
               how many slides want it and what the master offers instead
    unused     layouts the master defines that no slide in this deck lands on

Nothing here calls a model. Every judgment it reports was already made and
recorded by qc.applymaster and qc.layoutpick; this file counts and clusters. That matters for a
report a designer takes to a client conversation: "eleven slides want a layout
you do not have" has to be a count of the same evidence they can open and see,
not a second opinion produced by a second call.

A CLUSTER IS A STRUCTURAL SIGNATURE, not a similarity score. Two slides cluster
when they need the same thing built: the same number of content blocks, in the
same number of columns, with or without a title. Counting is honest about what
it can see and coarse on purpose - three text blocks and four in two columns are
the same request to a master - because a report with eleven clusters of one
slide each is the per-slide table again, and that table already exists.
"""

from dataclasses import dataclass, field

from .design import placed_shapes
from .extract import content_type

# --- calibration ----------------------------------------------------------
#
# Every number here decides whether two slides are counted as wanting the same
# layout, so each one is stated with what it is for.

# A shape covering most of the slide is a ground or a backdrop panel, not a
# column of content. Counting it as content makes every slide that has one look
# like it needs an extra box.
BACKDROP_AREA_SHARE = 0.7

# Content this small is furniture: a page number, a source line, a corner rule.
MIN_CONTENT_AREA_SHARE = 0.004

# A title sits in the top band and is not tall. Both conditions, because a
# full-height text column starting at the top of the slide is a column.
TITLE_BAND = 0.28
TITLE_MAX_HEIGHT_SHARE = 0.22

# Two items whose horizontal centres are closer than this are in the same
# column. At 8% of the slide width, a 13.3-inch slide reads centres within
# about an inch as one column, which is what a stack of blocks looks like.
COLUMN_TOLERANCE = 0.08

# Past four content blocks the exact count stops changing what the master needs
# to offer, and clustering on it would split one request into four.
CONTENT_BUCKET_CAP = 4


# A layout has to offer at least this many content boxes before its shape can
# be wrong. A layout with none makes no promises: Blank is not a layout that
# fails to fit a slide, it is a layout that holds nothing, and content sitting
# on it is the migration pass's business rather than this one's.
MIN_BOXES_TO_JUDGE = 1


@dataclass
class Misfit:
    """A slide that WAS placed, on a layout its content does not fit."""
    slide_index: int
    layout: str
    rule: str                   # how it was placed: name, archetype, reviewed
    label: str                  # what the slide is asking for, as a sentence
    offers: str                 # what the layout offers, as a sentence
    signature: dict = field(default_factory=dict)
    reason: str = ""
    # Set when the slides were actually looked at: "no fit" means a model saw
    # the slide against this layout and said no, which is evidence of a
    # different order from a signature comparison.
    review: str = ""


def fits(sig: dict, lay: dict) -> tuple[bool, str]:
    """Whether a slide of this shape fits a layout of that shape, and why not.

    The same three questions the proposal check asks, from the other side: does
    it have the columns, does it have the boxes, does it have a title where the
    slide carries one. Crude on purpose - it decides whether to ASK a designer
    to look, and a check that needs its own tolerance table would be a check
    nobody could predict.
    """
    want_columns = max(1, int(sig.get("columns") or 1))
    got_columns = max(1, int(lay.get("columns") or 0))
    want_blocks = max(1, int(sig.get("blocks") or 1))
    got_boxes = int(lay.get("bodies") or 0)

    if got_boxes < MIN_BOXES_TO_JUDGE:
        return True, ""
    if got_columns and got_columns != min(want_columns, CONTENT_BUCKET_CAP):
        return False, (f"the slide sits in {_plural(want_columns, 'column')} "
                       f"and the layout offers "
                       f"{_plural(got_columns, 'column')}")
    if got_boxes < want_blocks:
        return False, (f"the slide has {_plural(want_blocks, 'content block')} "
                       f"and the layout offers "
                       f"{_plural(got_boxes, 'box', 'boxes')}")
    if sig.get("title") and not lay.get("title"):
        return False, "the slide carries a heading and the layout has no title"
    return True, ""


def _offers(lay: dict) -> str:
    return (f"{_plural(int(lay.get('bodies') or 0), 'content box', 'content boxes')}"
            f" in {_plural(max(1, int(lay.get('columns') or 1)), 'column')}"
            + (" with a title" if lay.get("title") else " and no title"))


@dataclass
class Gap:
    """One thing the master is missing, and every slide that wants it."""
    label: str                  # the sentence a designer reads
    slides: list[int]           # zero-based indices
    source_layouts: list[str]   # what those slides were on in the submitted deck
    signature: dict
    closest: str | None = None  # nearest layout the master DOES offer
    closest_note: str = ""
    refused: int = 0            # slides a designer found no home for
    reviewed: int = 0           # slides that got an answer at all
    asked: int = 0              # slides whose OWN call produced that answer

    @property
    def places(self) -> int:
        return len(self.slides)


@dataclass
class Coverage:
    """The deck-level answer. `gaps` is the part that gets acted on."""
    slides: int
    by_rule: dict[str, int] = field(default_factory=dict)
    gaps: list[Gap] = field(default_factory=list)
    used_layouts: dict[str, int] = field(default_factory=dict)
    unused_layouts: list[str] = field(default_factory=list)
    reviewed: int = 0
    not_reviewed: int = 0
    review_ran: bool = False
    # Slides that WERE placed, on a layout their content does not fit. Separate
    # from `gaps`, because they are a different sentence to a designer: the
    # master has a layout for this slide and it is the wrong one, as against the
    # master has nothing for this slide at all.
    misfits: list = field(default_factory=list)
    # The same misfits grouped by what they need, so one proposal answers all of
    # them (misfit_gaps). Kept beside the per-slide list rather than instead of
    # it: the list is the evidence a designer opens, the clusters are what gets
    # built.
    misfit_clusters: list = field(default_factory=list)

    @property
    def unplaced(self) -> int:
        """Slides still on a layout nobody chose for them."""
        return self.by_rule.get("fallback", 0) + self.by_rule.get("none", 0)

    @property
    def matched(self) -> int:
        return self.slides - self.unplaced


# --- what a slide is asking for -------------------------------------------


def _boxes(slide) -> list[tuple]:
    """Top-level items with their slide-space boxes.

    Top level, because a group is ONE thing to place: a card with its icon and
    its label is a single block the rebuild has to find a home for, and counting
    its three members as three blocks is how a two-card slide starts looking
    like it needs six boxes."""
    return [(p.shape, p.box) for p in placed_shapes(slide)
            if not p.grouped and p.box is not None]


def _columns(centres: list[float], tolerance: float = COLUMN_TOLERANCE) -> int:
    """How many columns these horizontal centres sit in.

    Clustered rather than counted: a stack of four blocks down the left of a
    slide is one column, and four blocks across it are four."""
    if not centres:
        return 0
    columns = 0
    last = None
    for c in sorted(centres):
        if last is None or c - last > tolerance:
            columns += 1
        last = c
    return columns


def signature(slide, slide_w: int, slide_h: int) -> dict:
    """What the rebuild will actually have to place on this slide.

    Handed to the vision pass as its prompt payload AND used here to cluster,
    which is deliberate: the model is shown the same reading of the slide that
    the report groups by, so a designer comparing the two is not comparing two
    different descriptions of the same slide.
    """
    area = float(slide_w) * float(slide_h) or 1.0
    title = False
    kinds = {"text": 0, "image": 0, "chart": 0, "table": 0, "group": 0}
    centres: list[float] = []
    placeholders = 0

    for shape, box in _boxes(slide):
        left, top, right, bottom = box
        width, height = right - left, bottom - top
        if width <= 0 or height <= 0:
            continue
        share = (width * height) / area
        kind = content_type(shape)

        try:
            if shape.is_placeholder:
                placeholders += 1
                token = str(shape.placeholder_format.type or "")
                if "TITLE" in token.upper():
                    title = True
                    continue
        except Exception:
            pass

        if kind is None:
            continue
        if share >= BACKDROP_AREA_SHARE:
            continue                      # a ground, not a block of content
        if share < MIN_CONTENT_AREA_SHARE:
            continue                      # page furniture

        if (kind == "text" and not title and top / slide_h <= TITLE_BAND
                and height / slide_h <= TITLE_MAX_HEIGHT_SHARE):
            title = True
            continue

        kinds[kind] = kinds.get(kind, 0) + 1
        centres.append(((left + right) / 2) / slide_w)

    blocks = sum(kinds.values())
    return {"title": title, "blocks": blocks, "columns": _columns(centres),
            "text": kinds["text"], "images": kinds["image"],
            "charts": kinds["chart"], "tables": kinds["table"],
            "groups": kinds["group"], "placeholders": placeholders}


def _plural(n: int, one: str, many: str | None = None) -> str:
    return f"{n} {one}" if n == 1 else f"{n} {many or one + 's'}"


def describe(sig: dict) -> str:
    """The signature as a sentence, which is what the report is read in."""
    parts = []
    if sig["columns"] > 1:
        parts.append(_plural(sig["columns"], "column"))
    blocks = []
    for key, word in (("text", "text block"), ("images", "picture"),
                      ("charts", "chart"), ("tables", "table"),
                      ("groups", "grouped element")):
        if sig.get(key):
            blocks.append(_plural(sig[key], word))
    if blocks:
        parts.append(", ".join(blocks))
    elif not parts:
        parts.append("no content blocks")
    head = "a title over " if sig["title"] else "no title, "
    if sig["columns"] > 1:
        return head + " of ".join(parts)
    return head + parts[-1]


def cluster_key(sig: dict) -> tuple:
    """What makes two slides the same request.

    Coarse on the counts and exact on the structure: a master needs to offer a
    title or not, a number of columns, and roughly how many blocks. Whether a
    slide has three text blocks or four does not change which layout has to
    exist."""
    return (
        sig["title"],
        min(sig["columns"], CONTENT_BUCKET_CAP),
        min(sig["blocks"], CONTENT_BUCKET_CAP),
        bool(sig["images"]),
        bool(sig["charts"] or sig["tables"]),
    )


# --- what the master offers instead ---------------------------------------

_BODYISH = ("body", "obj", "pic", "chart", "tbl", "dgm", "media", "clipArt")


def layout_signature(entry: dict, slide_w: int) -> dict:
    """The same reading, taken from a layout's placeholders rather than a
    slide's shapes. Like for like is the only way a comparison between the two
    means anything."""
    title = False
    centres = []
    bodies = 0
    for ph in entry.get("placeholders") or []:
        token = (ph.get("type") or "").lower()
        if "title" in token or token == "ctrtitle":
            title = True
            continue
        if token in ("ftr", "sldnum", "dt"):
            continue            # page furniture the master carries on every slide
        if token and not any(token.startswith(b.lower()) for b in _BODYISH):
            continue
        bodies += 1
        pos = ph.get("position_emu") or {}
        left, width = pos.get("left"), pos.get("width")
        if left is not None and width:
            centres.append(((left + width / 2) / slide_w))
    return {"title": title, "bodies": bodies, "columns": _columns(centres)}


def _closest(sig: dict, layouts: list[dict], slide_w: int):
    """The layout a designer would reach for if they had to use this master as
    it stands, and why it is not right.

    Reported because "the master has nothing for this" is only actionable next
    to what it does have. The distance is deliberately crude: it exists to pick
    a nearest neighbour for the sentence, never to place a slide - placing is
    qc.applymaster's job and it already ran."""
    best, best_score, best_sig = None, None, None
    for entry in layouts:
        lay = layout_signature(entry, slide_w)
        score = (2 * abs(lay["columns"] - sig["columns"])
                 + abs(lay["bodies"] - min(sig["blocks"], CONTENT_BUCKET_CAP))
                 + (0 if lay["title"] == sig["title"] else 1))
        if best_score is None or score < best_score:
            best, best_score, best_sig = entry, score, lay
    if best is None:
        return None, ""
    title_note = ""
    if best_sig["title"] and not sig["title"]:
        title_note = ", and a title these slides do not have"
    elif sig["title"] and not best_sig["title"]:
        title_note = ", and no title for the one these slides carry"
    note = (f"offers "
            f"{_plural(best_sig['bodies'], 'content box', 'content boxes')}"
            f" in {_plural(best_sig['columns'], 'column')}{title_note}. "
            f"This wants "
            f"{_plural(min(sig['blocks'], CONTENT_BUCKET_CAP), 'block')} in "
            f"{_plural(sig['columns'], 'column')}.")
    return best.get("name"), note


# --- the report -----------------------------------------------------------


def _cluster(items, label_of, sig_of, source_of) -> dict:
    """Group anything by its structural signature. One implementation, because a
    gap and a misfit are the same kind of fact and must not group two ways."""
    buckets: dict[tuple, Gap] = {}
    for item in items:
        sig = sig_of(item)
        key = cluster_key(sig)
        gap = buckets.get(key)
        if gap is None:
            gap = buckets[key] = Gap(label=label_of(sig), slides=[],
                                     source_layouts=[], signature=sig)
        gap.slides.append(item.slide_index)
        source = source_of(item)
        if source and source not in gap.source_layouts:
            gap.source_layouts.append(source)
    return buckets


def misfit_gaps(items: list) -> list:
    """Misfits grouped into the same shape as a gap, so the proposal pass needs
    no second code path.

    `closest` is the layout they are ON rather than the nearest one they are
    not: for a misfit that IS the useful comparison, because the designer is
    looking at a layout that exists and deciding whether to fix it or add
    another.
    """
    buckets = _cluster(items, lambda sig: describe(sig),
                       lambda m: m.signature or {}, lambda m: m.layout)
    out = []
    for key, gap in buckets.items():
        members = [m for m in items if cluster_key(m.signature or {}) == key]
        first = members[0]
        gap.closest = first.layout
        gap.closest_note = f"{_capitalise(first.offers)}, and {first.reason}."
        gap.refused = sum(1 for m in members if m.review == "no fit")
        gap.reviewed = sum(1 for m in members
                           if m.review and m.review != "not reviewed")
        gap.asked = gap.reviewed
        out.append(gap)
    out.sort(key=lambda g: (-g.places, g.label))
    return out


def _capitalise(text: str) -> str:
    return (text[:1].upper() + text[1:]) if text else text


def misfits(deck_prs, master_layouts: list[dict], plans: list) -> list:
    """Every slide that was placed on a layout its content does not fit.

    Only slides that WERE placed: a fallback is already reported as having no
    home and would be counted twice. And only against layouts that offer boxes
    (MIN_BOXES_TO_JUDGE), because a layout with none cannot be the wrong shape.
    """
    slide_w = int(deck_prs.slide_width or 0) or 1
    slide_h = int(deck_prs.slide_height or 0) or 1
    slides = list(deck_prs.slides)
    by_name = {(l.get("name") or "").strip().casefold(): l
               for l in master_layouts}

    out = []
    for plan in plans:
        if plan.match_rule in ("fallback", "none") or not plan.target_layout:
            continue
        if plan.slide_index >= len(slides):
            continue
        entry = by_name.get(plan.target_layout.strip().casefold())
        if entry is None:
            continue
        try:
            sig = signature(slides[plan.slide_index], slide_w, slide_h)
        except Exception:
            continue
        lay = layout_signature(entry, slide_w)
        ok, why = fits(sig, lay)
        if ok:
            continue
        review = getattr(plan, "review", "") or ""
        if review == "confirmed":
            # A model looked at this slide against this layout and said it
            # belongs. A signature comparison does not get to overrule that:
            # the whole reason for looking is that the shapes on a slide are not
            # the whole story about whether it fits.
            continue
        out.append(Misfit(
            slide_index=plan.slide_index, layout=plan.target_layout,
            rule=plan.match_rule, label=describe(sig), offers=_offers(lay),
            signature=sig, reason=why, review=review))
    return out


def report(deck_prs, master_layouts: list[dict], plans: list) -> Coverage:
    """Roll the format pass's plans up into the deck-level answer.

    Takes the OPEN presentation and the plans rather than bytes, because both
    callers already have them: the format route plans before it applies, and the
    coverage pass plans without applying anything at all. Re-reading the deck
    here would be a second read that could disagree with the first."""
    slide_w = int(deck_prs.slide_width or 0) or 1
    slide_h = int(deck_prs.slide_height or 0) or 1
    slides = list(deck_prs.slides)

    cov = Coverage(slides=len(plans))
    for plan in plans:
        cov.by_rule[plan.match_rule] = cov.by_rule.get(plan.match_rule, 0) + 1
        if plan.target_layout:
            cov.used_layouts[plan.target_layout] = (
                cov.used_layouts.get(plan.target_layout, 0) + 1)
        review = getattr(plan, "review", "") or ""
        if review:
            cov.review_ran = True
        if review and review != "not reviewed":
            cov.reviewed += 1

    cov.unused_layouts = sorted(
        entry["name"] for entry in master_layouts
        if entry.get("name") and entry["name"] not in cov.used_layouts)

    # Only the slides nobody chose a layout for. A name match is not a gap, and
    # neither is a slide the review placed: something in the master fit it.
    open_plans = [p for p in plans if p.match_rule in ("fallback", "none")]
    cov.not_reviewed = sum(1 for p in open_plans
                           if (getattr(p, "review", "") or "") in
                           ("", "not reviewed"))

    buckets: dict[tuple, Gap] = {}
    for plan in open_plans:
        if plan.slide_index >= len(slides):
            continue
        try:
            sig = signature(slides[plan.slide_index], slide_w, slide_h)
        except Exception:
            # A slide this pass cannot read is not silently dropped from the
            # count: it goes in its own bucket, labelled, so the totals still
            # add up to the deck.
            sig = {"title": False, "blocks": 0, "columns": 0, "text": 0,
                   "images": 0, "charts": 0, "tables": 0, "groups": 0,
                   "placeholders": 0, "unreadable": True}
        key = cluster_key(sig)
        gap = buckets.get(key)
        if gap is None:
            gap = buckets[key] = Gap(label=describe(sig), slides=[],
                                     source_layouts=[], signature=sig)
        gap.slides.append(plan.slide_index)
        if plan.source_layout and plan.source_layout not in gap.source_layouts:
            gap.source_layouts.append(plan.source_layout)
        review = getattr(plan, "review", "") or ""
        if review == "no fit":
            gap.refused += 1
        if review and review != "not reviewed":
            gap.reviewed += 1
            # Every decision is now its own. A vision pass used to answer one
            # slide and carry that answer to every slide built the same way, so
            # "looked at" and "took a look-alike's answer" had to be counted
            # apart; a designer answers each slide on the layout page, so the
            # two counts are the same count (31/08/2026).
            gap.asked += 1

    cov.misfits = misfits(deck_prs, master_layouts, plans)
    cov.misfit_clusters = misfit_gaps(cov.misfits)

    for gap in buckets.values():
        gap.closest, gap.closest_note = _closest(gap.signature, master_layouts,
                                                 slide_w)
    cov.gaps = sorted(buckets.values(), key=lambda g: (-g.places, g.label))
    return cov


def headline(cov: Coverage) -> str:
    """One sentence for the top of a page or the top of an email."""
    if not cov.slides:
        return "This deck has no slides."
    if not cov.unplaced:
        if cov.misfits:
            return (f"Every slide has a layout in this master, but "
                    f"{_plural(len(cov.misfits), 'slide')} sit on one their "
                    f"content does not fit.")
        return (f"Every one of the {cov.slides} slides has a layout in this "
                f"master.")
    biggest = cov.gaps[0] if cov.gaps else None
    lead = (f"{_plural(cov.unplaced, 'slide')} of {cov.slides} have no layout "
            f"in this master")
    if biggest and biggest.places > 1:
        return (f"{lead}, and {biggest.places} of them want the same thing: "
                f"{biggest.label}.")
    return lead + "."
