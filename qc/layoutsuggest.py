"""What layout the master is missing, drawn so a designer can build it.

qc.layoutgap says a master has no home for eleven slides and clusters them: six
want a title over two columns, five want a full-bleed statement. That is the
diagnosis. This is the prescription, and it is the half a designer acts on -
"add a two-column comparison layout" is a sentence they can take to the master,
where "eleven slides fell back" is a sentence they have to interpret first.

THE MODEL NAMES THE STRUCTURE. CODE PUTS IT ON THE MASTER'S OWN FRAME. The model
gets the gap - what the slides in it hold, how many columns, whether they carry a
title - and answers with a layout described structurally: a name, an archetype,
and a list of boxes by KIND and COLUMN. It never returns a coordinate. The
geometry comes from the master's own stated frame divided into the columns it
asked for, so a proposed layout lands on the client's margins rather than on
numbers a model invented (same division of labour as qc.components).

AND IT PROPOSES, IT DOES NOT BUILD. Nothing here writes to the master. A client's
master is not a file this tool edits on a hunch, and a layout is a design object
with type styles, guides and brand furniture on it - the parts a designer adds
and this pass has no opinion about. What comes back is a wireframe, a name and a
reason, next to the slides that would use it.

Every proposal is checked against the gap it was asked about before it is shown:
a layout with one body box does not answer a gap that wants two columns, however
well it is described. A proposal that does not serve its gap is discarded rather
than shown, for the reason qc.layoutpick drops an invented layout name.
"""

from dataclasses import dataclass, field

from .llm import ask_json

EMU_IN = 914400

# The OOXML archetype tokens a proposal may claim. Closed because the token is
# what the format pass matches on (qc.applymaster.plan_assignments), so an
# invented one would produce a layout that never matches anything.
ARCHETYPES = ("title", "obj", "twoObj", "twoTxTwoObj", "secHead", "titleOnly",
              "objTx", "picTx", "tbl", "chart", "blank", "objOnly")

# What a proposed box can be. Mapped to the placeholder types PowerPoint has, so
# a designer building it is picking from the same menu.
KINDS = {
    "title": "Title",
    "body": "Content or text",
    "picture": "Picture",
    "chart": "Chart",
    "table": "Table",
    "caption": "Text (caption)",
}

MAX_BOXES = 8
MAX_COLUMNS = 4
# One call per gap, and a deck rarely has more than a handful. Past this the
# report stops being a list of things to build and becomes a redesign.
MAX_SUGGESTIONS = 6

# The gutter between columns when the master's grid does not state one. 0.25in
# is the narrowest gap that still reads as a gap at presentation size.
DEFAULT_GUTTER = EMU_IN // 4
# A title band deep enough for two lines at display size.
TITLE_BAND = int(EMU_IN * 1.05)
# Air between the title and the content under it.
TITLE_GAP = EMU_IN // 4

SPEC_SCHEMA = {
    "type": "object", "additionalProperties": False,
    "required": ["name", "archetype", "columns", "boxes", "why"],
    "properties": {
        "name": {"type": "string"},
        "archetype": {"type": "string"},
        "columns": {"type": "integer"},
        "boxes": {
            "type": "array",
            "items": {
                "type": "object", "additionalProperties": False,
                "required": ["kind", "column"],
                "properties": {
                    "kind": {"type": "string"},
                    # 0 means "across every column", which is what a title is.
                    "column": {"type": "integer"},
                    "label": {"type": "string"},
                },
            },
        },
        "why": {"type": "string"},
    },
}

_SYSTEM = """You are a senior presentation designer at Prezlab. A client's master
has no layout for a group of slides, and you are naming the layout it should
have.

You are told what those slides hold - how many content blocks, in how many
columns, with or without a title - and which layouts the master already has. You
answer with the layout to ADD, described structurally.

Give it a name a designer would recognise in PowerPoint's layout gallery, two to
four words, in the client's register rather than a description of the mechanics.
"Two-column comparison" or "Statement" or "Metrics row", not "Layout with two
content placeholders".

`archetype` is the OOXML token, from the list you are given and nothing else. It
is what the format pass matches on, so an invented token produces a layout that
never gets used.

`boxes` is the placeholders it needs, each a `kind` from the list you are given
and a `column`: 1 for the first column, 2 for the second, and 0 for a box that
spans every column - which is what a title does. Do not give coordinates or
sizes. The layout is placed on the master's own frame by code, so it lands on
the client's margins.

Match the group you were shown. A layout with one content box does not serve a
group that wants two columns, and a proposal that does not fit its group is
thrown away.

`why` is one sentence of US English, no em dashes, saying what this layout is
for. A designer reads it to decide whether to build it."""


@dataclass
class Suggestion:
    """One layout to add, and everything a designer needs to decide."""
    name: str
    archetype: str
    columns: int
    boxes: list          # [{"kind", "column", "label", "box": (l, t, r, b)}]
    why: str
    serves: list = field(default_factory=list)   # slide indices, zero-based
    gap_label: str = ""
    collides_with: str | None = None   # an existing layout of the same name

    @property
    def places(self) -> int:
        return len(self.serves)


# --- the master's own frame ------------------------------------------------


def frame_of(prs, master, space=None) -> tuple:
    """The rectangle a layout's content belongs inside, in EMU.

    The master's stated presentation space when it has one, because that is the
    designer's own answer to this question and it outranks anything inferred.
    Failing that, the slide inset by half an inch - not a guess at the brand's
    margins, just somewhere honest to draw a wireframe.
    """
    if space:
        box = space.get("box_emu") if isinstance(space, dict) else space
        if box and len(box) == 4 and box[2] > box[0] and box[3] > box[1]:
            return tuple(int(v) for v in box)
    width = int(prs.slide_width or 0) or int(EMU_IN * 13.333)
    height = int(prs.slide_height or 0) or int(EMU_IN * 7.5)
    inset = EMU_IN // 2
    return (inset, inset, width - inset, height - inset)


def _place(spec: dict, frame: tuple, gutter: int = DEFAULT_GUTTER) -> list:
    """Turn a structural description into real boxes on the master's frame.

    Deliberately dumb: columns of equal width, content filling the height under
    the title, stacked where a column holds more than one box. A designer moves
    these; what matters is that the wireframe shows the STRUCTURE at the client's
    margins rather than at invented coordinates.
    """
    left, top, right, bottom = frame
    columns = max(1, min(int(spec.get("columns") or 1), MAX_COLUMNS))
    boxes = spec.get("boxes") or []

    has_title = any(b.get("column") == 0 for b in boxes)
    body_top = top + (TITLE_BAND + TITLE_GAP if has_title else 0)
    col_width = (right - left - gutter * (columns - 1)) // columns

    per_column: dict[int, list] = {}
    for box in boxes:
        column = int(box.get("column") or 0)
        per_column.setdefault(column, []).append(box)

    placed = []
    for box in per_column.get(0, []):
        placed.append(dict(box, box=(left, top, right, top + TITLE_BAND)))

    for index in range(1, columns + 1):
        members = per_column.get(index) or []
        if not members:
            continue
        col_left = left + (index - 1) * (col_width + gutter)
        share = (bottom - body_top - gutter * (len(members) - 1)) // len(members)
        for n, box in enumerate(members):
            box_top = body_top + n * (share + gutter)
            placed.append(dict(box, box=(col_left, box_top,
                                         col_left + col_width, box_top + share)))
    return placed


# --- validation -----------------------------------------------------------


def _serves(spec: dict, signature: dict) -> bool:
    """Whether this proposal answers the gap it was asked about.

    Columns have to match, because that is the shape of the request. The box
    count is allowed to differ by one: a group whose slides carry three blocks
    in two columns is well served by a two-column layout, and insisting on three
    boxes would reject the right answer.
    """
    want_columns = max(1, int(signature.get("columns") or 1))
    got_columns = max(1, min(int(spec.get("columns") or 1), MAX_COLUMNS))
    if got_columns != min(want_columns, MAX_COLUMNS):
        return False

    boxes = spec.get("boxes") or []
    content = [b for b in boxes if int(b.get("column") or 0) > 0]
    want_blocks = max(1, int(signature.get("blocks") or 1))
    if abs(len(content) - want_blocks) > 1 and len(content) < want_blocks:
        return False

    if signature.get("title") and not any(int(b.get("column") or 0) == 0
                                          for b in boxes):
        return False
    return True


def validate(spec: dict, signature: dict, existing: dict) -> dict | None:
    """A proposal cleaned up, or None when it cannot be used.

    `existing` maps a normalised layout name to the master's real name, so a
    proposal that repeats one is not offered as new - a designer told to add a
    layout the master already has stops trusting the page. It is reported as a
    collision instead, because that IS the finding: the layout exists and the
    slides did not reach it, which is a naming problem rather than a missing
    layout.
    """
    name = " ".join(str(spec.get("name") or "").split())[:60]
    if not name:
        return None
    archetype = str(spec.get("archetype") or "").strip()
    if archetype not in ARCHETYPES:
        return None

    boxes = []
    for box in (spec.get("boxes") or [])[:MAX_BOXES]:
        kind = str(box.get("kind") or "").strip().lower()
        if kind not in KINDS:
            continue
        column = int(box.get("column") or 0)
        if not (0 <= column <= MAX_COLUMNS):
            continue
        boxes.append({"kind": kind, "column": column,
                      "label": " ".join(str(box.get("label") or "").split())[:40]})
    if not boxes:
        return None

    cleaned = {"name": name, "archetype": archetype,
               "columns": max(1, min(int(spec.get("columns") or 1), MAX_COLUMNS)),
               "boxes": boxes,
               "why": " ".join(str(spec.get("why") or "").split())[:240]}
    if not _serves(cleaned, signature):
        return None
    cleaned["collides_with"] = existing.get(_norm(name))
    return cleaned


def _norm(name: str) -> str:
    return " ".join((name or "").split()).casefold()


# --- the call -------------------------------------------------------------


class Unreachable:
    """The model could not be ASKED about this gap.

    Distinct from None, which _ask never returns for that case any more, and
    distinct from a validate() rejection. The three were one outcome until
    30/08/2026 and the page said the same sentence about all of them: "a
    proposal that did not answer the group it was asked about is discarded
    rather than shown" - which asserts the model answered. Under a 429 it never
    answered, and telling a designer their proposal was rejected on quality when
    the truth is an exhausted quota sends them to fix the wrong thing.
    """

    def __init__(self, reason: str):
        self.reason = reason


def _ask(gap, existing_names: list, signature: dict):
    """The answer, or an Unreachable saying why it could not be asked."""
    import json

    try:
        return ask_json(
            system=_SYSTEM,
            prompt=(
                "The master already has these layouts:\n"
                + "\n".join(f"- {n}" for n in existing_names[:24])
                + "\n\nArchetype tokens you may use:\n"
                + ", ".join(ARCHETYPES)
                + "\n\nBox kinds you may use:\n"
                + ", ".join(f"{k} ({v})" for k, v in KINDS.items())
                + f"\n\nThe group of slides with no layout: {gap.label}. "
                  f"{gap.places} slide(s), and in the deck they were on: "
                + ", ".join(gap.source_layouts[:6] or ["(unknown)"])
                + ".\n\nWhat the slides in this group hold:\n"
                + json.dumps(signature, sort_keys=True)),
            schema=SPEC_SCHEMA,
            max_tokens=1024,
        )
    except Exception as exc:
        return Unreachable(f"{type(exc).__name__}: {exc}")


def suggest(coverage, prs, master_layouts: list, space=None,
            master=None) -> tuple[list, int, str]:
    """One layout proposal per gap, best first.

    Returns (suggestions, asked, unreachable_reason) - the last being "" when
    every call was actually made, and otherwise the first failure's reason.

    THE THIRD VALUE IS WHY THE PAGE CAN TELL THE TRUTH. Two things produce no
    proposal and they are opposite facts: the model answered and the answer did
    not serve the gap (a real, if disappointing, result), or the model was never
    reached at all. Both used to come out as "a proposal that did not answer the
    group it was asked about is discarded rather than shown", which under a 429
    tells a designer their proposal was rejected on quality when the truth is
    that no proposal was ever made.

    Only gaps get a proposal: a slide the master already places is not a
    question. Anything the model cannot answer, or answers in a way that does
    not serve the gap, simply produces no suggestion for that gap - the gap
    itself still stands in the report, which is the state before this pass ran.
    """
    # Unplaced slides first, then slides placed on a layout that does not fit.
    # Both are one missing layout stated several times, and the proposal is the
    # same shape either way (qc.layoutgap.misfit_gaps).
    gaps = [g for g in list(coverage.gaps or [])
            + list(getattr(coverage, "misfit_clusters", None) or [])
            if g.places][:MAX_SUGGESTIONS]
    if not gaps:
        return [], 0, ""

    existing = {_norm(l.get("name")): l.get("name") for l in master_layouts
                if l.get("name")}
    names = [l["name"] for l in master_layouts if l.get("name")]
    frame = frame_of(prs, master, space)

    out, asked, unreachable = [], 0, ""
    for gap in gaps:
        answer = _ask(gap, names, gap.signature or {})
        asked += 1
        if isinstance(answer, Unreachable):
            # The first reason is kept and the rest are not: they are the same
            # outage restated, and a page listing six copies of one 429 is
            # noise. Asking stops too - once the provider is refusing, five more
            # calls buy nothing.
            unreachable = unreachable or answer.reason
            break
        if answer is None:
            continue
        spec = validate(answer, gap.signature or {}, existing)
        if spec is None:
            continue
        out.append(Suggestion(
            name=spec["name"], archetype=spec["archetype"],
            columns=spec["columns"], boxes=_place(spec, frame),
            why=spec["why"], serves=list(gap.slides),
            gap_label=gap.label, collides_with=spec["collides_with"]))
    out.sort(key=lambda s: -s.places)
    return out, asked, unreachable


# --- the picture ----------------------------------------------------------

_FILL = {"title": "#0b3d42", "body": "#cfe8ea", "picture": "#e3ddf0",
         "chart": "#f6e0c8", "table": "#dfe6ec", "caption": "#eef3f4"}


def wireframe(suggestion: Suggestion, prs) -> str:
    """The proposed layout as an inline SVG, at the slide's own aspect ratio.

    A wireframe rather than a render, because there is nothing to render: this
    layout does not exist yet. What a designer needs to see is where the boxes
    sit relative to the frame they know, which is exactly what a wireframe
    shows and a list of placeholder types does not.
    """
    width = int(prs.slide_width or 0) or int(EMU_IN * 13.333)
    height = int(prs.slide_height or 0) or int(EMU_IN * 7.5)
    parts = [f'<svg viewBox="0 0 {width} {height}" '
             f'style="width:100%;height:auto;display:block;'
             f'background:#fff;border:1px solid #d8e0e2;border-radius:8px" '
             f'role="img" aria-label="Wireframe of {_esc(suggestion.name)}">']
    for box in suggestion.boxes:
        left, top, right, bottom = box["box"]
        fill = _FILL.get(box["kind"], "#eef3f4")
        label = box.get("label") or KINDS.get(box["kind"], box["kind"])
        # Type scaled off the slide width so it reads the same at any render
        # size, and clamped so a narrow column does not get illegible text.
        size = max(int(width * 0.018), 100000)
        parts.append(
            f'<rect x="{left}" y="{top}" width="{max(0, right - left)}" '
            f'height="{max(0, bottom - top)}" rx="45000" fill="{fill}" '
            f'stroke="#0b3d42" stroke-opacity="0.25" stroke-width="9525"/>')
        parts.append(
            f'<text x="{left + EMU_IN // 6}" y="{top + size + EMU_IN // 8}" '
            f'font-size="{size}" fill="#002528" '
            f'font-family="Segoe UI, Helvetica, Arial, sans-serif">'
            f'{_esc(label)}</text>')
    parts.append("</svg>")
    return "".join(parts)


def _esc(text: str) -> str:
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))
