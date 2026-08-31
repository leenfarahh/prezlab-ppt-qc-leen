"""The editable shape of a profile: what a designer may change, and in what
units.

A profile is a JSON document and it was only ever writable two ways - read a
master and let qc.bootstrap infer one, or accept a triage answer that bumped one
value (qc.assist). Neither lets anyone change the thing they actually want to
change, which is usually one number: the body size is 17 not 18, this navy
belongs in the palette, the client's margins are tighter than ours. The answer
to all three was "read the master again", which rebuilds the whole profile and
throws away every earlier decision.

ONE SPEC DRIVES BOTH DIRECTIONS. The form is rendered from FIELDS and parsed
from FIELDS, so a field cannot appear on the page without being readable back or
be readable back without appearing. The alternative - a template listing inputs
and a parser listing keys - is two lists that agree on the day they are written.

UNITS ARE THE POINT OF THIS FILE. A profile stores EMU because that is what
OOXML measures in, and 457200 is not a number anybody can check. The form shows
inches and points, converts on the way in and on the way out, and round-trips:
saving a profile you did not edit must leave the file byte-identical in every
field this module owns. That is what `test_profiles` asserts, and it is the
whole reason the conversion lives in one place.

VALIDATION REFUSES, IT DOES NOT REPAIR. A hex code that is not six hex digits, a
size that is not a number, a margin wider than the slide - each comes back as a
message against that field with the rest of the form preserved. Silently
coercing a bad value is how a profile ends up with a 0pt body font that nobody
typed.
"""

import copy
import re
from dataclasses import dataclass, field as dc_field

EMU_IN = 914400
EMU_PT = 12700

_HEX = re.compile(r"^[0-9A-Fa-f]{6}$")

# Roles the font rules are stated per. Fixed rather than read off the profile:
# a role the audit modules do not know about is a rule nothing enforces, and a
# form that lets one be invented produces exactly that.
FONT_ROLES = ("title", "subtitle", "body", "caption")
WEIGHTS = ("light", "regular", "semibold", "bold")


@dataclass
class Field:
    """One input. `path` is dotted into profile["config"], except for the
    handful prefixed "/" which address the profile document itself."""
    path: str
    label: str
    kind: str            # text | int | number | inches | points | bool | select
    help: str = ""
    options: tuple = ()
    minimum: float | None = None
    maximum: float | None = None
    placeholder: str = ""

    @property
    def name(self) -> str:
        """The form field name. Dots are legal in an HTML name but awkward
        everywhere else, so they travel as double underscores."""
        return self.path.replace("/", "top:").replace(".", "__")


@dataclass
class Group:
    title: str
    blurb: str
    fields: list = dc_field(default_factory=list)


def _font_fields(role: str) -> list:
    return [
        Field(f"font.roles.{role}.latin", "Latin families", "text",
              "In order of preference, separated by commas. The first is what "
              "a fix applies; the rest are accepted without comment.",
              placeholder="Georgia, Times New Roman"),
        Field(f"font.roles.{role}.complex_script", "Arabic families", "text",
              "The complex-script face. A deck with Arabic runs is audited "
              "against this, never against the Latin list.",
              placeholder="Dubai, Noto Naskh Arabic"),
        Field(f"font.roles.{role}.size_pt", f"{role.title()} size (pt)",
              "number", minimum=1, maximum=400),
        Field(f"font.roles.{role}.allowed_weights", "Allowed weights",
              "select", options=WEIGHTS,
              help="Anything else on this role is flagged."),
    ]


FIELDS: list = [
    Group("Identity", "What this profile is called and who it is for.", [
        Field("/name", "Profile name", "text",
              "Shown in every picker. Say the client, not the file."),
        Field("/client_scope", "Client", "text",
              "Optional. Free text, for finding it later."),
        Field("/project_scope", "Project", "text", "Optional."),
        Field("/is_default", "Offer this one first", "bool",
              "Only one profile can be the default; setting this clears it "
              "on the others."),
    ]),
    Group("Type", "The families and sizes each role is held to. A run in the "
                  "wrong family is an error; a size outside the tolerance is a "
                  "warning.", [
        Field("font.theme_font_refs_allowed",
              "Accept theme font references", "bool",
              "A run set to +mj-lt rather than a family name resolves through "
              "the theme. Off means the family has to be stated."),
        Field("font.size_tolerance_pt", "Size tolerance (pt)", "number",
              "How far off the stated size a run may be before it is "
              "reported.", minimum=0, maximum=12),
    ] + [f for role in FONT_ROLES for f in _font_fields(role)]),
    Group("Palette", "The colours this client's work is allowed to use. "
                     "Anything else on a shape is measured against these and "
                     "reported by how far off it is.", [
        Field("color_palette.on_palette_mode", "Match colours", "select",
              options=("by_name", "by_hex"),
              help="by_name allows a tint or shade of a named colour; by_hex "
                   "requires the exact value."),
        Field("color_palette.match_tolerance_deltaE",
              "On-palette within (deltaE)", "number",
              "Perceptual distance. Under 1 is invisible; 2 is the usual "
              "working tolerance.", minimum=0, maximum=100),
        Field("color_palette.auto_replace_max_deltaE",
              "Auto-correct up to (deltaE)", "number",
              "A colour this close to a palette entry is snapped to it "
              "without asking. Above it, you are asked.", minimum=0,
              maximum=100),
        Field("color_palette.ambiguity_band_deltaE",
              "Ask rather than guess up to (deltaE)", "number",
              "Past this, a colour is treated as deliberate and only "
              "reported.", minimum=0, maximum=100),
    ]),
    Group("Margins and alignment", "The frame content sits inside, and how "
                                   "close counts as aligned. Shown in inches "
                                   "and points; stored as EMU.", [
        Field("geometry.safe_zone_margins_emu.left", "Left margin (in)",
              "inches", minimum=0, maximum=20),
        Field("geometry.safe_zone_margins_emu.right", "Right margin (in)",
              "inches", minimum=0, maximum=20),
        Field("geometry.safe_zone_margins_emu.top", "Top margin (in)",
              "inches", minimum=0, maximum=20),
        Field("geometry.safe_zone_margins_emu.bottom", "Bottom margin (in)",
              "inches", minimum=0, maximum=20),
        Field("geometry.grid.enabled", "Hold content to a grid", "bool"),
        Field("geometry.grid.columns", "Grid columns", "int", minimum=1,
              maximum=48),
        Field("geometry.grid.gutter_emu", "Gutter (in)", "inches", minimum=0,
              maximum=5),
        Field("geometry.alignment.edge_tolerance_emu", "Edge tolerance (pt)",
              "points", "Two edges within this are treated as intended to "
                        "line up.", minimum=0, maximum=72),
        Field("geometry.alignment.center_tolerance_emu",
              "Centre tolerance (pt)", "points", minimum=0, maximum=72),
        Field("geometry.alignment.spacing_tolerance_emu",
              "Spacing tolerance (pt)", "points", minimum=0, maximum=72),
    ]),
    Group("Shapes", "When two shapes are meant to be the same size.", [
        Field("shape_size.size_tolerance_emu", "Size tolerance (pt)", "points",
              minimum=0, maximum=72),
        Field("shape_size.min_cohort_size", "Smallest group to judge", "int",
              "Below this many similar shapes, nothing is reported: two "
              "shapes of different sizes are not evidence of a cohort.",
              minimum=2, maximum=20),
        Field("shape_size.preserve_picture_aspect",
              "Never distort a picture", "bool",
              "A resize that would change a picture's aspect ratio is held "
              "back for your approval."),
        Field("shape_size.dominant_size_strategy", "Size to settle on",
              "select", options=("median", "mode", "largest", "smallest")),
    ]),
    Group("Master and layouts", "How strictly a deck has to stay on the "
                                "master's own layouts.", [
        Field("master_slide.enforce_existing_only",
              "Only the master's own layouts", "bool",
              "On, a slide on a layout the master does not define is "
              "reported."),
        Field("master_slide.layout_allowlist", "Allowed layouts", "text",
              "Comma-separated layout names. Empty means all of them."),
        Field("master_slide.geometry_tolerance_emu",
              "Placeholder tolerance (pt)", "points",
              "How far a placeholder may sit from where the layout puts it.",
              minimum=0, maximum=72),
    ]),
    Group("Page furniture", "The footer, the slide number and the date.", [
        Field("header_footer.template.footer_text", "Footer text", "text",
              "Empty means the deck sets its own."),
        Field("header_footer.template.slide_number", "Slide numbers", "bool"),
        Field("header_footer.template.date.enabled", "Date", "bool"),
        Field("header_footer.template.date.format", "Date format", "select",
              options=("DD/MM/YYYY", "D MMMM YYYY", "MM/DD/YYYY",
                       "YYYY-MM-DD")),
        Field("header_footer.template.font_role", "Furniture type role",
              "select", options=FONT_ROLES),
    ]),
]


def groups() -> list:
    return FIELDS


def _dig(node, path: str):
    for key in path.split("."):
        if not isinstance(node, dict) or key not in node:
            return None
        node = node[key]
    return node


def _plant(node: dict, path: str, value) -> None:
    keys = path.split(".")
    for key in keys[:-1]:
        nxt = node.get(key)
        if not isinstance(nxt, dict):
            nxt = {}
            node[key] = nxt
        node = nxt
    node[keys[-1]] = value


def _round(value: float) -> float:
    """A float with its float noise trimmed. 0.5000000000000001 in a form field
    reads as a bug in the tool, and it is one."""
    rounded = round(value, 4)
    return int(rounded) if rounded == int(rounded) else rounded


def _stored(profile: dict, f: Field):
    """This field's value as the profile holds it, before any unit conversion."""
    if f.path.startswith("/"):
        return profile.get(f.path[1:])
    return _dig(profile.get("config") or {}, f.path)


def display(profile: dict, f: Field):
    """What goes in the input for this field, from the profile as stored."""
    raw = _stored(profile, f)

    if f.kind == "bool":
        return bool(raw)
    if raw is None:
        return [] if f.kind == "select" and f.path.endswith("allowed_weights") \
            else ""
    if f.kind == "inches":
        return _round(float(raw) / EMU_IN)
    if f.kind == "points":
        return _round(float(raw) / EMU_PT)
    if isinstance(raw, list):
        return raw if f.path.endswith("allowed_weights") else ", ".join(
            str(x) for x in raw)
    return raw


def multiple(f: Field) -> bool:
    """Whether this select takes several answers. Only the weight sets do, and
    they are the reason `select` is not simply a dropdown."""
    return f.path.endswith("allowed_weights")


class Invalid(ValueError):
    """One field the designer has to fix, named."""

    def __init__(self, field_name: str, message: str):
        super().__init__(message)
        self.field_name = field_name


def _number(f: Field, raw: str, cast):
    text = (raw or "").strip()
    if text == "":
        raise Invalid(f.name, f"{f.label} needs a value.")
    try:
        value = cast(text)
    except ValueError:
        raise Invalid(f.name, f"{f.label} has to be a number, not "
                              f"{text!r}.") from None
    if f.minimum is not None and value < f.minimum:
        raise Invalid(f.name, f"{f.label} cannot be below {_round(f.minimum)}.")
    if f.maximum is not None and value > f.maximum:
        raise Invalid(f.name, f"{f.label} cannot be above {_round(f.maximum)}.")
    return value


def _unchanged(f: Field, text: str, current) -> bool:
    """Whether this input still holds exactly what was rendered into it.

    THE ROUND TRIP TURNS ON THIS. A margin of 274638 EMU is 0.3003543 inches,
    and no form shows a designer seven decimal places - so the field renders
    0.3004, and converting that back lands on 274594. Forty-four EMU is nothing
    to look at and it is not nothing to do: every open-and-save walks the value,
    and a profile saved a few times has margins nobody set.

    So a field the designer did not touch is not rewritten. The comparison is
    against what was RENDERED, not against a tolerance: type 0.3004 yourself
    into an untouched field and it stays 274638, which is the honest reading of
    "you left it alone". Type 0.31 and it moves.
    """
    if current is None:
        return False
    try:
        return float(text) == float(display_number(f, current))
    except (TypeError, ValueError):
        return False


def display_number(f: Field, raw):
    """The number this field shows for a stored value. Split out of display()
    so the parser can ask the same question the renderer answered."""
    if f.kind == "inches":
        return _round(float(raw) / EMU_IN)
    if f.kind == "points":
        return _round(float(raw) / EMU_PT)
    return raw


def _parse(f: Field, form, current=None) -> object:
    if f.kind == "bool":
        return bool(form.get(f.name))
    if multiple(f):
        picked = [w for w in form.getlist(f.name) if w in f.options]
        if not picked:
            raise Invalid(f.name, f"{f.label} needs at least one weight, "
                                  f"or nothing on that role is allowed.")
        return picked

    raw = form.get(f.name)
    raw = "" if raw is None else str(raw)

    if f.kind == "select":
        text = raw.strip()
        if text not in f.options:
            raise Invalid(f.name, f"{f.label} is not one of "
                                  f"{', '.join(f.options)}.")
        return text
    if f.kind == "int":
        return _number(f, raw, int)
    if f.kind == "number":
        return _number(f, raw, float)
    if f.kind in ("inches", "points"):
        inches = _number(f, raw, float)     # validated before it is trusted
        if _unchanged(f, raw.strip(), current):
            return current
        scale = EMU_IN if f.kind == "inches" else EMU_PT
        return int(round(inches * scale))

    text = raw.strip()
    # A comma-separated field is a list in the profile even when it holds one
    # entry, because the modules iterate it. An empty one is an empty list, not
    # a list holding "".
    if (f.path.endswith(("latin", "complex_script", "layout_allowlist"))):
        return [part.strip() for part in text.split(",") if part.strip()]
    if f.path.startswith("/") or f.path.endswith("footer_text"):
        return text or None
    return text


def parse_colors(form) -> list:
    """The palette rows, as named_colors entries.

    Rows arrive as parallel lists (every colour_name[] beside its colour_hex[])
    because the form lets a designer add and remove rows in the browser, so the
    indices are not contiguous by the time it posts. A row with neither a name
    nor a hex is a row they emptied out and is dropped; a row with one of the
    two is a mistake and is named.
    """
    names = form.getlist("color_name")
    hexes = form.getlist("color_hex")
    out, seen = [], set()
    for i, (name, hexval) in enumerate(zip(names, hexes)):
        name = (name or "").strip()
        hexval = (hexval or "").strip().lstrip("#").upper()
        if not name and not hexval:
            continue
        if not name:
            raise Invalid("color_name", f"The colour {hexval} has no name. "
                                        f"Name it or clear the row.")
        if not _HEX.match(hexval):
            raise Invalid("color_hex", f"{name} has {hexval!r} for a hex code. "
                                       f"It needs six hex digits, like "
                                       f"1F4E79.")
        key = name.casefold()
        if key in seen:
            raise Invalid("color_name", f"Two colours are both called "
                                        f"{name!r}. Names are how a fix picks "
                                        f"one, so they have to differ.")
        seen.add(key)
        out.append({"name": name, "hex": hexval, "theme_ref": None,
                    "allowed_tints": [], "allowed_shades": []})
    return out


def apply_form(profile: dict, form) -> dict:
    """A NEW profile dict with the form's values written into it.

    A copy, never in place: a save that fails validation halfway would
    otherwise leave the loaded profile half-edited, and the page re-renders
    from it.

    Everything this module does not own is carried through untouched. A profile
    holds keys no form should be able to reach - the id, the stored master's
    filename, whatever a later version adds - and a save that rewrote the
    document from the form alone would delete them.
    """
    out = copy.deepcopy(profile)
    config = out.setdefault("config", {})

    for group in FIELDS:
        for f in group.fields:
            value = _parse(f, form, _stored(profile, f))
            if f.path.startswith("/"):
                out[f.path[1:]] = value
            else:
                _plant(config, f.path, value)

    palette = config.setdefault("color_palette", {})
    palette["named_colors"] = parse_colors(form)

    name = (out.get("name") or "").strip()
    if not name:
        raise Invalid("top:name", "A profile needs a name.")
    out["name"] = name
    out["version"] = int(profile.get("version") or 1) + 1
    return out


def partial(profile: dict, form) -> dict:
    """The profile with everything the form got RIGHT written into it.

    What the page re-renders after a validation failure. A designer who changed
    nine fields and mistyped the tenth has to get their nine back; re-rendering
    the stored document instead would throw the edit away and blame them for it.

    Every field is attempted independently and a bad one is left at its stored
    value, which is exactly where the error message points. Nothing is written
    to disk from here - it is a view, not a save.
    """
    out = copy.deepcopy(profile)
    config = out.setdefault("config", {})
    for group in FIELDS:
        for f in group.fields:
            try:
                value = _parse(f, form, _stored(profile, f))
            except Invalid:
                continue
            if f.path.startswith("/"):
                out[f.path[1:]] = value
            else:
                _plant(config, f.path, value)
    try:
        config.setdefault("color_palette", {})["named_colors"] = \
            parse_colors(form)
    except Invalid:
        # The rows as typed, so the offending one is still on screen to fix.
        # Unvalidated on purpose: this dict is rendered, never stored.
        names, hexes = form.getlist("color_name"), form.getlist("color_hex")
        config.setdefault("color_palette", {})["named_colors"] = [
            {"name": (n or "").strip(),
             "hex": (h or "").strip().lstrip("#").upper()}
            for n, h in zip(names, hexes) if (n or "").strip()
            or (h or "").strip()]
    return out


def summary(profile: dict) -> str:
    """One line for the list page: what this profile actually holds someone to.

    Read off the config rather than stored, so it cannot go stale against the
    values underneath it."""
    config = profile.get("config") or {}
    body = _dig(config, "font.roles.body.latin") or []
    colors = _dig(config, "color_palette.named_colors") or []
    left = _dig(config, "geometry.safe_zone_margins_emu.left")
    bits = []
    if body:
        bits.append(f"{body[0]} body")
    if colors:
        bits.append(f"{len(colors)} palette colour"
                    f"{'s' if len(colors) != 1 else ''}")
    if left:
        bits.append(f"{_round(left / EMU_IN)}in margins")
    return " &middot; ".join(bits) or "no rules set"
