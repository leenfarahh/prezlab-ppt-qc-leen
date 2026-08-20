"""Moving a deck's content into the master it was just given.

Applying a layout remaps placeholder content only; a deck of free-floating
shapes keeps them exactly where they were. These tests build that situation
directly (blank layout, everything a free text box) because it is what an
export-tool deck actually looks like, and it is the case that made an applied
master look like nothing had happened.
"""

import io

from pptx import Presentation
from pptx.util import Emu, Pt

from qc.migrate import migrate_deck
from qc.util import iter_shapes_deep

IN = 914400


# Header geometry copied from the real client master this was built against
# (its "USE" layout): a title band across the top and a thin subtitle strip
# under it. The stock template's layouts do not combine those - its title
# slide centres both mid-canvas - so the placeholders are positioned here
# explicitly rather than inheriting whatever the sample template happens to
# do. Getting this wrong is what made the first version of these tests fail:
# the content never overlapped a placeholder that sat in the middle.
_TITLE_BOX = (0.48, 0.42, 12.40, 0.92)
_SUBTITLE_BOX = (0.48, 1.40, 12.40, 0.35)


def _deck(*, with_furniture=False, tall_content=False):
    """A slide whose content is all free shapes, over a master-style header."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # ctrTitle + subTitle
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)

    def tb(x, y, w, h, text, size):
        shape = slide.shapes.add_textbox(Emu(int(x)), Emu(int(y)),
                                         Emu(int(w)), Emu(int(h)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(size)
        return shape

    tb(0.6 * IN, 0.35 * IN, 2 * IN, 0.3 * IN, "CORE FEATURES", 11)
    tb(2.4 * IN, 0.30 * IN, 7 * IN, 0.6 * IN, "Full Lifecycle Management", 30)
    tb(0.6 * IN, 1.05 * IN, 8 * IN, 0.4 * IN, "Every layer supports it.", 13)
    height = 6.0 if tall_content else 3.0
    for i in range(3):
        tb((0.6 + i * 4.1) * IN, 1.9 * IN, 3.8 * IN, height * IN, f"Card {i+1}", 14)
    if with_furniture:
        tb(0.5 * IN, 6.9 * IN, 1.5 * IN, 0.25 * IN, "Strategy&", 9)
        tb(12.7 * IN, 6.9 * IN, 0.4 * IN, 0.25 * IN, "1", 9)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _run(**kw):
    out, changes = migrate_deck(_deck(**kw))
    # migrate_deck catches per-slide exceptions so one bad slide cannot fail a
    # deck. That resilience once hid a NameError as "nothing moved", so every
    # test asserts the pass actually ran rather than silently skipped.
    skipped = [c for c in changes if c.action == "migration skipped"]
    assert not skipped, f"the migration raised: {skipped[0].detail}"
    return Presentation(io.BytesIO(out)).slides[0], changes


def _actions(changes):
    return [c.action for c in changes]


def _ph_text(slide, *want):
    """Placeholder text by type name. TITLE and CENTER_TITLE are both titles;
    a layout picks one and the migration handles either."""
    for ph in slide.placeholders:
        if str(ph.placeholder_format.type).startswith(want):
            return ph.text_frame.text
    return None


def _bottoms(deck_bytes):
    slide = Presentation(io.BytesIO(deck_bytes)).slides[0]
    return {s.shape_id: (s.top or 0) + (s.height or 0) for s in slide.shapes}


# ------------------------------------------------------------ text migration


def test_largest_type_in_the_title_band_becomes_the_title():
    slide, changes = _run()
    assert "title into placeholder" in _actions(changes)
    assert _ph_text(slide, "TITLE", "CENTER_TITLE") == "Full Lifecycle Management"
    # the heading is gone as a free shape, not duplicated
    free = [s.text_frame.text for s in slide.shapes
            if not s.is_placeholder and s.has_text_frame]
    assert "Full Lifecycle Management" not in free


def test_standfirst_reaches_the_subtitle_even_without_overlap():
    """A hand-laid deck rarely puts its standfirst exactly where the master's
    subtitle box sits, so overlap alone misses it."""
    slide, changes = _run()
    assert "subtitle into placeholder" in _actions(changes)
    assert _ph_text(slide, "SUBTITLE") == "Every layer supports it."


def test_the_eyebrow_is_not_mistaken_for_the_title():
    """CORE FEATURES sits in the title band too, but it is 11pt against 30pt."""
    slide, _ = _run()
    assert _ph_text(slide, "TITLE", "CENTER_TITLE") != "CORE FEATURES"


# --------------------------------------------------------------- the block


def test_remaining_content_moves_as_one_block():
    """Relative arrangement is the design; the cards must keep their spacing."""
    before = Presentation(io.BytesIO(_deck())).slides[0]
    card_lefts = sorted(s.left for s in before.shapes
                        if s.has_text_frame and s.text_frame.text.startswith("Card"))
    gaps_before = [b - a for a, b in zip(card_lefts, card_lefts[1:])]

    slide, _changes = _run()
    after = sorted(s.left for s in slide.shapes
                   if s.has_text_frame and s.text_frame.text.startswith("Card"))
    # Asserted on the outcome, not on a "content block moved" record: with the
    # eyebrow now swept, the body may already sit on the margin and need no
    # move at all. Spacing preserved is the invariant either way.
    assert [b - a for a, b in zip(after, after[1:])] == gaps_before


def test_migration_never_pushes_a_shape_further_off_the_canvas():
    """Regression: dragging footer-band shapes with the content block moved
    them to 8.45in on a 7.5in canvas, out of sight.

    Asserted as "no worse than before", not "everything fits": a source deck
    can arrive with content already overflowing, and this pass is not allowed
    to fix that by scaling. What it must never do is make it worse.

    Scoped to a master that states no body ceiling, which is the case here (the
    stock template draws no guides). Where a master DOES draw one, that line
    binds and the overflow it creates is reported instead - see
    test_the_ceiling_binds_even_when_the_block_then_overflows."""
    source = _deck(with_furniture=True, tall_content=True)
    height = Presentation(io.BytesIO(source)).slide_height
    before_overflow = max(
        [b - height for b in _bottoms(source).values() if b > height] or [0])

    out, _ = migrate_deck(source)
    after_overflow = max(
        [b - height for b in _bottoms(out).values() if b > height] or [0])

    assert after_overflow <= before_overflow
    slide = Presentation(io.BytesIO(out)).slides[0]
    assert all((s.top or 0) >= 0 for s in slide.shapes)


def test_footer_furniture_is_never_dragged_downward():
    """The specific shape of the original bug: a footer at 6.9in must not end
    up below the canvas because the content above it moved."""
    source = _deck(with_furniture=False, tall_content=False)
    prs = Presentation(io.BytesIO(source))
    slide = prs.slides[0]
    tb = slide.shapes.add_textbox(Emu(int(0.5 * IN)), Emu(int(6.9 * IN)),
                                  Emu(int(3 * IN)), Emu(int(0.25 * IN)))
    tb.text_frame.paragraphs[0].add_run().text = "Source: internal analysis 2026"
    buf = io.BytesIO()
    prs.save(buf)

    out, _ = migrate_deck(buf.getvalue())
    after = Presentation(io.BytesIO(out)).slides[0]
    note = [s for s in after.shapes
            if s.has_text_frame and s.text_frame.text.startswith("Source:")]
    assert note, "the footnote must survive"
    assert note[0].top == Emu(int(6.9 * IN)), "left exactly where it was"


def test_oversized_content_is_reported_not_scaled():
    """Shrinking a text box does not shrink its type, so a silent scale would
    produce overflowing text that looks fine in the XML."""
    slide, changes = _run(tall_content=True)
    assert "content does not fit" in _actions(changes)
    cards = [s for s in slide.shapes
             if s.has_text_frame and s.text_frame.text.startswith("Card")]
    assert all(c.height == Emu(int(6.0 * IN)) for c in cards)


# ------------------------------------------------------------ housekeeping


def test_a_bare_page_number_duplicate_is_removed():
    """The slide's own hand-typed page number, where the master stamps one."""
    slide, changes = _run(with_furniture=True)
    removed = [c for c in changes if c.action == "removed duplicate furniture"]
    assert [c.detail.split()[0] for c in removed] == ["page"]
    text = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    assert "Strategy&" in text, "an unmatched footer string is content, not a duplicate"


def test_footer_band_shapes_are_parked_not_moved():
    """Anything already in the footer band is page furniture by position and
    stays there; only the content above it forms the moving block.

    Uses normal-height content on purpose: with oversized content the block
    already starts below the region top, so no move happens and there is
    nothing to park it against."""
    slide, _changes = _run(with_furniture=True, tall_content=False)
    footer = [s for s in slide.shapes
              if s.has_text_frame and s.text_frame.text == "Strategy&"]
    assert footer and footer[0].top == Emu(int(6.9 * IN)),         "a footer-band shape must never be dragged by the content block"


def test_no_empty_placeholder_survives():
    """An empty placeholder shows its prompt text in the editor, which is what
    made the applied master look like nothing had happened."""
    slide, _ = _run()
    for ph in slide.placeholders:
        if str(ph.placeholder_format.type).startswith(
                ("FOOTER", "SLIDE_NUMBER", "DATE")):
            continue  # furniture placeholders are meant to be empty
        assert ph.text_frame.text.strip(), "an empty placeholder was left behind"


def test_an_unfillable_placeholder_is_removed_and_reported():
    """A layout offering more placeholders than the slide has content for."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
    buf = io.BytesIO()
    prs.save(buf)

    out, changes = migrate_deck(buf.getvalue())
    assert [c.action for c in changes].count("removed empty placeholder") == 2
    assert not list(Presentation(io.BytesIO(out)).slides[0].placeholders)


def test_a_slide_with_nothing_to_move_is_left_alone():
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    buf = io.BytesIO()
    prs.save(buf)
    out, changes = migrate_deck(buf.getvalue())
    assert Presentation(io.BytesIO(out))
    assert not [c for c in changes if c.action.startswith(("title", "subtitle"))]


def test_every_change_names_its_slide_and_reads_as_a_sentence():
    _, changes = _run(with_furniture=True)
    for c in changes:
        assert c.slide_index == 0
        assert c.detail and c.action
        assert str(c).startswith("slide 1: ")


def test_a_grouped_title_is_still_found():
    """A converter or a designer may group the eyebrow with the heading. A
    group carries no text of its own, so a top-level-only scan never sees the
    shape that holds the title."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)

    group = slide.shapes.add_group_shape()
    for x, y, w, h, text, size in (
            (0.6, 0.35, 2.0, 0.3, "CORE FEATURES", 11),
            (2.4, 0.30, 7.0, 0.6, "Full Lifecycle Management", 30)):
        shape = group.shapes.add_textbox(Emu(int(x * IN)), Emu(int(y * IN)),
                                         Emu(int(w * IN)), Emu(int(h * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(size)
    buf = io.BytesIO()
    prs.save(buf)

    out, changes = migrate_deck(buf.getvalue())
    slide = Presentation(io.BytesIO(out)).slides[0]

    assert _ph_text(slide, "TITLE", "CENTER_TITLE") == "Full Lifecycle Management"
    lifted = [c for c in changes if c.action == "title into placeholder"]
    assert lifted and "lifted out of a group" in lifted[0].detail
    # the rest of the group survives
    remaining = [s.text_frame.text for s, _ in iter_shapes_deep(slide.shapes)
                 if s.has_text_frame]
    assert "CORE FEATURES" in remaining


# ------------------------------------------------------------- collisions


def _header_deck(*, heading_pt=28, eyebrow_pt=11, standfirst_pt=13,
                 content_to_bottom=True):
    """The real failure from the client deck: an eyebrow and a heading both
    inside the master's title band, a standfirst under them, and content that
    already reaches the bottom edge so the block cannot shift down."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)

    def tb(x, y, w, h, text, size=None):
        shape = slide.shapes.add_textbox(Emu(int(x * IN)), Emu(int(y * IN)),
                                         Emu(int(w * IN)), Emu(int(h * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        if size:
            run.font.size = Pt(size)
        return shape

    tb(0.6, 0.45, 3.0, 0.28, "TECHNICAL ARCHITECTURE", eyebrow_pt)
    tb(0.6, 0.62, 7.0, 0.55, "Under the hood", heading_pt)
    tb(0.6, 1.05, 9.0, 0.35, "A lean service boundary.", standfirst_pt)
    tb(0.6, 1.6, 11.0, 5.7 if content_to_bottom else 2.0, "DIAGRAM")
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _overlaps(slide):
    boxes = [(s, (s.left, s.top, s.left + s.width, s.top + s.height))
             for s in slide.shapes if s.left is not None]
    hits = []
    for i, (a, ba) in enumerate(boxes):
        for b, bb in boxes[i + 1:]:
            x = min(ba[2], bb[2]) - max(ba[0], bb[0])
            y = min(ba[3], bb[3]) - max(ba[1], bb[1])
            if x > 0 and y > 0:
                hits.append((a, b))
    return hits


def test_nothing_is_left_printing_over_a_filled_placeholder():
    """The defect from the client deck: the heading went into the title
    placeholder and the eyebrow stayed on top of it."""
    out, changes = migrate_deck(_header_deck())
    slide = Presentation(io.BytesIO(out)).slides[0]

    assert _ph_text(slide, "TITLE", "CENTER_TITLE") == "Under the hood"
    placeholders = [s for s in slide.shapes if s.is_placeholder]
    for a, b in _overlaps(slide):
        assert not (a in placeholders or b in placeholders), \
            "a shape is still printing over a placeholder"
    # The guarantee is that nothing overlaps a placeholder, not that a
    # NUDGE achieved it: header remnants are now placed under the header
    # band by the block pass, so a nudge is the fallback, not the norm.
    assert not any(c.action == "migration skipped" for c in changes)


def test_a_nudge_clears_the_whole_header_band_not_just_one_placeholder():
    """Shifting past the title alone dropped the eyebrow onto the subtitle,
    trading one collision for another."""
    out, changes = migrate_deck(_header_deck())
    slide = Presentation(io.BytesIO(out)).slides[0]

    # The eyebrow no longer needs nudging because it is swept as unplaced text.
    # What must hold either way: nothing free is left inside the header band.
    assert any(c.action == "removed unplaced text" for c in changes)
    floor = max(p.top + p.height for p in slide.placeholders
                if str(p.placeholder_format.type).startswith(
                    ("TITLE", "CENTER_TITLE", "SUBTITLE")))
    for shape in slide.shapes:
        if shape.is_placeholder or shape.top is None:
            continue
        assert shape.top + shape.height > floor,             f"{shape.text_frame.text!r} is still inside the header band"


def test_one_shape_gets_one_nudge_and_one_log_line():
    """Iterating placeholder-first moved a shape once per placeholder it hit."""
    _out, changes = migrate_deck(_header_deck())
    nudges = [c for c in changes if c.action == "nudged clear of a placeholder"]
    labels = [c.detail.split("'")[1] for c in nudges]
    assert len(labels) == len(set(labels)), f"a shape was nudged twice: {labels}"


def test_a_full_slide_block_is_never_consumed_into_the_subtitle():
    """The subtitle fallback CONSUMES the shape it picks. An unbounded version
    swallowed the whole diagram and deleted it."""
    out, changes = migrate_deck(_header_deck())
    slide = Presentation(io.BytesIO(out)).slides[0]

    assert _ph_text(slide, "SUBTITLE") == "A lean service boundary."
    text = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
    assert "DIAGRAM" in text, "the diagram must survive as content"


def test_with_no_explicit_sizes_the_higher_line_becomes_the_title():
    """A deck that sets no font sizes has every line resolving to the same
    inherited size, so they genuinely tie and rule 2 decides: the line nearer
    the top wins. Worth knowing as a limitation rather than a bug - the
    eyebrow, being uppermost, takes the title slot. A deck that sets real
    sizes ranks correctly (see the test above)."""
    out, _ = migrate_deck(_header_deck(heading_pt=None, eyebrow_pt=None,
                                       standfirst_pt=None))
    slide = Presentation(io.BytesIO(out)).slides[0]
    assert _ph_text(slide, "TITLE", "CENTER_TITLE") == "TECHNICAL ARCHITECTURE"


def test_duplicated_text_over_a_placeholder_is_removed():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)
    for y, size in ((0.5, 28), (0.75, 28)):
        shape = slide.shapes.add_textbox(Emu(int(0.6 * IN)), Emu(int(y * IN)),
                                         Emu(int(7 * IN)), Emu(int(0.5 * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = "Under the hood"
        run.font.size = Pt(size)
    buf = io.BytesIO()
    prs.save(buf)

    out, changes = migrate_deck(buf.getvalue())
    slide = Presentation(io.BytesIO(out)).slides[0]
    assert any(c.action == "removed duplicated text" for c in changes)
    kept = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
    assert kept.count("Under the hood") == 1


def test_content_on_content_overlap_is_reported_not_moved():
    """Choosing which of two content blocks gives way is a layout decision;
    the audit owns it. This pass must not guess, but must not hide it either."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tops = []
    for y in (3.0, 3.1):
        shape = slide.shapes.add_textbox(Emu(int(2 * IN)), Emu(int(y * IN)),
                                         Emu(int(5 * IN)), Emu(int(1 * IN)))
        shape.text_frame.paragraphs[0].add_run().text = f"Block at {y}"
        tops.append(y)
    buf = io.BytesIO()
    prs.save(buf)

    out, changes = migrate_deck(buf.getvalue())
    flagged = [c for c in changes if c.action == "text overlaps text"]
    assert flagged and "layout decision" in flagged[0].detail
    slide = Presentation(io.BytesIO(out)).slides[0]
    assert sorted(round(s.top / IN, 1) for s in slide.shapes) == tops


# --------------------------------------------------- background and bounds

_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _set_solid_bg(container, hexval):
    from lxml import etree

    cSld = container._element.find(f"{_P}cSld")
    for old in cSld.findall(f"{_P}bg"):
        cSld.remove(old)
    bg = etree.Element(f"{_P}bg")
    bgPr = etree.SubElement(bg, f"{_P}bgPr")
    solid = etree.SubElement(bgPr, f"{_A}solidFill")
    etree.SubElement(solid, f"{_A}srgbClr").set("val", hexval)
    etree.SubElement(bgPr, f"{_A}effectLst")
    cSld.insert(0, bg)


def _has_own_bg(container) -> bool:
    cSld = container._element.find(f"{_P}cSld")
    return bool(cSld is not None and cSld.findall(f"{_P}bg"))


def test_a_slide_background_override_is_dropped_so_the_master_shows():
    """An exported deck stamps a white background on every slide, and a
    slide-level background beats the master's. That is how a deck adopts
    every other part of a master and still comes out the wrong colour."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    _set_solid_bg(prs.slide_masters[0], "1E2E61")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_solid_bg(slide, "FFFFFF")
    buf = io.BytesIO()
    prs.save(buf)
    assert _has_own_bg(Presentation(io.BytesIO(buf.getvalue())).slides[0])

    out, changes = migrate_deck(buf.getvalue())
    slide = Presentation(io.BytesIO(out)).slides[0]
    assert not _has_own_bg(slide), "the override must be gone"
    assert any(c.action == "dropped background override" for c in changes)


def test_a_slide_keeps_its_background_when_the_master_declares_none():
    """Stripping an override with nothing to inherit would leave the slide
    worse off, not more on-brand."""
    prs = Presentation()
    master_cSld = prs.slide_masters[0]._element.find(f"{_P}cSld")
    for old in master_cSld.findall(f"{_P}bg"):
        master_cSld.remove(old)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_solid_bg(slide, "FFFFFF")
    buf = io.BytesIO()
    prs.save(buf)

    out, changes = migrate_deck(buf.getvalue())
    assert _has_own_bg(Presentation(io.BytesIO(out)).slides[0])
    assert not any(c.action == "dropped background override" for c in changes)


def test_the_block_move_never_pushes_content_over_the_footer():
    """Regression from the client deck: bullets ended up printing across the
    footer text because the move was bounded by the canvas, not the region."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)

    def tb(y, h, text, size=None):
        shape = slide.shapes.add_textbox(Emu(int(0.6 * IN)), Emu(int(y * IN)),
                                         Emu(int(9 * IN)), Emu(int(h * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        if size:
            run.font.size = Pt(size)
        return shape

    tb(0.5, 0.6, "A heading", 28)
    tb(1.1, 0.3, "A standfirst.", 13)
    # content that already runs close to the bottom
    tb(1.7, 4.9, "BULLETS", 14)
    footer = tb(6.95, 0.25, "Strategy&", 9)
    footer_top = footer.top
    buf = io.BytesIO()
    prs.save(buf)

    out, _ = migrate_deck(buf.getvalue())
    slide = Presentation(io.BytesIO(out)).slides[0]
    bullets = next(s for s in slide.shapes
                   if s.has_text_frame and s.text_frame.text == "BULLETS")
    assert bullets.top + bullets.height <= footer_top, \
        "content was pushed into the footer band"


# ------------------------------------------------- margins, not movements


def _margin_deck(*, eyebrow_top, body_top):
    """A client-shaped slide: eyebrow, heading, standfirst, then a content
    cluster. eyebrow_top and body_top vary independently so the body's final
    position can be tested for independence from what sits above it."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)

    def tb(x, y, w, h, text, size=None):
        shape = slide.shapes.add_textbox(Emu(int(x * IN)), Emu(int(y * IN)),
                                         Emu(int(w * IN)), Emu(int(h * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        if size:
            run.font.size = Pt(size)

    tb(0.6, eyebrow_top, 3.0, 0.28, "FUTURE WORK", 11)
    tb(0.6, 0.55, 9.0, 0.60, "Next: Personalized Daily Digests", 28)
    tb(0.6, 1.25, 10.0, 0.40, "A second LLM integration.", 13)
    for i in range(3):
        tb(0.6 + i * 4.1, body_top, 3.8, 1.9, f"Card {i + 1}", 12)
    tb(0.6, body_top + 2.1, 10.0, 0.8, "Bullets", 12)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _tops(deck_bytes, prefix):
    slide = Presentation(io.BytesIO(deck_bytes)).slides[0]
    return sorted(s.top for s in slide.shapes
                  if s.has_text_frame and s.text_frame.text.startswith(prefix))


def test_body_is_pulled_up_to_the_margin_not_only_pushed_down():
    """Content sitting below the master's body margin must rise to meet it.
    Only ever moving down could never close a gap, which is what left 1.5in of
    dead space under the subtitle on the client deck."""
    source = _margin_deck(eyebrow_top=0.30, body_top=3.30)
    before = _tops(source, "Card")
    out, changes = migrate_deck(source)
    after = _tops(out, "Card")

    assert after[0] < before[0], "the body should have moved UP"
    moved = [c for c in changes if c.action == "content block moved"]
    assert moved and "body now starts at" in moved[0].detail
    # and the report names the frame it was seated on, so a stale stored master
    # cannot look like a bug
    assert "placeholder extents" in moved[0].detail


def test_body_position_is_independent_of_what_sits_above_it():
    """The property behind 'margins, not movements': how far a leftover header
    shape had to travel must not displace the body at all."""
    high = migrate_deck(_margin_deck(eyebrow_top=0.30, body_top=3.30))[0]
    low = migrate_deck(_margin_deck(eyebrow_top=1.00, body_top=3.30))[0]

    assert _tops(high, "Card") == _tops(low, "Card")


def test_relative_arrangement_inside_the_body_still_holds():
    """Margins govern where the body STARTS; inside it, the designer's
    spacing is still the design."""
    source = _margin_deck(eyebrow_top=0.30, body_top=3.30)
    before_cards = _tops(source, "Card")
    before_gap = _tops(source, "Bullets")[0] - before_cards[0]

    out, _ = migrate_deck(source)
    after_cards = _tops(out, "Card")
    after_gap = _tops(out, "Bullets")[0] - after_cards[0]

    assert len(set(after_cards)) == 1, "the cards must stay on one line"
    assert after_gap == before_gap


def test_an_unplaced_header_line_is_swept_rather_than_parked_in_the_body():
    """Stacking a remnant above the body was the earlier behaviour; it is now
    removed and reported, so it can never end up sitting inside the content."""
    out, changes = migrate_deck(_margin_deck(eyebrow_top=0.30, body_top=3.30))
    slide = Presentation(io.BytesIO(out)).slides[0]

    swept = [c for c in changes if c.action == "removed unplaced text"]
    assert [c.removed_text for c in swept] == ["FUTURE WORK"]
    assert "FUTURE WORK" not in [s.text_frame.text for s in slide.shapes
                                 if s.has_text_frame]


# ------------------------------------------- largest is title, second is sub


def _header_slide(entries):
    """A slide with the master-style header band and the given text boxes,
    as (top_in, height_in, text, size_pt)."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)
    for top, height, text, size in entries:
        shape = slide.shapes.add_textbox(Emu(int(0.6 * IN)), Emu(int(top * IN)),
                                         Emu(int(9 * IN)), Emu(int(height * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        if size:
            run.font.size = Pt(size)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _placed(deck_bytes):
    out, changes = migrate_deck(deck_bytes)
    slide = Presentation(io.BytesIO(out)).slides[0]
    return (_ph_text(slide, "TITLE", "CENTER_TITLE"),
            _ph_text(slide, "SUBTITLE"), out, changes)


def test_largest_type_is_the_title_and_second_largest_the_subtitle():
    title, subtitle, _out, _ch = _placed(_header_slide([
        (0.45, 0.28, "SMALL EYEBROW", 11),
        (0.62, 0.60, "The Heading", 28),
        (1.20, 0.35, "The standfirst line", 16),
        (2.60, 3.00, "CARDS", 12),
    ]))
    assert title == "The Heading"
    assert subtitle == "The standfirst line"


def test_rank_ignores_source_order_entirely():
    """Same three lines declared smallest-first must rank identically."""
    a = _placed(_header_slide([
        (0.62, 0.60, "The Heading", 28), (0.45, 0.28, "EYEBROW", 11),
        (1.20, 0.35, "The standfirst", 16), (2.60, 3.00, "CARDS", 12)]))[:2]
    b = _placed(_header_slide([
        (0.45, 0.28, "EYEBROW", 11), (1.20, 0.35, "The standfirst", 16),
        (0.62, 0.60, "The Heading", 28), (2.60, 3.00, "CARDS", 12)]))[:2]
    assert a == b == ("The Heading", "The standfirst")


def test_equal_sizes_break_by_position_higher_wins():
    """The safety net: at the same type size the line nearer the top of the
    slide is the heading."""
    title, subtitle, _out, _ch = _placed(_header_slide([
        (0.85, 0.30, "lower down", 20),
        (0.45, 0.30, "nearer the top", 20),
        (2.60, 3.00, "CARDS", 12),
    ]))
    assert title == "nearer the top"
    assert subtitle == "lower down"


def test_a_chart_figure_lower_down_never_becomes_the_title():
    """Ranking is bounded to the header band. Taking the globally largest text
    would let a big number in a chart outrank the heading."""
    title, _sub, _out, _ch = _placed(_header_slide([
        (0.62, 0.60, "The Heading", 24),
        (3.50, 1.20, "98%", 72),
    ]))
    assert title == "The Heading"


# --------------------------------------- unplaced text is removed and flagged


def test_unplaced_header_text_is_removed_and_flagged_with_its_full_text():
    """Rule: text the master has no placeholder for is removed, but reported
    loudly and in full so a designer can put it back."""
    full = "CORE FEATURES / STRATEGY AND OPERATIONS"
    _title, _sub, out, changes = _placed(_header_slide([
        (0.45, 0.28, full, 11),
        (0.62, 0.60, "The Heading", 28),
        (1.20, 0.35, "The standfirst", 16),
        (2.60, 3.00, "CARDS", 12),
    ]))

    alerts = [c for c in changes if c.severity == "alert"]
    assert len(alerts) == 1
    assert alerts[0].action == "removed unplaced text"
    # Full text, not a preview: a designer must not have to retype it.
    assert alerts[0].removed_text == full
    assert str(alerts[0]).startswith("slide 1: !! ")

    slide = Presentation(io.BytesIO(out)).slides[0]
    assert full not in [s.text_frame.text for s in slide.shapes
                        if s.has_text_frame]


def test_body_content_is_never_removed_as_unplaced():
    """Only the header band is swept. Body content has a home and keeps it."""
    _title, _sub, out, changes = _placed(_header_slide([
        (0.62, 0.60, "The Heading", 28),
        (1.20, 0.35, "The standfirst", 16),
        (2.60, 1.00, "Body paragraph one", 12),
        (3.80, 1.00, "Body paragraph two", 12),
    ]))
    assert not [c for c in changes if c.action == "removed unplaced text"]
    kept = [s.text_frame.text for s in
            Presentation(io.BytesIO(out)).slides[0].shapes if s.has_text_frame]
    assert "Body paragraph one" in kept and "Body paragraph two" in kept


def test_the_report_surfaces_removals_before_routine_moves():
    from qc.ui_format import render_format_result

    _t, _s, _out, changes = _placed(_header_slide([
        (0.45, 0.28, "AN EYEBROW", 11),
        (0.62, 0.60, "The Heading", 28),
        (1.20, 0.35, "The standfirst", 16),
        (2.60, 3.00, "CARDS", 12),
    ]))
    html = render_format_result(deck_name="d.pptx", profile_name="p",
                                job_id="j", plans=[], errors={}, applied=1,
                                content_changes=changes)
    assert "were removed" in html
    assert "AN EYEBROW" in html
    # the alert banner must appear before the per-change table
    assert html.index("were removed") < html.index("What moved into the master")


# ---------------------------------------------------- left / right margins

_P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"


def _plant_guides(master, vertical_in=(), horizontal_in=()):
    """Guides as desktop PowerPoint stores them: eighths of a point from the
    top-left edge, orient omitted for vertical. python-pptx cannot draw them,
    so the format captured from a COM probe is planted directly."""
    import itertools

    from lxml import etree

    ext_lst = etree.SubElement(master._element, f"{_P}extLst")
    ext = etree.SubElement(ext_lst, f"{_P}ext")
    ext.set("uri", "{GUIDES}")
    lst = etree.SubElement(ext, f"{{{_P15}}}sldGuideLst")
    gid = itertools.count(1)
    for pos, horz in ([(v, False) for v in vertical_in]
                      + [(h, True) for h in horizontal_in]):
        g = etree.SubElement(lst, f"{{{_P15}}}guide")
        g.set("id", str(next(gid)))
        if horz:
            g.set("orient", "horz")
        if pos:
            g.set("pos", str(int(pos * 72 * 8)))


def _guided_deck(*, card_left=0.73, card_width=3.8, guides=(0.60, 12.73)):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    _plant_guides(prs.slide_masters[0], vertical_in=guides,
                  horizontal_in=(0.42, 7.05))
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)

    def tb(x, y, w, h, text, size=None):
        shape = slide.shapes.add_textbox(Emu(int(x * IN)), Emu(int(y * IN)),
                                         Emu(int(w * IN)), Emu(int(h * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        if size:
            run.font.size = Pt(size)

    tb(0.6, 0.55, 9.0, 0.6, "The Heading", 28)
    tb(0.6, 1.20, 9.0, 0.35, "The standfirst", 16)
    for i in range(3):
        tb(card_left + i * (card_width + 0.3), 2.2, card_width, 2.5,
           f"Card {i + 1}", 12)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _declared_margins(deck_bytes):
    """The left margin and right edge the master actually declares.

    Read back rather than hand-computed in inches: guides are stored in
    EIGHTHS OF A POINT, so 0.60in (43.2pt) quantises to 43.125pt and a literal
    Emu(0.6 * 914400) is ~950 EMU off. The format's granularity is the truth
    here, not the number typed into PowerPoint."""
    from qc.stylespec import dominant_master, infer_grid

    prs = Presentation(io.BytesIO(deck_bytes))
    grid = infer_grid(prs, dominant_master(prs))
    margins = grid["margins_emu"]
    return margins["left"], prs.slide_width - margins["right"]


def _card_bounds(deck_bytes):
    slide = Presentation(io.BytesIO(deck_bytes)).slides[0]
    cards = [s for s in slide.shapes
             if s.has_text_frame and s.text_frame.text.startswith("Card")]
    return (min(c.left for c in cards),
            max(c.left + c.width for c in cards))


def test_the_left_margin_binds_from_the_masters_guides():
    """Content indented inside the margin is pulled out to meet it. Correcting
    only breaches left the cards a few millimetres inside the title's own left
    edge, which reads as misalignment even though no margin was crossed."""
    source = _guided_deck(card_left=0.73)
    margin_left, margin_right = _declared_margins(source)
    out, changes = migrate_deck(source)
    left, right = _card_bounds(out)

    assert left == margin_left, "the block should sit on the left guide"
    assert right <= margin_right, "and stay inside the right guide"
    moved = [c for c in changes if c.action == "content block moved"]
    assert moved and "-0.13in" in moved[0].detail


def test_guides_beat_placeholder_extents_as_the_margin_source():
    """A designer's guides are a stated intention; placeholder extents are only
    an inference from where the title happens to sit."""
    wide_src = _guided_deck(guides=(0.60, 12.73))
    tight_src = _guided_deck(guides=(1.20, 12.13))

    assert _card_bounds(migrate_deck(wide_src)[0])[0] ==         _declared_margins(wide_src)[0]
    assert _card_bounds(migrate_deck(tight_src)[0])[0] ==         _declared_margins(tight_src)[0]
    # and the two really are different margins, so the source is being read
    assert _declared_margins(wide_src)[0] != _declared_margins(tight_src)[0]


def test_full_bleed_content_is_exempt_from_the_left_margin():
    """A band running edge to edge is not indented content, and dragging it to
    the margin would destroy the effect."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    _plant_guides(prs.slide_masters[0], vertical_in=(0.60, 12.73),
                  horizontal_in=(0.42, 7.05))
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)
    band = slide.shapes.add_textbox(Emu(0), Emu(int(2.5 * IN)),
                                    prs.slide_width, Emu(int(2 * IN)))
    band.text_frame.paragraphs[0].add_run().text = "FULL BLEED BAND"
    buf = io.BytesIO()
    prs.save(buf)

    out, _ = migrate_deck(buf.getvalue())
    shape = next(s for s in Presentation(io.BytesIO(out)).slides[0].shapes
                 if s.has_text_frame and s.text_frame.text == "FULL BLEED BAND")
    assert shape.left == 0, "full-bleed content must not be indented"


# ------------------------------------------------- the reserved header band
#
# A client master draws four horizontal guides: the page's top and bottom
# margins, and between them the floor its subtitle may not cross and the
# ceiling its body may not cross. The strip between those two stays empty on
# every slide, so the ceiling is where the body starts - not the top margin
# (that is where the PAGE starts) and not the header placeholder's floor (that
# is where one layout's title box happens to end).


def _banded_deck(*, card_top=1.4, card_height=2.0, ceiling=2.40):
    """A guided master whose body ceiling sits BELOW the header placeholder's
    floor plus its gap, so the two sources give different answers and the test
    can tell which one the block was seated on."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    _plant_guides(prs.slide_masters[0], vertical_in=(0.60, 12.73),
                  horizontal_in=(0.45, 1.65, ceiling, 6.80))
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        ph.text_frame.clear()
        kind = str(ph.placeholder_format.type)
        box = _TITLE_BOX if kind.startswith(("TITLE", "CENTER_TITLE")) else (
            _SUBTITLE_BOX if kind.startswith("SUBTITLE") else None)
        if box:
            ph.left, ph.top, ph.width, ph.height = (
                Emu(int(v * IN)) for v in box)

    def tb(x, y, w, h, text, size=None):
        shape = slide.shapes.add_textbox(Emu(int(x * IN)), Emu(int(y * IN)),
                                         Emu(int(w * IN)), Emu(int(h * IN)))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = text
        if size:
            run.font.size = Pt(size)

    tb(0.6, 0.55, 9.0, 0.6, "The Heading", 28)
    tb(0.6, 1.20, 9.0, 0.35, "The standfirst", 16)
    for i in range(3):
        tb(0.6 + i * 4.1, card_top, 3.8, card_height, f"Card {i + 1}", 12)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _declared_ceiling(deck_bytes):
    from qc.stylespec import dominant_master, infer_grid

    prs = Presentation(io.BytesIO(deck_bytes))
    return infer_grid(prs, dominant_master(prs))["body_top_emu"]


def test_the_body_lands_on_the_masters_ceiling_not_the_header_floor():
    """Both are real statements, and the tighter one wins - but the ceiling has
    to be consulted at all. Seeding from the top margin left the header
    placeholder's floor as the only thing holding the body down."""
    source = _banded_deck(card_top=1.4)
    out, changes = migrate_deck(source)

    assert _tops(out, "Card") == [_declared_ceiling(source)] * 3
    moved = [c for c in changes if c.action == "content block moved"]
    assert moved, "the block should have been seated on the ceiling"


def test_the_ceiling_binds_even_when_the_block_then_overflows():
    """The strip the master keeps clear stays clear on the busy slides too.
    Clamping the move at the bottom margin collapses to no move at all on
    exactly those slides, which is how a deck comes back with content sitting
    in the reserved strip on every full slide. The overflow is reported
    instead - loudly, because content past the slide edge will not print."""
    source = _banded_deck(card_top=1.4, card_height=6.0)
    out, changes = migrate_deck(source)

    assert _tops(out, "Card") == [_declared_ceiling(source)] * 3
    alerts = [c for c in changes if c.action == "content does not fit"]
    assert alerts and alerts[0].severity == "alert"
    assert "past the bottom margin" in alerts[0].detail
    assert "past the slide edge" in alerts[0].detail
    assert "still clear" in alerts[0].detail


def test_a_master_with_no_stated_ceiling_keeps_the_bottom_clamp():
    """Where the master draws no body ceiling there is no statement to honour,
    so the old restraint holds: an overflowing block is never pushed further
    down. (The unguided case is covered by
    test_migration_never_pushes_a_shape_further_off_the_canvas.)"""
    guided = _guided_deck()          # two horizontal guides: margins only
    assert _declared_ceiling(guided) is None


def test_content_wider_than_the_margins_is_reported_not_narrowed():
    """Narrowing a text box reflows its text, so a block that cannot fit
    between both margins is aligned left and flagged for a designer."""
    out, changes = migrate_deck(_guided_deck(card_left=0.6, card_width=4.6))
    alerts = [c for c in changes if c.action == "wider than the margins"]

    assert alerts and alerts[0].severity == "alert"
    assert "cannot sit inside both margins" in alerts[0].detail
    before, after = _card_bounds(_guided_deck(card_left=0.6, card_width=4.6)), _card_bounds(out)
    assert (after[1] - after[0]) == (before[1] - before[0]), "must not be narrowed"
