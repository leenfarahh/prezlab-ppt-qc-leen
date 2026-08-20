"""Stage 1 pages: submit a finished master, review the Style Spec it yields.

Rendering only; route logic lives in qc/web.py. Reuses the Prezlab brand
shell from qc/ui.py so this reads as one product with the audit flow.

The review surface exists because a Style Spec is only trustworthy if a
design lead can check it against the master they built, and nobody checks
raw JSON. Every extracted value is shown in the terms a designer would use
(colour swatches, point sizes, inches) with the EMU kept alongside for
whoever needs to debug a number. Where a value was INFERRED rather than
declared, the page says so, because those are the values worth doubting.
"""

from .ui import _shell, esc

EMU_PER_IN = 914400
EMU_PER_PT = 12700


def _in(emu) -> str:
    """EMU rendered in inches; designers think in inches, not EMU."""
    if emu is None:
        return "&mdash;"
    return f"{emu / EMU_PER_IN:.2f}in"


def _warn(message: str) -> str:
    return f'<div class="banner warn">{esc(message)}</div>' if message else ""


# ------------------------------------------------------------------ intake


def render_master_intake(message: str = "") -> str:
    body = f"""
<h1>Read a master.</h1>
<p class="sub">Upload the finished master slide. This reads the slide master, its
layouts, and the theme, and returns the visual system they declare: theme colours
and fonts, placeholder geometry per layout, page furniture, and the grid.
Slide content is never read, so a master with no slides is exactly what this
expects.</p>
{_warn(message)}
<form action="/master" method="post" enctype="multipart/form-data" id="f">
  <div class="drop" id="drop" tabindex="0" role="button"
       aria-label="Drop a .pptx master here or press Enter to browse">
    <strong>Drop the master .pptx here</strong> or click to browse
    <div class="hint">Read locally and deleted straight after. The extracted
    spec stays in memory on this machine.</div>
    <div class="file" id="fname" aria-live="polite"></div>
    <input type="file" name="master" id="master" accept=".pptx" required hidden>
  </div>
  <div class="actions">
    <button class="btn primary" id="go" type="submit" disabled>Read the master</button>
  </div>
</form>
<script>
const drop = document.getElementById('drop'), input = document.getElementById('master'),
      fname = document.getElementById('fname'), go = document.getElementById('go');
function arm() {{
  if (input.files.length) {{ fname.textContent = input.files[0].name; go.disabled = false; }}
}}
drop.addEventListener('click', () => input.click());
drop.addEventListener('keydown', e => {{ if (e.key === 'Enter' || e.key === ' ') input.click(); }});
input.addEventListener('change', arm);
document.getElementById('f').addEventListener('submit', () => {{
  go.disabled = true;
  showBusy('Reading ' + (input.files[0] ? input.files[0].name : 'master'),
           'Walking the slide master, every layout, and the theme part.');
}});
['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => {{
  e.preventDefault(); drop.classList.add('armed'); }}));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => {{
  e.preventDefault(); drop.classList.remove('armed'); }}));
drop.addEventListener('drop', e => {{
  if (e.dataTransfer.files.length) {{ input.files = e.dataTransfer.files; arm(); }} }});
</script>"""
    return _shell("Read a master", body)


# ------------------------------------------------------------------- review


# Styled inline rather than via the shared .swatch class: that rule is scoped
# as `.legend .swatch`, so outside a legend it contributes no width or height
# and an empty span collapses to a hairline.
_SWATCH = ("display:inline-block;width:13px;height:13px;border-radius:3px;"
           "margin-right:0.6rem;vertical-align:-2px;"
           "border:1px solid var(--line)")


def _swatches(colors: dict) -> str:
    if not colors:
        return '<p class="note">No theme colours found.</p>'
    cells = []
    for slot, hexval in colors.items():
        cells.append(
            f'<div style="min-width:7.2rem">'
            f'<span style="{_SWATCH};background:#{esc(hexval)}"></span>'
            f'<b style="font-size:0.8rem">{esc(slot)}</b><br>'
            f'<span class="note" style="font-family:monospace">#{esc(hexval)}</span></div>')
    return ('<div style="display:flex;flex-wrap:wrap;gap:0.9rem 1.2rem">'
            + "".join(cells) + "</div>")


def _fonts_block(fonts: dict) -> str:
    if not fonts:
        return ""
    rows = []
    for kind in ("major", "minor"):
        f = fonts.get(kind) or {}
        label = "Major (headings)" if kind == "major" else "Minor (body)"
        scripts = " &middot; ".join(
            f"{name}: <b>{esc(f.get(key) or '—')}</b>"
            for name, key in (("Latin", "latin"),
                              ("Complex script", "complex_script"),
                              ("East Asian", "east_asian"))
            if f.get(key) or key != "east_asian")
        rows.append(f"<tr><td>{label}</td><td>{scripts}</td></tr>")
    return f"<table class='w3'><tbody>{''.join(rows)}</tbody></table>"


def _text_styles_block(styles: dict) -> str:
    if not styles:
        return ""
    rows = []
    for role in ("title", "body", "other"):
        s = styles.get(role)
        if not s:
            continue
        size = f"{s['size_pt']:g}pt" if s.get("size_pt") else "&mdash;"
        face = esc(s.get("latin") or "—")
        cs = f" / {esc(s['complex_script'])}" if s.get("complex_script") else ""
        weight = " bold" if s.get("bold") else ""
        rows.append(f"<tr><td>{esc(role)}</td><td>{size}</td>"
                    f"<td>{face}{cs}{weight}</td></tr>")
    if not rows:
        return ""
    return ("<table class='w3'><thead><tr><th>Level</th><th>Size</th>"
            f"<th>Typeface</th></tr></thead><tbody>{''.join(rows)}</tbody></table>")


_GRID_SOURCE_NOTE = {
    "presentation_space": (
        "Read from the <b>presentation space</b> rectangle on the master. This "
        "is the strongest statement a master can make about where content "
        "lives: it is drawn, not interpreted, so nothing has to choose between "
        "several sets of margins."),
    "guides": ("Read from the drawing guides on the master. These are a stated "
               "intention, so they are the most trustworthy numbers here. To "
               "remove any doubt where a master carries several sets of "
               "margins, draw a rectangle around the content area, give it no "
               "fill and no line, and mark it either by naming it "
               "<b>Presentation space</b> or by setting its alt text to "
               "<code>ToolsToo_PS</code> (which is what ToolsToo writes when "
               "you set a presentation space with it): the read will use that "
               "instead."),
    "placeholders": ("<b>Inferred.</b> This master has no presentation-space "
                     "rectangle and no drawing guides, so the margins come "
                     "from the extent of its own content placeholders. Worth a "
                     "designer's eye before anything relies on them."),
}


def _grid_block(grid: dict) -> str:
    if not grid or not grid.get("source"):
        return ('<p class="note">No drawing guides, and no master placeholders '
                'to infer a content area from. Nothing downstream gets a grid '
                'from this master.</p>')
    m = grid["margins_emu"]
    g = grid["guides"]
    cols = (f"<b>{grid['columns']}</b> columns, "
            f"{_in(grid['gutter_emu'])} gutter" if grid.get("columns")
            else "<span class='note'>no even column grid detected</span>")
    # The body ceiling gets its own row because it is the line most decks
    # actually break: the top margin is where the page starts, this is where
    # CONTENT starts, and the strip above it is meant to stay empty.
    body = grid.get("body_top_emu")
    floor = grid.get("subtitle_floor_emu")
    if body and floor:
        band_row = (f"<tr><td>Body begins</td><td>{_in(body)} from the top, with "
                    f"a {_in(body - floor)} strip above it (subtitle floor at "
                    f"{_in(floor)}) that stays empty on every slide</td></tr>")
    elif body:
        band_row = (f"<tr><td>Body begins</td><td>{_in(body)} from the top; the "
                    f"master states no reserved strip above it</td></tr>")
    else:
        band_row = ("<tr><td>Body begins</td><td><span class='note'>not stated: "
                    "the master's horizontal guides do not name a single line "
                    "where content starts, so content is only held to the top "
                    "margin</span></td></tr>")
    space = grid.get("presentation_space")
    space_row = ""
    if space:
        where = ("on the master" if space.get("source") == "master"
                 else f"on {space.get('source')}")
        warn = ("" if not space.get("prints") else
                " <b>It has a fill or a line, so it will print on every "
                "slide.</b> Set it to no fill and no line.")
        # How it was declared, because the two markers are found and fixed in
        # different places: one is the shape's name, the other its alt text.
        how = (f" (alt text <code>{esc(space.get('alt_text') or '')}</code>)"
               if space.get("marker") == "alt_text" else " (named)")
        space_row = (f"<tr><td>Presentation space</td><td>"
                     f"'{esc(space.get('name') or '')}'{how} {esc(where)}: "
                     f"{_in(space['box_emu'][2] - space['box_emu'][0])} &times; "
                     f"{_in(space['box_emu'][3] - space['box_emu'][1])}"
                     f"{warn}</td></tr>")
        if space.get("source") != "master":
            space_row += ("<tr><td></td><td class='note'>Move it to the slide "
                          "master so every layout inherits it.</td></tr>")
        # A second marker that disagrees is named rather than silently losing:
        # the frame it states would govern every slide of every deck on this
        # profile, so "which rectangle won" has to be answerable.
        for rival in space.get("rivals") or []:
            box = rival["box_emu"]
            where_r = rival.get("source") or "the same container"
            space_row += (
                f"<tr><td></td><td class='note'>A second marker disagrees: "
                f"'{esc(rival['name'])}' on {esc(where_r)} states "
                f"{_in(box[2] - box[0])} &times; {_in(box[3] - box[1])} at "
                f"{_in(box[0])} from the left, {_in(box[1])} from the top. The "
                f"one above was used. Delete whichever is stale.</td></tr>")
    return f"""
<table class="w3"><tbody>
{space_row}
<tr><td>Margins</td><td>left {_in(m['left'])} &middot; right {_in(m['right'])}
    &middot; top {_in(m['top'])} &middot; bottom {_in(m['bottom'])}</td></tr>
{band_row}
<tr><td>Columns</td><td>{cols}</td></tr>
<tr><td>Guides</td><td>{len(g['vertical_emu'])} vertical,
    {len(g['horizontal_emu'])} horizontal</td></tr>
</tbody></table>
<p class="note">{_GRID_SOURCE_NOTE.get(grid['source'], '')}</p>"""


def _background_cell(bg: dict | None) -> str:
    """A background is a design decision, so it gets shown, not stringified.
    A picture background renders as an actual preview when the spec carries
    its bytes, which is the fastest way for a designer to confirm the right
    image was captured."""
    if not bg:
        return "<span class='note'>none declared; inherits or is empty</span>"

    kind = bg.get("kind")
    if kind == "solid":
        return (f'<span style="{_SWATCH};background:#{esc(bg["hex"])}"></span>'
                f'solid <span style="font-family:monospace">#{esc(bg["hex"])}</span>')

    if kind == "theme_ref":
        themed = bg.get("theme_fill_kind")
        note = ""
        if themed and themed != "solid":
            note = (f' <span class="banner warn" style="display:inline-block;'
                    f'margin:0;padding:0.1rem 0.5rem">theme fill is a '
                    f'{esc(themed)}, not a flat colour</span>')
        chip = (f'<span style="{_SWATCH};background:#{esc(bg["hex"])}"></span>'
                if bg.get("hex") else "")
        return (f'{chip}theme fill reference '
                f'<span class="note">(idx {esc(str(bg.get("idx")))})</span>{note}')

    if kind == "image":
        img = bg.get("image") or {}
        fill = bg.get("fill") or {}
        if img.get("unavailable"):
            return (f'picture background &mdash; <span class="banner warn" '
                    f'style="display:inline-block;margin:0;padding:0.1rem 0.5rem">'
                    f'{esc(img["unavailable"])}</span>')

        bits = [f"<b>{esc(fill.get('mode') or 'unspecified')}</b>"]
        if fill.get("alpha_pct") is not None:
            bits.append(f"{fill['alpha_pct']:g}% opacity")
        if fill.get("crop_pct"):
            bits.append("cropped")
        if fill.get("recolor"):
            bits.append("recoloured: " + esc(", ".join(fill["recolor"])))
        if fill.get("tile"):
            bits.append("tiled")

        px = img.get("px") or {}
        facts = (f"{esc(str(img.get('format') or '?')).upper()} "
                 f"{px.get('width', '?')}&times;{px.get('height', '?')}px, "
                 f"{(img.get('bytes') or 0) / 1024:.0f} KB")

        preview = ""
        if img.get("data_base64"):
            preview = (
                f'<div style="margin-top:0.5rem"><img alt="Background image" '
                f'src="data:{esc(img.get("content_type") or "image/png")};'
                f'base64,{img["data_base64"]}" '
                f'style="max-width:22rem;max-height:9rem;border-radius:8px;'
                f'border:1px solid var(--line)"></div>')
        elif img.get("embed_skipped"):
            preview = (f'<div class="note" style="margin-top:0.3rem">Bytes not '
                       f'carried in the spec: {esc(img["embed_skipped"])}. '
                       f'Stage 2 will need the master for this asset.</div>')

        return (f"picture background, {' &middot; '.join(bits)}"
                f"<br><span class='note'>{facts} &middot; "
                f"<span style='font-family:monospace'>"
                f"{esc((img.get('sha1') or '')[:16])}</span></span>{preview}")

    return f"<b>{esc(str(kind))}</b> fill "\
           "<span class='note'>(no single colour; captured as-is)</span>"


def _furniture_row(label: str, item: dict) -> str:
    if not item or not item.get("present"):
        return (f"<tr><td>{label}</td>"
                f"<td><span class='note'>not on the master</span></td></tr>")
    if item.get("field"):
        what = f"live <b>{esc(item['field'])}</b> field"
    elif item.get("text"):
        what = f"text: <b>{esc(item['text'])}</b>"
    else:
        what = "present, empty"
    pos = item.get("position_emu") or {}
    where = (f" <span class='note'>at {_in(pos.get('left'))}, "
             f"{_in(pos.get('top'))}</span>") if pos.get("left") is not None else ""
    return f"<tr><td>{label}</td><td>{what}{where}</td></tr>"


def _logo_block(brand: dict) -> str:
    logo = (brand or {}).get("logo")
    if not logo:
        return ('<p class="note">No logo found on the master or its layouts. '
                'Only embedded pictures are detected: a logo drawn as grouped '
                'vector shapes will not be found (known gap).</p>')
    pos = logo["position_emu"]
    where = {"master": "on the slide master, so every layout inherits it",
             "layouts": f"on {len(logo['layouts'])} layout(s): "
                        + esc(", ".join(logo["layouts"])),
             "slides": "stamped on slides, not on the master"}.get(logo["scope"], "")
    drift = ('<p class="banner warn">Position varies between occurrences, so '
             'there is no single home position to propagate.</p>'
             if logo["position_varies"] else "")
    return f"""
<table class="w3"><tbody>
<tr><td>Placement</td><td>{where}</td></tr>
<tr><td>Position</td><td>{_in(pos['left'])}, {_in(pos['top'])}</td></tr>
<tr><td>Size</td><td>{_in(pos['width'])} &times; {_in(pos['height'])}</td></tr>
<tr><td>Image</td><td><span style="font-family:monospace">
    {esc(logo['image_sha1'][:16])}</span></td></tr>
</tbody></table>{drift}"""


def _layout_bg_cell(bg: dict | None) -> str:
    """Compact background indicator for the layouts table. A layout with no
    p:bg inherits the master's, which is the common and correct case, so it
    reads as inherited rather than as missing."""
    if not bg:
        return "<span class='note'>inherits</span>"
    kind = bg.get("kind")
    if kind == "solid":
        return (f'<span style="{_SWATCH};background:#{esc(bg["hex"])}"></span>'
                f'<span style="font-family:monospace;font-size:0.75rem">'
                f'#{esc(bg["hex"])}</span>')
    if kind == "image":
        img = bg.get("image") or {}
        if img.get("data_base64"):
            return (f'<img alt="" src="data:'
                    f'{esc(img.get("content_type") or "image/png")};base64,'
                    f'{img["data_base64"]}" style="height:18px;width:32px;'
                    f'object-fit:cover;border-radius:3px;vertical-align:-4px;'
                    f'border:1px solid var(--line)"> picture')
        return "picture"
    if kind == "theme_ref":
        chip = (f'<span style="{_SWATCH};background:#{esc(bg["hex"])}"></span>'
                if bg.get("hex") else "")
        return f"{chip}theme fill"
    return esc(str(kind))


_TITLE_TOKENS = ("title", "ctrTitle")


def _title_declaration(lay: dict) -> tuple:
    """(size_pt, autofit) the layout declares for its title, either None when
    it inherits the master's."""
    for ph in lay["placeholders"]:
        if ph["type"] in _TITLE_TOKENS:
            return ph.get("size_pt"), ph.get("autofit")
    return None, None


def _title_size_note(layouts: list) -> str:
    """Whether the layouts agree about the title size.

    Asked because a designer asked it the other way round: "some titles have
    bigger font sizes than others, why is that?" Nothing resizes them - each
    slide inherits the size of whatever layout it was given, and this master's
    layouts disagree. That is a master-level defect, fixable once here rather
    than per deck (design lead, 20/08/2026)."""
    stated = {}
    shrink = []
    for lay in layouts:
        size, autofit = _title_declaration(lay)
        if size:
            stated.setdefault(size, []).append(lay["name"])
        if autofit == "normAutofit":
            shrink.append(lay["name"])
    notes = []
    if len(stated) > 1:
        listing = "; ".join(
            f"<b>{size:g}pt</b> on {esc(', '.join(names))}"
            for size, names in sorted(stated.items()))
        notes.append(
            f"The layouts declare <b>{len(stated)} different title sizes</b> "
            f"({listing}). A slide's title takes the size of whichever layout "
            f"it is given, so decks built on this master will show titles at "
            f"different sizes with nothing having resized them. The layouts "
            f"that state no size inherit the master's.")
    if shrink:
        notes.append(
            f"{len(shrink)} layout(s) put the title on <b>shrink text on "
            f"overflow</b> ({esc(', '.join(shrink))}), so a long title renders "
            f"smaller again at whatever scale PowerPoint picks.")
    # The listings are built with their own markup and their names escaped, so
    # these go out raw rather than through _warn, which escapes.
    return "".join(f'<div class="banner warn">{n}</div>' for n in notes)


_FRAME_SLACK_EMU = 36000    # 1mm: a box a designer placed exactly on the line
_CONTENT_PH = ("title", "ctrTitle", "subTitle", "body")
_HEADER_PH = ("title", "ctrTitle", "subTitle")


def _space_agreement_note(grid: dict, layouts: list) -> str:
    """Whether the presentation space and the master's own placeholders agree
    about where content starts.

    They have to, and nothing can make them. A slide's title and body live in
    the master's placeholders, whose geometry IS the master's statement and
    which this tool never moves; free content is seated on the frame. If the
    rectangle says 1.20in and the title box says 0.48in, a formatted deck shows
    headers on one line and body on another - "the presentation space was
    copied but not applied", which is exactly how it reads (design lead,
    21/08/2026). One of the two has to move, in the master.

    A rectangle drawn around the BODY is the exception, and only for its top
    edge: the header lives above such a frame by design, so a title sitting
    above it agrees with the master rather than contradicting it. Its sides and
    its floor are still compared, because those are page-wide statements
    whatever the frame is drawn around (qc.stylespec.infer_grid space_states)."""
    space = (grid or {}).get("presentation_space")
    if not space or space.get("problem"):
        return ""
    states_body = (grid or {}).get("space_states") == "body"
    sl, st, sr, sb = space["box_emu"]
    # Worst deviation per placeholder role per side, plus how many layouts show
    # it. One line per role reads; one line per layout per side does not.
    worst: dict = {}
    for lay in layouts:
        for ph in lay["placeholders"]:
            if ph["type"] not in _CONTENT_PH:
                continue
            box = ph.get("position_emu") or {}
            if None in (box.get("left"), box.get("top"), box.get("width"),
                        box.get("height")):
                continue
            l, t = box["left"], box["top"]
            r, b = l + box["width"], t + box["height"]
            for side, over in (("left of it", sl - l),
                               ("past its right edge", r - sr),
                               ("above it", st - t),
                               ("below it", b - sb)):
                if over <= _FRAME_SLACK_EMU:
                    continue
                if (side == "above it" and states_body
                        and ph["type"] in _HEADER_PH):
                    continue
                key = (ph["type"], side)
                seen, layout_names = worst.get(key, (0, []))
                worst[key] = (max(seen, over), layout_names + [lay["name"]])
    if not worst:
        return ""
    lines = "".join(
        f"<li>The <b>{esc(kind)}</b> placeholder reaches up to "
        f"<b>{over / 914400:.2f}in</b> {esc(side)}, on {len(names)} layout(s): "
        f"{esc(', '.join(sorted(set(names))[:4]))}"
        f"{f' and more' if len(set(names)) > 4 else ''}.</li>"
        for (kind, side), (over, names) in sorted(
            worst.items(), key=lambda kv: -kv[1][0]))
    return f"""
<div class="banner warn">
  <b>The presentation space and this master's own placeholders disagree.</b>
  A slide's title and body sit in the master's placeholders, and formatting
  never moves those; free content is seated on the rectangle. Where the two
  differ, a formatted deck shows headers on one line and body on another, and
  no tool can reconcile it - one of the two has to move, here in the master.
  <ul style="margin:0.5rem 0 0 1.1rem">{lines}</ul>
</div>"""


def _layout_rows(layouts: list) -> str:
    rows = []
    for lay in layouts:
        phs = lay["placeholders"]
        explicit = sum(1 for p in phs if p["geometry_source"] == "explicit")
        kinds = ", ".join(sorted({p["type"] or "?" for p in phs
                                  if p["type"] not in ("ftr", "sldNum", "dt")}))
        geom = (f"{explicit} explicit / {len(phs) - explicit} inherited"
                if phs else "&mdash;")
        size, autofit = _title_declaration(lay)
        title = f"{size:g}pt" if size else "<span class='note'>inherits</span>"
        if autofit == "normAutofit":
            title += " <span class='note'>+ shrink</span>"
        rows.append(
            f"<tr><td>{esc(lay['name'])}</td>"
            f"<td><span class='tag' style='margin:0'>{esc(lay['type'] or '—')}</span></td>"
            f"<td>{esc(kinds) or '&mdash;'}</td>"
            f"<td>{title}</td>"
            f"<td>{geom}</td>"
            f"<td>{_layout_bg_cell(lay.get('background'))}</td></tr>")
    return "".join(rows)


def render_style_spec(spec: dict, spec_id: str, can_save: bool = True,
                      message: str = "") -> str:
    meta = spec["meta"]
    theme = spec.get("theme") or {}
    master = spec.get("master") or {}
    size = meta["slide_size_emu"]

    no_slides = meta["slide_count"] == 0
    provenance = (
        "This file carries no slides, which is exactly what a master "
        "submission should look like."
        if no_slides else
        f"This file also carries {meta['slide_count']} slide(s). They were "
        "ignored: only the master, its layouts, and the theme were read.")

    multi_master = ""
    if meta["master_count"] > 1:
        multi_master = _warn(
            f"This file has {meta['master_count']} slide masters. The spec "
            "describes the dominant one; the others were not read.")

    save_form = ""
    if can_save:
        save_form = f"""
<div class="card">
  <div class="tag">Use it</div>
  <h2 style="margin-top:0">Save as a formatting profile</h2>
  <p class="sub">Turns this spec into a profile the audit engine reads, so decks
  can be checked against this master. The spec stays the source of truth; the
  profile is a view of it.</p>
  <form action="/spec/{esc(spec_id)}/profile" method="post" class="actions"
        style="gap:0.6rem;align-items:center">
    <input name="name" required placeholder="Profile name, e.g. Client X master"
      style="border:1px solid var(--line);border-radius:10px;padding:0.42rem 0.75rem;
             background:#fff;color:var(--teal);font-size:0.88rem;min-width:18rem">
    <button class="btn primary" type="submit">Save profile</button>
  </form>
</div>"""

    body = f"""
<h1>{esc(meta.get('source_file') or 'Style Spec')}</h1>
<p class="sub">{provenance} Canvas {_in(size['width'])} &times; {_in(size['height'])},
{len(spec['layouts'])} layouts. <a href="/spec/{esc(spec_id)}.json">Download the
Style Spec JSON</a> &middot; <a href="/master">read another master</a></p>
{_warn(message)}{multi_master}

<div class="card">
  <div class="tag">Theme</div>
  <h2 style="margin-top:0">Colours the theme declares</h2>
  <p class="sub">Read from the theme part, not from what any shape happens to use.
  These are the roles Stage 2 substitutes when it restyles a layout.</p>
  {_swatches(theme.get('colors') or {})}
  <h3 style="margin-bottom:0.3rem">Fonts</h3>
  {_fonts_block(theme.get('fonts') or {})}
  <h3 style="margin-bottom:0.3rem">Type sizes the master declares</h3>
  {_text_styles_block(master.get('text_styles') or {})}
</div>

<div class="card">
  <div class="tag">Grid</div>
  <h2 style="margin-top:0">Margins and columns</h2>
  {_grid_block(spec.get('grid') or {})}
  {_space_agreement_note(spec.get('grid') or {}, spec['layouts'])}
</div>

<div class="card">
  <div class="tag">Master</div>
  <h2 style="margin-top:0">Fixed elements</h2>
  <table class="w3"><tbody>
  <tr><td>Background</td><td>{_background_cell(master.get('background'))}</td></tr>
  {_furniture_row('Footer', master.get('footer'))}
  {_furniture_row('Slide number', master.get('slide_number'))}
  {_furniture_row('Date', master.get('date'))}
  </tbody></table>
  <h3 style="margin-bottom:0.3rem">Logo</h3>
  {_logo_block(spec.get('brand') or {})}
</div>

<div class="card">
  <div class="tag">Layouts</div>
  <h2 style="margin-top:0">{len(spec['layouts'])} layouts</h2>
  <p class="sub">"Explicit" means the layout pins its own placeholder geometry;
  "inherited" means it follows the master, so a master edit moves it. The title
  column is the size each layout DECLARES: that is what a slide's title
  inherits when it lands there.</p>
  {_title_size_note(spec['layouts'])}
  <table class="w3">
    <thead><tr><th>Name</th><th>Archetype</th><th>Placeholders</th>
    <th>Title size</th><th>Geometry</th><th>Background</th></tr></thead>
    <tbody>{_layout_rows(spec['layouts'])}</tbody>
  </table>
</div>
{save_form}"""
    return _shell(f"Style Spec: {meta.get('source_file') or spec_id}", body)
