"""color_palette detection module (audit-only, v1).

Checks two color surfaces per shape against the profile palette:
the shape solid fill (p:spPr/a:solidFill) and each run's explicit text
color (a:rPr/a:solidFill). Theme-colored surfaces are judged by slot
name (on_palette_mode "by_name"); literal srgbClr surfaces are resolved
through tint/shade/lum transforms and matched perceptually (CIEDE2000)
against the named palette.
"""

from qc.records import make_record
from qc.util import iter_shapes_deep
from spike.color_resolver import clr_map, nearest_palette_match, resolve_solid_fill
from spike.ns import find

MODULE = "color_palette"


def _hex(rgb: tuple[int, int, int]) -> str:
    return "{:02X}{:02X}{:02X}".format(*rgb)


def _build_palette(cfg: dict) -> dict[str, tuple[int, int, int]]:
    palette = {}
    for entry in cfg.get("named_colors", []):
        h = entry["hex"].lstrip("#")
        palette[entry["name"]] = tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    return palette


def _check_surface(*, parent_el, surface: str, master, cfg, palette,
                   slide_index: int, shape, shape_path, arabic: bool,
                   locator: str | None = None) -> list:
    """Audit one fill parent element (spPr or rPr). Returns 0 or 1 records.

    `locator` addresses the run for text colors ("p0/r2"), and is what makes
    them fixable: without it a record says "this shape's text is off-palette"
    and the fixer has no way to find WHICH run to repaint."""
    solid = find(parent_el, "a:solidFill")
    if solid is None:
        # Gradient/pattern/picture/none: no single color exists, out of scope.
        return []

    common = dict(slide_index=slide_index, shape_id=shape.shape_id,
                  shape_path=shape_path, module=MODULE, locator=locator,
                  property=f"{surface}.solidFill", action="flagged",
                  arabic_flag=arabic)

    scheme_el = find(solid, "a:schemeClr")
    if scheme_el is not None:
        # by_name mode: a theme-colored surface is on-palette when its
        # clrMap-resolved slot is allowed, regardless of tint/lumMod.
        name = scheme_el.get("val")
        slot = clr_map(master).get(name, name)
        if slot in cfg.get("theme_color_slots", []):
            return []
        return [make_record(
            issue_type="color_palette.disallowed_theme_slot",
            severity="warning", confidence="high",
            old_value=slot,
            profile_rule_id="color_palette.theme_color_slots",
            message=f"Theme color slot '{slot}' (schemeClr '{name}') is not an "
                    f"allowed palette slot.",
            **common)]

    srgb = find(solid, "a:srgbClr")
    if srgb is None:
        return []

    resolved = resolve_solid_fill(parent_el, master)
    if resolved is None:
        return []
    resolved_hex = _hex(resolved)
    nearest_name, delta_e = nearest_palette_match(resolved, palette)
    nearest_hex = _hex(palette[nearest_name])

    match_tol = cfg.get("match_tolerance_deltaE", 2.0)
    replace_max = cfg.get("auto_replace_max_deltaE", 5.0)
    ambiguity = cfg.get("ambiguity_band_deltaE", 10.0)

    if delta_e <= match_tol:
        return []
    if delta_e <= replace_max:
        return [make_record(
            issue_type="color_palette.off_palette_rgb",
            severity="warning", confidence="high",
            old_value=resolved_hex, new_value=nearest_hex,
            profile_rule_id="color_palette.named_colors",
            message=f"Color #{resolved_hex} is off-palette; replace with "
                    f"'{nearest_name}' (#{nearest_hex}, deltaE {delta_e:.1f}).",
            **common)]
    # Past the auto-replace band the nearest palette color is still NAMED as the
    # target, where before these two tiers carried none and so could never be
    # fixed - the row read "no automatic fix" for a colour the tool could name
    # and reach (design lead, 24/08/2026). What changes past this line is not
    # whether a fix exists but who decides: the swap is visible, so it is never
    # pre-ticked and the tick is the designer's approval (qc.fixer
    # _COLOR_TICK_IS_APPROVAL). Deciding between two candidates, rather than
    # accepting one, is what the design page's palette card is for.
    if delta_e <= ambiguity:
        return [make_record(
            issue_type="color_palette.off_palette_rgb",
            severity="warning", confidence="medium",
            old_value=resolved_hex, new_value=nearest_hex,
            profile_rule_id="color_palette.named_colors",
            message=f"Color #{resolved_hex} is off-palette; closest named color "
                    f"'{nearest_name}' (#{nearest_hex}) is ambiguous at deltaE "
                    f"{delta_e:.1f}. The fix uses it; ticking it is your "
                    f"approval of a visible change.",
            **common)]
    # Severity error, confidence medium, and the two are answering different
    # questions: the tool is CERTAIN this colour is not in the palette and only
    # guessing that the nearest one is what was meant. Confidence is about the
    # fix, so it drops here even though the finding is the strongest of the
    # three - which is also what keeps this tier off the pre-ticked list.
    return [make_record(
        issue_type="color_palette.off_palette_rgb",
        severity="error", confidence="medium",
        old_value=resolved_hex, new_value=nearest_hex,
        profile_rule_id="color_palette.named_colors",
        message=f"Color #{resolved_hex} has no near palette match (nearest "
                f"'{nearest_name}' #{nearest_hex}, deltaE {delta_e:.1f}) - far "
                f"enough to be a different colour, not a variant of one. The "
                f"fix replaces it; ticking it is your approval, and it may be "
                f"deliberate.",
        **common)]


def detect(ctx) -> list:
    cfg = ctx.profile.module_config("color_palette")
    palette = _build_palette(cfg)
    if not palette:
        return []

    records = []
    for s_idx, slide in enumerate(ctx.prs.slides):
        master = slide.slide_layout.slide_master
        for shape, shape_path in iter_shapes_deep(slide.shapes):
            arabic = ctx.shape_has_arabic(s_idx, shape.shape_id)
            spPr = find(shape._element, "p:spPr")
            if spPr is not None:
                records.extend(_check_surface(
                    parent_el=spPr, surface="spPr", master=master, cfg=cfg,
                    palette=palette, slide_index=s_idx, shape=shape,
                    shape_path=shape_path, arabic=arabic))
            if getattr(shape, "has_text_frame", False):
                for p_idx, para in enumerate(shape.text_frame.paragraphs):
                    for r_idx, run in enumerate(para.runs):
                        rPr = find(run._r, "a:rPr")
                        if rPr is None:
                            continue
                        records.extend(_check_surface(
                            parent_el=rPr, surface="rPr", master=master,
                            cfg=cfg, palette=palette, slide_index=s_idx,
                            shape=shape, shape_path=shape_path, arabic=arabic,
                            locator=f"p{p_idx}/r{r_idx}"))
    return records
