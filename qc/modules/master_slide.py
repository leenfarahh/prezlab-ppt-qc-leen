"""master_slide module: layout conformance and placeholder geometry audit (U1).

v1 is audit-only: every record is action="flagged". Detection reuses the U1
spike primitives (layout_census, find_geometry_deviations); this module only
maps their output onto FindingRecords and applies the Arabic guard.
"""

from qc.records import FindingRecord, make_record
from spike.u1_master import find_geometry_deviations, layout_census

MODULE = "master_slide"
DEFAULT_TOLERANCE_EMU = 9525


def detect(ctx) -> list[FindingRecord]:
    records: list[FindingRecord] = []
    layout_names, broken = _survey_layouts(ctx.prs)
    records.extend(_no_usable_master_records(broken))
    records.extend(_foreign_master_records(ctx))
    records.extend(_layout_outlier_records(ctx, layout_names))
    records.extend(_geometry_records(ctx))
    return records


def _survey_layouts(prs):
    """Per-slide layout name, plus slides whose layout/master chain is broken.
    Access is wrapped defensively: a slide with a dangling layout relationship
    raises deep inside python-pptx and must not abort the whole audit."""
    layout_names: dict[int, str] = {}
    broken: dict[int, str] = {}
    for s_idx, slide in enumerate(prs.slides):
        try:
            layout = slide.slide_layout
            master = layout.slide_master
            if layout is None or master is None:
                raise ValueError("layout or master missing")
            layout_names[s_idx] = layout.name
        except Exception as exc:
            broken[s_idx] = f"{type(exc).__name__}: {exc}"
    return layout_names, broken


def _foreign_master_records(ctx) -> list[FindingRecord]:
    """Slides living on a non-dominant master (the copy-paste pollution the
    designers unify by hand). Three routes, mirrored in the fixer:

    - clone master, structural twin layout -> deterministic package-level
      repoint (visual no-op, so the Arabic guard is deliberately NOT raised:
      the slide's own XML is untouched);
    - different master, same-named layout -> PowerPoint re-applies the layout
      via COM (medium confidence, designer approval + before/after review;
      Arabic slides keep the guard because placeholders can move);
    - different master, no matching layout -> flagged for manual re-layout.
    """
    from qc.unify import analyze, com_available

    try:
        analysis = analyze(ctx.deck_path.read_bytes())
    except Exception:
        return []  # unreadable package plumbing; other checks still ran
    if not analysis.multiple_masters:
        return []

    records: list[FindingRecord] = []
    n_masters = len(analysis.masters)
    for fs in analysis.foreign:
        arabic = _slide_has_arabic(ctx, fs.slide_index)
        common = dict(slide_index=fs.slide_index, shape_id="-", module=MODULE,
                      issue_type="master_slide.foreign_master",
                      severity="warning", action="flagged",
                      property="slideLayout.master",
                      old_value=fs.layout_name,
                      profile_rule_id="master_slide.single_master")
        if fs.twin_layout_part:
            records.append(make_record(
                **{**common, "severity": "error"},
                confidence="deterministic",
                new_value=fs.twin_layout_name,
                locator=f"dedup:{fs.twin_layout_part}",
                arabic_flag=False,
                message=f"Slide sits on a duplicate copy of the main master "
                        f"(deck has {n_masters}). Re-pointing it to the "
                        f"identical '{fs.twin_layout_name}' layout on the main "
                        "master changes nothing visually; the emptied duplicate "
                        "master is then removed.",
            ))
        elif fs.name_match_layout:
            message = (f"Slide uses layout '{fs.layout_name}' from a different "
                       f"master (deck has {n_masters}). PowerPoint can re-apply "
                       f"the main master's '{fs.name_match_layout}' layout "
                       "(same behavior as doing it in the layout gallery); "
                       "review the before/after preview, placeholders may move.")
            if arabic:
                message += " Arabic content, manual review."
            if not com_available():
                message += " Needs the desktop (PowerPoint) box to apply."
            records.append(make_record(
                **common, confidence="medium",
                new_value=fs.name_match_layout,
                locator=f"com:{fs.name_match_layout}",
                arabic_flag=arabic,
                message=message,
            ))
        else:
            records.append(make_record(
                **common, confidence="high", new_value=None,
                arabic_flag=arabic,
                message=f"Slide uses layout '{fs.layout_name}' from a different "
                        f"master (deck has {n_masters}) and the main master has "
                        "no layout with that name. Re-apply a layout manually "
                        "(one slide at a time), then re-audit.",
            ))
    return records


def _no_usable_master_records(broken: dict[int, str]) -> list[FindingRecord]:
    return [
        make_record(
            slide_index=s_idx, shape_id="-", module=MODULE,
            issue_type="master_slide.no_usable_master",
            severity="error", action="flagged", confidence="high",
            property="slideLayout",
            message=f"Slide layout/master chain is missing or unreadable ({reason}); "
                    "slide excluded from layout and geometry checks.",
        )
        for s_idx, reason in sorted(broken.items())
    ]


def _slide_has_arabic(ctx, slide_index: int) -> bool:
    return any(s_idx == slide_index for s_idx, _ in ctx.arabic_shapes)


def _layout_outlier_records(ctx, layout_names: dict[int, str]) -> list[FindingRecord]:
    records: list[FindingRecord] = []
    allowlist = ctx.profile.get("master_slide.layout_allowlist") or []

    if allowlist:
        for s_idx, name in sorted(layout_names.items()):
            if name in allowlist:
                continue
            records.append(make_record(
                slide_index=s_idx, shape_id="-", module=MODULE,
                issue_type="master_slide.layout_outlier",
                severity="warning", action="flagged", confidence="high",
                property="slideLayout",
                old_value=name, new_value=", ".join(allowlist),
                arabic_flag=_slide_has_arabic(ctx, s_idx),
                profile_rule_id="master_slide.layout_allowlist",
                message=f"Slide layout '{name}' is not in the profile layout allowlist.",
            ))
        return records

    if not layout_names:
        return records
    try:
        census = layout_census(ctx.prs)
    except Exception:
        # Broken layout chains are already reported as no_usable_master.
        return records
    dominant = census["dominant"]
    for s_idx in census["outlier_slide_indices"]:
        records.append(make_record(
            slide_index=s_idx, shape_id="-", module=MODULE,
            issue_type="master_slide.layout_outlier",
            severity="warning", action="flagged", confidence="medium",
            property="slideLayout",
            old_value=layout_names.get(s_idx), new_value=dominant,
            arabic_flag=_slide_has_arabic(ctx, s_idx),
            profile_rule_id="master_slide.layout_allowlist",
            message=f"Slide uses layout '{layout_names.get(s_idx)}' but the deck's "
                    f"dominant layout '{dominant}' was inferred from a layout census "
                    "(profile allowlist is empty).",
        ))
    return records


def _geometry_records(ctx) -> list[FindingRecord]:
    tolerance = ctx.profile.get(
        "master_slide.geometry_tolerance_emu", DEFAULT_TOLERANCE_EMU)
    try:
        deviations = find_geometry_deviations(ctx.prs, tolerance)
    except Exception:
        # Broken layout chains are already reported as no_usable_master.
        return []

    records: list[FindingRecord] = []
    for dev in deviations:
        shape_id = "-"
        for ph in ctx.prs.slides[dev.slide_index].placeholders:
            if ph.placeholder_format.idx == dev.ph_idx:
                shape_id = str(ph.shape_id)
                break
        arabic = ctx.shape_has_arabic(dev.slide_index, shape_id)
        message = (f"Placeholder idx {dev.ph_idx} has an explicit xfrm deviating "
                   f"from its inherited layout/master geometry by more than "
                   f"{tolerance} EMU.")
        if arabic:
            message += (" Contains Arabic text (geometry-only fix, text "
                        "untouched).")
        records.append(make_record(
            slide_index=dev.slide_index, shape_id=shape_id, module=MODULE,
            issue_type="master_slide.placeholder_geometry_off",
            severity="warning", action="flagged", confidence="deterministic",
            property="spPr.xfrm",
            old_value=dev.slide_xfrm, new_value=dev.baseline_xfrm,
            arabic_flag=arabic,
            profile_rule_id="master_slide.geometry_tolerance_emu",
            message=message,
        ))
    return records
