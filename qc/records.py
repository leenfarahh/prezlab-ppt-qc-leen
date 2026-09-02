"""FindingRecord: the system of record (PRD Appendix A.2).

Every detection or change is one record, emitted by the module that found
it. Reports, summaries, and the change manifest are projections of the
record set; nothing is ever reconstructed by diffing files.
"""

import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone

SEVERITIES = ("error", "warning", "info")
# Where a finding came from: arithmetic over the file, or a judgment about a
# picture of it. See FindingRecord.source.
SOURCES = ("measured", "vision")
ACTIONS = ("flagged", "changed", "skipped")
CONFIDENCES = ("deterministic", "high", "medium", "low")
MODULES = ("master_slide", "font", "margin_alignment", "color_palette",
           "shape_size", "header_footer", "typography")
# "preflight" is an engine-level pseudo-module for cross-cutting records
# (issue_type preflight.*); it is not a selectable audit module.
RECORD_MODULES = MODULES + ("preflight",)


@dataclass
class FindingRecord:
    record_id: str
    job_id: str | None
    slide_index: int  # zero-based
    shape_id: str
    shape_path: str | None
    module: str
    issue_type: str
    property: str | None
    old_value: str | None
    new_value: str | None
    severity: str
    action: str
    confidence: str
    arabic_flag: bool
    profile_rule_id: str | None
    message: str
    # Internal addressing beyond Appendix A.2 (superset is allowed): locates
    # the exact paragraph/run inside the shape for fix application, e.g.
    # "p2/r1". None for shape- or slide-level records.
    locator: str | None = None
    # WHO NOTICED. "measured" is a rule in qc/modules or qc/design.py that
    # compared two numbers; "vision" is a judgment the model made about a
    # picture of the slide, which code then re-measured before it became this
    # record (qc.copilot, qc.components).
    #
    # Recorded because the two are different KINDS of claim and a report that
    # interleaves them reads as one long list of equal-weight complaints. A
    # measured finding is a fact - this run is Arial and the profile says
    # Georgia - and it is usually also boring. A vision finding is the answer to
    # "what would a designer change about this slide", which is the question
    # someone actually opened the tool to ask, and it belongs at the top
    # (design lead, 31/08/2026).
    #
    # It is NOT a confidence and must not be read as one: `confidence` already
    # says how sure the tool is, and a vision finding re-measured in code can be
    # more certain than a geometric one inferred from proximity.
    source: str = "measured"
    created_at: str = field(default="")

    def to_dict(self) -> dict:
        return asdict(self)


def make_record(*, slide_index: int, shape_id, module: str, issue_type: str,
                message: str, severity: str = "warning", action: str = "flagged",
                confidence: str = "high", property: str | None = None,
                old_value=None, new_value=None, shape_path: str | None = None,
                arabic_flag: bool = False, profile_rule_id: str | None = None,
                job_id: str | None = None, locator: str | None = None,
                source: str = "measured") -> FindingRecord:
    assert severity in SEVERITIES, severity
    assert action in ACTIONS, action
    assert confidence in CONFIDENCES, confidence
    assert module in RECORD_MODULES, module
    assert source in SOURCES, source
    return FindingRecord(
        record_id=uuid.uuid4().hex,
        job_id=job_id,
        slide_index=slide_index,
        shape_id=str(shape_id),
        shape_path=shape_path,
        module=module,
        issue_type=issue_type,
        property=property,
        old_value=None if old_value is None else str(old_value),
        new_value=None if new_value is None else str(new_value),
        severity=severity,
        action=action,
        confidence=confidence,
        arabic_flag=arabic_flag,
        profile_rule_id=profile_rule_id,
        message=message,
        locator=locator,
        source=source,
        created_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
    )


# --------------------------------------------------- who wins when two agree


# Issue types that answer the SAME QUESTION about a shape, grouped by the
# question. Only records inside one family can supersede each other, and the
# family is named explicitly rather than inferred from `property`, because the
# property namespace is shared by findings that are not about the same thing at
# all: header_footer.footer_off_canvas ("the footer is off the page") and
# margin_alignment.edge_misaligned ("this sits 3mm off its line") both write
# spPr.xfrm.off.y, and letting one stand in for the other would delete a real
# defect to make room for a suggestion about a different one.
#
# Anything not listed has no family and can never be superseded. That is the
# safe default and it covers most of an audit: fonts, colours, margins,
# overflow, overlap - the findings nobody asked a model about.
_CLAIM_FAMILIES = {
    # WHICH LINE DOES THIS SHAPE SIT ON. Six names for it across the measured
    # modules and the two vision passes; one question.
    "margin_alignment.edge_misaligned": "edge",
    "margin_alignment.component_edge_misaligned": "edge",
    "margin_alignment.panel_row_misaligned": "edge",
    "margin_alignment.space_edge_misaligned": "edge",
    "margin_alignment.cluster_rhythm": "edge",
    "margin_alignment.recurring_off_position": "edge",
    "shape_size.off_grid": "edge",
    # HOW ARE THESE SPACED.
    "margin_alignment.uneven_spacing": "spacing",
    # HOW BIG SHOULD THIS BE.
    "shape_size.size_mismatch": "size",
}

# The axis a geometry write moves the shape on. "spPr.xfrm.off" carries no
# axis, so it claims both: a record that repositions a shape outright collides
# with one that moves it on either axis alone.
_AXES = {
    "spPr.xfrm.off.x": ("x",),
    "spPr.xfrm.off.y": ("y",),
    "spPr.xfrm.ext": ("size",),
    "spPr.xfrm.off": ("x", "y"),
    "spPr.xfrm": ("x", "y", "size"),
}


def claim_keys(record: dict) -> set:
    """What this record CLAIMS: one key per (family, shape, axis) it speaks for.

    A set rather than a single key because a record whose property is
    "spPr.xfrm.off" moves the shape on both axes and therefore claims both.

    Empty for a record with no family, which is how "this can never be
    superseded" is expressed - there is nothing for another record to match.
    """
    family = _CLAIM_FAMILIES.get(record.get("issue_type"))
    if family is None:
        return set()
    axes = _AXES.get(record.get("property"))
    if not axes:
        return set()
    return {(record["slide_index"], str(record["shape_id"]), family, axis)
            for axis in axes}


def vision_wins(records: list[dict]) -> list[dict]:
    """`records` with every measured record a vision record already claims
    removed.

    THE MODEL'S ANSWER OUTRANKS THE RULE'S (design lead, 02/09/2026). Both
    halves were already being kept and the duplicate dropped; which one got
    dropped was decided by arrival order, and the measured half always arrives
    first - it comes out of the audit, and a vision pass is a button pressed
    afterwards - so the model's answer was the one that disappeared, every
    time, on every path.

    That is backwards for what the two claims ARE. A measured record says a
    number is off a threshold. A vision record says what a designer would
    change about the slide and why, re-verified against the same geometry the
    measured one read - so it carries the rule's precision AND a reason a
    person can act on. When they name the same shape and the same move, the one
    with the reason on it is the one to show.

    Narrow by construction (see _CLAIM_FAMILIES): only findings answering the
    same question about the same shape on the same axis can displace each
    other. A measured finding the model said nothing about is untouched, which
    is most of them.
    """
    claimed = set()
    for rec in records:
        if rec.get("source") == "vision":
            claimed |= claim_keys(rec)
    if not claimed:
        return list(records)
    return [rec for rec in records
            if rec.get("source") == "vision"
            or not (claim_keys(rec) & claimed)]
