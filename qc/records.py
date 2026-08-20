"""FindingRecord: the system of record (PRD Appendix A.2).

Every detection or change is one record, emitted by the module that found
it. Reports, summaries, and the change manifest are projections of the
record set; nothing is ever reconstructed by diffing files.
"""

import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone

SEVERITIES = ("error", "warning", "info")
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
    created_at: str = field(default="")

    def to_dict(self) -> dict:
        return asdict(self)


def make_record(*, slide_index: int, shape_id, module: str, issue_type: str,
                message: str, severity: str = "warning", action: str = "flagged",
                confidence: str = "high", property: str | None = None,
                old_value=None, new_value=None, shape_path: str | None = None,
                arabic_flag: bool = False, profile_rule_id: str | None = None,
                job_id: str | None = None, locator: str | None = None) -> FindingRecord:
    assert severity in SEVERITIES, severity
    assert action in ACTIONS, action
    assert confidence in CONFIDENCES, confidence
    assert module in RECORD_MODULES, module
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
        created_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
    )
