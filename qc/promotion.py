"""Designer-validated fix promotion.

The rule the whole fix pipeline runs on: a fix is PRE-SELECTED only when it
has earned it. Two ways to earn it:

1. Deterministic by construction: the target state is explicit and
   unambiguous (confidence "deterministic"). Pre-selected from day one.
2. Validated by designers: the check's detections have been triaged at
   least MIN_REVIEWED times across decks with a false-alarm rate at or
   below MAX_FP_RATE. Computed live from the triage log, so promotion (and
   demotion) follows the evidence automatically.

Everything else stays a suggestion: tickable, because a designer ticking a
box IS per-change validation, but never ticked for them. Arabic-flagged
findings are never fixable regardless of promotion (hard guard in fixer).
"""

MIN_REVIEWED = 20
MAX_FP_RATE = 0.05


def promoted_issue_types() -> set[str]:
    from .triage import stats

    return {row["issue_type"] for row in stats()
            if row["reviewed"] >= MIN_REVIEWED and row["fp_rate"] <= MAX_FP_RATE}


def promotion_status(issue_type: str) -> dict:
    """For UI tooltips: where this check stands on the road to pre-selection."""
    from .triage import stats

    for row in stats():
        if row["issue_type"] == issue_type:
            promoted = (row["reviewed"] >= MIN_REVIEWED
                        and row["fp_rate"] <= MAX_FP_RATE)
            return {"reviewed": row["reviewed"], "fp_rate": row["fp_rate"],
                    "promoted": promoted,
                    "needed": max(0, MIN_REVIEWED - row["reviewed"])}
    return {"reviewed": 0, "fp_rate": 0.0, "promoted": False,
            "needed": MIN_REVIEWED}
