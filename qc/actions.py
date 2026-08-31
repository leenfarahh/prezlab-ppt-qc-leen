"""What the assistant may be asked to DO, and how a sentence becomes a plan.

The ask box answered questions and pointed at cards (qc.chat). A designer with
forty slides and six hundred findings does not want to be pointed at cards. They
want to say "fix the fonts on slide 7" and have it done, and the tool already
knows how to do it - the knowledge was just locked behind knowing which of five
pages carries the button (design lead, 27/08/2026).

So the assistant acts. Three rules keep that from being a second, unreviewed
write path into a client's deck, and they are the same three the rest of the
tool runs on.

THE MODEL PICKS FROM A CLOSED SET, NAMED IN IDS THAT EXIST. It is handed this
job's own vocabulary - the issue types actually present, the finding ids
actually open, the remedy ids each of those findings actually offers - and it
answers with a name and some ids. It never writes a selector, never names a
shape, never states a coordinate. An id that is not in the vocabulary is
discarded, exactly as qc.layoutpick drops an invented layout name.

CODE RESOLVES THE PLAN, AND THE SUMMARY IS BUILT FROM WHAT IT RESOLVED. The
sentence a designer confirms is not the model's sentence. It is generated here
from the records that were actually selected, so "six fixes on slide 7" is a
count of six real record ids and cannot be a plausible number. A request that
resolves to nothing is refused with what WAS available, rather than performed
as an empty success.

AND NOTHING HERE PERFORMS ANYTHING. This module plans; the route performs, by
calling the very same function the button on the page calls. That is deliberate
and it is the whole safety story: there is no fix in this package that the
assistant can reach and a designer cannot, no fix that skips the re-audit, and
no fix that lands outside the Undo list. Asking is a new door onto the tool, not
a new tool.

The confirmation gate stays. A plan comes back with a summary and a handle, and
nothing happens until the designer presses the button - one press, on the page
they are already on, instead of a page they had to find. A chat that quietly
rewrote a client deck because a sentence was ambiguous is the failure this whole
tool is built to prevent, and being one click faster is not worth it.
"""

from dataclasses import dataclass, field

# How many of each list the model is shown. Enough to name the thing a designer
# just asked about without spending the window on the tail of a 2000-record
# audit; every cap that bites is stated in the vocabulary itself, so the model
# can say "and more beyond these" rather than implying the list is complete.
MAX_ISSUE_TYPES = 24
MAX_OPEN_FINDINGS = 24
MAX_APPLIED = 20
MAX_REMOVALS = 20


class Refused(ValueError):
    """The request cannot be planned against this job.

    A refusal is an ANSWER, not an error: "there is no fixable font finding on
    slide 7, there are two on slide 9" is a useful thing to be told. So the
    message is written for a designer to read and always says what was actually
    there.
    """


@dataclass
class Plan:
    """One resolved request: exactly what would happen, and to what."""

    name: str
    summary: str                # built from the resolved targets, not the model
    changes: bool = True        # whether it writes to the deck
    record_ids: list = field(default_factory=list)
    picks: list = field(default_factory=list)      # [(finding, remedy)]
    finding_ids: list = field(default_factory=list)
    remove_ids: list = field(default_factory=list)
    scope: str = "deck"
    slide: int | None = None    # zero-based, None for the whole deck
    include_holds: bool = False


# --- the vocabulary the model is given ------------------------------------

# Every action, with the sentence that tells a model when it is the right one.
# Kept beside the planners below so a name added to one and not the other is a
# visible mistake rather than a silent no-op.
WHAT_THEY_DO = {
    "fix_findings": ("Apply the audit's own fixes. These are deterministic "
                     "conformance to the profile - a font that is not in the "
                     "allowed set, a shape off the margin - and each one is "
                     "reversible. Name issue types, and a slide when the "
                     "designer named one."),
    "decide": ("Hand the whole thing over: apply every audit fix AND take "
               "every design decision the tool has an answer for, on one slide "
               "or on the whole deck. This is the right action for 'sort this "
               "out', 'do what you can', 'clean it up'."),
    "take_remedy": ("Take one specific way out of one specific design finding. "
                    "Use this when the designer named the finding or the "
                    "remedy, because a design finding has more than one right "
                    "answer and picking is the designer's call."),
    "undo": ("Take back design decisions that were already applied, exactly, "
             "putting the deck back as it was."),
    "recheck": ("Re-read the deck as it now stands, so the counts on the page "
                "match the file. Changes nothing."),
    "remove_pieces": ("Take out pieces the migration proposed removing - "
                      "content the master's own furniture now duplicates. "
                      "Nothing leaves a slide unless this is asked for."),
}


def _fixable(job) -> list:
    from .fixer import is_fixable

    return [r for r in ((job.get("manifest") or {}).get("records") or [])
            if r.get("module") != "preflight" and is_fixable(r)]


def _open_findings(job) -> list:
    answered = {a.finding_id for a in job.get("design_applied") or []}
    return [f for f in job.get("design") or []
            if f.finding_id not in answered]


def _pending_removals(job) -> list:
    done = set(job.get("removed") or [])
    return [c for c in job.get("changes") or []
            if getattr(c, "remove_id", None)
            and getattr(c, "remove_op", None)
            and c.remove_id not in done]


def _slide_count(job) -> int:
    manifest = job.get("manifest") or {}
    if manifest.get("slides"):
        return int(manifest["slides"])
    return len(job.get("plans") or [])


def vocabulary(job) -> dict:
    """The ids this job can actually be asked about, and nothing else.

    Assembled per job rather than declared once, because that is what makes a
    hallucinated id impossible to act on: the model is never shown a finding
    that is not open or an issue type this deck does not have, so the common
    failure mode is the model asking for something reasonable and being refused
    with what WAS there, not the deck being changed in a way nobody meant.
    """
    fixable = _fixable(job)
    by_type: dict[str, int] = {}
    for record in fixable:
        by_type[record["issue_type"]] = by_type.get(record["issue_type"], 0) + 1
    ordered = sorted(by_type.items(), key=lambda kv: -kv[1])

    findings = _open_findings(job)
    applied = list(job.get("design_applied") or [])
    removals = _pending_removals(job)

    out = {
        "slides": _slide_count(job),
        "can_change_the_deck": job.get("deck") is not None,
        # Whether this job was audited at all. A format-only run has plans and
        # a review but no records, so nothing that reads the manifest can be
        # asked for on it - and offering it anyway would produce a refusal for
        # every reasonable request a designer makes on that page.
        "has_audit": bool(job.get("manifest")),
        "issue_types_you_may_name": [
            {"issue_type": name, "fixable_now": n}
            for name, n in ordered[:MAX_ISSUE_TYPES]],
        "issue_types_omitted": max(0, len(ordered) - MAX_ISSUE_TYPES),
        "open_design_findings": [
            {"finding": f.finding_id, "headline": f.headline,
             "severity": f.severity, "kind": f.kind,
             "slides": [i + 1 for i in f.slides[:8]],
             "remedies": [{"remedy": o.remedy_id, "label": o.label}
                          for o in f.options]}
            for f in findings[:MAX_OPEN_FINDINGS]],
        "design_findings_omitted": max(0, len(findings) - MAX_OPEN_FINDINGS),
        "decisions_already_applied": [
            {"finding": a.finding_id,
             "headline": getattr(a, "headline", "") or ""}
            for a in applied[-MAX_APPLIED:] if getattr(a, "undo", None)],
        "pending_removals": [
            {"removal": c.remove_id, "slide": c.slide_index + 1,
             "what": c.detail}
            for c in removals[:MAX_REMOVALS]],
        "removals_omitted": max(0, len(removals) - MAX_REMOVALS),
    }
    out["actions"] = {name: why for name, why in WHAT_THEY_DO.items()
                      if _offered(name, out)}
    return out


def _offered(name: str, vocab: dict) -> bool:
    """Whether an action is worth showing at all for this job.

    An action with nothing to act on is not a capability, it is a trap: the
    model picks it because it reads as the right verb, and the designer gets a
    refusal instead of an answer."""
    if name == "recheck":
        return vocab["has_audit"] and vocab["can_change_the_deck"]
    if not vocab["can_change_the_deck"]:
        return False
    if name == "fix_findings":
        return bool(vocab["issue_types_you_may_name"])
    if name in ("decide", "take_remedy"):
        return bool(vocab["open_design_findings"]
                    or vocab["issue_types_you_may_name"])
    if name == "undo":
        return bool(vocab["decisions_already_applied"])
    if name == "remove_pieces":
        return bool(vocab["pending_removals"])
    return True


# The action the model may return beside its answer. `null` is the normal case:
# most questions are questions.
ACTION_SCHEMA = {
    "type": "object", "additionalProperties": False,
    "required": ["name"],
    "properties": {
        "name": {"type": "string"},
        "issue_types": {"type": "array", "items": {"type": "string"}},
        "slide": {"type": "integer"},          # 1-based, as a designer says it
        "finding": {"type": "string"},
        "remedy": {"type": "string"},
        "findings": {"type": "array", "items": {"type": "string"}},
        "removals": {"type": "array", "items": {"type": "string"}},
        "include_holds": {"type": "boolean"},
    },
}


# --- turning a request into a plan ----------------------------------------


def _slide_arg(request: dict, job) -> int | None:
    """The slide a request names, zero-based, or None for the whole deck.

    A slide outside the deck is refused rather than clamped. Clamping would
    turn "fix slide 40" on a 12-slide deck into a silent change to slide 12,
    which is the exact class of surprise this tool exists to prevent."""
    raw = request.get("slide")
    if raw is None:
        return None
    total = _slide_count(job)
    if not isinstance(raw, int) or not (1 <= raw <= total):
        raise Refused(f"There is no slide {raw} in this deck; it has {total}.")
    return raw - 1


def _plural(n: int, one: str, many: str | None = None) -> str:
    return f"{n} {one if n == 1 else (many or one + 's')}"


def _where(slide: int | None) -> str:
    return f"slide {slide + 1}" if slide is not None else "the whole deck"


def _plan_fix(job, request: dict) -> Plan:
    from .fixer import needs_explicit_tick

    slide = _slide_arg(request, job)
    wanted = {str(t) for t in (request.get("issue_types") or [])}
    fixable = _fixable(job)
    if not fixable:
        raise Refused("Nothing in this audit is one the fixer can perform, so "
                      "there is nothing to apply.")

    here = [r for r in fixable
            if slide is None or r["slide_index"] == slide]
    if not here:
        others = sorted({r["slide_index"] + 1 for r in fixable})[:8]
        raise Refused(
            f"There is no fixable finding on {_where(slide)}. The fixable ones "
            f"are on slide{'s' if len(others) != 1 else ''} "
            + ", ".join(str(n) for n in others) + ".")

    chosen = [r for r in here if not wanted or r["issue_type"] in wanted]
    if not chosen:
        have = sorted({r["issue_type"] for r in here})[:8]
        raise Refused(
            f"None of those are on {_where(slide)}. What is fixable there: "
            + ", ".join(have) + ".")

    holds = [r for r in chosen if needs_explicit_tick(r)]
    include_holds = bool(request.get("include_holds"))
    if not include_holds:
        chosen = [r for r in chosen if not needs_explicit_tick(r)]
    if not chosen:
        raise Refused(
            f"All {_plural(len(holds), 'fix', 'fixes')} on {_where(slide)} ask "
            f"for your explicit approval - Arabic font substitutions and whole "
            f"slide body moves are never applied on a sentence. Tick them on "
            f"the slide, or say to include the ones held for approval.")

    counts: dict[str, int] = {}
    for record in chosen:
        counts[record["issue_type"]] = counts.get(record["issue_type"], 0) + 1
    detail = ", ".join(f"{n} {name}" for name, n in
                       sorted(counts.items(), key=lambda kv: -kv[1])[:5])
    summary = (f"Apply {_plural(len(chosen), 'audit fix', 'audit fixes')} on "
               f"{_where(slide)}: {detail}. The deck is re-audited afterwards, "
               f"and every one of them is reversible.")
    if holds and not include_holds:
        summary += (f" {_plural(len(holds), 'fix', 'fixes')} held for your "
                    f"explicit approval {'are' if len(holds) != 1 else 'is'} "
                    f"not included.")
    return Plan(name="fix_findings", summary=summary,
                record_ids=[r["record_id"] for r in chosen],
                slide=slide, scope="slide" if slide is not None else "deck")


def _plan_decide(job, request: dict) -> Plan:
    from .design import auto_choice
    from .fixer import needs_explicit_tick

    slide = _slide_arg(request, job)
    include_holds = bool(request.get("include_holds"))
    records = [r for r in _fixable(job)
               if slide is None or r["slide_index"] == slide]
    if not include_holds:
        records = [r for r in records if not needs_explicit_tick(r)]
    findings = [f for f in _open_findings(job)
                if slide is None or f.slides == [slide]]
    decides = [f for f in findings if auto_choice(f) is not None]
    left = [f for f in findings if auto_choice(f) is None]

    if not records and not decides:
        raise Refused(
            f"Nothing on {_where(slide)} is the tool's to decide. "
            + (f"{_plural(len(left), 'design decision')} there "
               f"{'need' if len(left) != 1 else 'needs'} a designer's eye, and "
               f"the cards say why." if left
               else "There is nothing open there at all."))

    bits = []
    if records:
        bits.append(_plural(len(records), "audit fix", "audit fixes"))
    if decides:
        bits.append(_plural(len(decides), "design decision"))
    summary = (f"Decide {_where(slide)}: apply " + " and ".join(bits)
               + ". The audit fixes land first and the deck is re-audited "
                 "between the two passes, so the design judgments are made "
                 "about the deck as the fixes leave it.")
    if left:
        summary += (f" {_plural(len(left), 'finding')} would be left for you, "
                    f"because more than one answer is right and only you are "
                    f"looking at the slide.")
    return Plan(name="decide", summary=summary,
                scope="slide" if slide is not None else "deck",
                slide=slide, include_holds=include_holds,
                record_ids=[r["record_id"] for r in records],
                picks=[(f, auto_choice(f)) for f in decides])


def _plan_remedy(job, request: dict) -> Plan:
    findings = {f.finding_id: f for f in _open_findings(job)}
    finding = findings.get(str(request.get("finding") or ""))
    if finding is None:
        raise Refused("That design finding is not open on this deck. It may "
                      "have been answered already, or the deck may have "
                      "changed since it was found.")
    wanted = str(request.get("remedy") or "")
    remedy = next((o for o in finding.options if o.remedy_id == wanted), None)
    if remedy is None:
        offers = ", ".join(f"{o.label}" for o in finding.options[:4])
        raise Refused(f"That is not one of the ways out of this finding. It "
                      f"offers: {offers}.")
    where = ", ".join(str(i + 1) for i in finding.slides[:6])
    summary = (f"{remedy.label}, for \"{finding.headline}\" on "
               f"slide{'s' if len(finding.slides) != 1 else ''} {where}. "
               f"{remedy.note} It lands in the same list as a hand-picked "
               f"decision, with the same Undo.")
    return Plan(name="take_remedy", summary=summary,
                picks=[(finding, remedy)],
                slide=finding.slides[0] if len(finding.slides) == 1 else None)


def _plan_undo(job, request: dict) -> Plan:
    applied = [a for a in (job.get("design_applied") or [])
               if getattr(a, "undo", None)]
    if not applied:
        raise Refused("No decision on this deck has been applied yet, so there "
                      "is nothing to take back.")
    wanted = [str(f) for f in (request.get("findings") or [])]
    known = {a.finding_id for a in applied}
    chosen = [f for f in wanted if f in known]
    if wanted and not chosen:
        raise Refused("Those decisions are not in the applied list for this "
                      "deck, so there is nothing there to reverse.")
    if not chosen:
        # No id named: the most recent one, because "undo that" means the last
        # thing that happened and guessing wider would reverse work nobody
        # asked about.
        chosen = [applied[-1].finding_id]
    summary = (f"Take back {_plural(len(chosen), 'decision')}, putting the "
               f"deck back exactly as it was. Anything that touched the same "
               f"shape comes back with it, and the page says what did.")
    return Plan(name="undo", summary=summary, finding_ids=chosen)


def _plan_recheck(job, request: dict) -> Plan:
    if not job.get("manifest"):
        raise Refused("This run applied the master but did not audit the "
                      "result, so there are no counts to refresh. The Prepare "
                      "a deck page does both in one pass.")
    return Plan(name="recheck", changes=False,
                summary="Read the deck again as it now stands, so the counts "
                        "on this page come from the current file. Nothing is "
                        "changed.")


def _plan_remove(job, request: dict) -> Plan:
    pending = {c.remove_id: c for c in _pending_removals(job)}
    if not pending:
        raise Refused("Nothing is waiting to be removed from this deck. The "
                      "migration only proposes a removal when the master's own "
                      "furniture duplicates something already on the slide.")
    slide = _slide_arg(request, job)
    wanted = [str(r) for r in (request.get("removals") or [])]
    chosen = [pending[r] for r in wanted if r in pending]
    if not chosen and slide is not None:
        chosen = [c for c in pending.values() if c.slide_index == slide]
    if not chosen and not wanted and slide is None:
        chosen = list(pending.values())
    if not chosen:
        raise Refused(f"There is nothing proposed for removal on "
                      f"{_where(slide)}.")
    where = sorted({c.slide_index + 1 for c in chosen})[:8]
    summary = (f"Remove {_plural(len(chosen), 'piece')} the migration flagged "
               f"as duplicated by the master, on slide"
               f"{'s' if len(where) != 1 else ''} "
               + ", ".join(str(n) for n in where)
               + ". Each one comes back as an ordinary change carrying its own "
                 "Undo, so a removal you regret is reversed like any other.")
    return Plan(name="remove_pieces", summary=summary,
                remove_ids=[c.remove_id for c in chosen])


_PLANNERS = {
    "fix_findings": _plan_fix,
    "decide": _plan_decide,
    "take_remedy": _plan_remedy,
    "undo": _plan_undo,
    "recheck": _plan_recheck,
    "remove_pieces": _plan_remove,
}


def plan(job, request: dict) -> Plan:
    """Resolve one requested action against this job. Raises Refused with a
    sentence a designer can act on."""
    name = str((request or {}).get("name") or "")
    planner = _PLANNERS.get(name)
    if planner is None:
        raise Refused("That is not something this assistant can do. It can "
                      "apply audit fixes, take a design remedy, decide a slide "
                      "or the whole deck, undo a decision, remove a flagged "
                      "piece, or re-read the deck.")
    if name != "recheck" and job.get("deck") is None:
        raise Refused("The deck is no longer held in memory, so nothing can be "
                      "changed. Run it again and the same request will work.")
    return planner(job, request or {})
