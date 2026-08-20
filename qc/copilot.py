"""Design copilot: Claude looks at rendered slides and proposes layout
actions a designer would make; code verifies and executes them.

Strict division of labor (the lesson of the ground-truth calibration):
- Claude (vision) supplies JUDGMENT: which shapes should be distributed,
  aligned, or size-matched. It chooses from a CLOSED action vocabulary and
  may only reference shape ids from the inventory we hand it.
- Code supplies PRECISION: every observation is re-verified against the
  actual geometry and materialized as an ordinary FindingRecord with a
  computed target, so it flows through the existing pipeline - tickable
  (never pre-selected), collision-guarded, before/after rendered.
- The designer supplies APPROVAL, as with every other fix.

Confidentiality: unlike the metadata-only assistant, this sends SLIDE
IMAGES to the Anthropic API. The UI says so explicitly; use it only on
decks approved for cloud processing.
"""

import base64
import io
import json

from pptx import Presentation

from .config import COPILOT_MODEL
from .records import make_record

MAX_SLIDES = 20
MAX_OBS_PER_SLIDE = 3
TOL_EMU = 28575           # perceptual floor, same as calibrated profiles
WINDOW_EMU = 137160       # sanity ceiling for "meant to align"

ACTIONS = ("distribute_row", "distribute_col", "align_left", "align_top",
           "match_widths", "match_heights")

OBSERVATIONS_SCHEMA = {
    "type": "object", "additionalProperties": False,
    "required": ["observations"],
    "properties": {
        "observations": {
            "type": "array",
            "items": {
                "type": "object", "additionalProperties": False,
                "required": ["action", "shape_ids", "rationale"],
                "properties": {
                    "action": {"enum": list(ACTIONS)},
                    "shape_ids": {"type": "array",
                                  "items": {"type": "string"}},
                    "rationale": {"type": "string"},
                },
            },
        },
    },
}

_SYSTEM = """You are a senior presentation designer at Prezlab reviewing a
slide for layout consistency. You see the rendered slide image and an
inventory of its shapes (id, kind, position and size as fractions of the
slide, whether it holds text).

Propose at most {max_obs} layout actions a designer would make, chosen ONLY
from: distribute_row / distribute_col (spread a line of shapes to equal
gaps), align_left / align_top (line up edges that should match),
match_widths / match_heights (make sibling cards the same size). Reference
shapes ONLY by ids from the inventory, at least three per action.

Rules: propose an action only when the improvement would be clearly
visible and safe; analogous elements (a row of cards, a set of columns)
are good targets, decorative compositions are not. If the slide already
looks professionally composed, return an empty list. Write rationales in
clear US English without em dashes."""


def inventory(slide, slide_w: int, slide_h: int) -> list[dict]:
    """Shape inventory Claude reasons over: ids, kind, normalized geometry.
    No text content - the image already shows it; the inventory stays
    minimal."""
    out = []
    for shape in slide.shapes:
        l, t, w, h = shape.left, shape.top, shape.width, shape.height
        if None in (l, t, w, h):
            continue
        has_text = bool(getattr(shape, "has_text_frame", False)
                        and shape.text_frame.text.strip())
        out.append({
            "id": str(shape.shape_id),
            "kind": str(shape.shape_type).split(" ")[0].lower(),
            "x": round(l / slide_w, 3), "y": round(t / slide_h, 3),
            "w": round(w / slide_w, 3), "h": round(h / slide_h, 3),
            "text": has_text,
        })
    return out


def _ask_vision(png: bytes, inv: list[dict]) -> list[dict]:
    import anthropic

    client = anthropic.Anthropic()
    response = client.messages.create(
        model=COPILOT_MODEL,
        max_tokens=8192,
        thinking={"type": "adaptive"},
        output_config={"format": {"type": "json_schema",
                                  "schema": OBSERVATIONS_SCHEMA}},
        system=_SYSTEM.format(max_obs=MAX_OBS_PER_SLIDE),
        messages=[{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": "image/png",
                            "data": base64.b64encode(png).decode()}},
                {"type": "text",
                 "text": "Shape inventory:\n"
                         + json.dumps(inv, sort_keys=True)},
            ],
        }],
    )
    if response.stop_reason == "refusal":
        return []
    text = "".join(b.text for b in response.content if b.type == "text")
    return json.loads(text)["observations"][:MAX_OBS_PER_SLIDE]


def synthesize(slide, s_idx: int, observations: list[dict],
               existing: list[dict]) -> list[dict]:
    """Code is the precision gate: re-verify each observation against the
    real geometry and emit ordinary FindingRecords with computed targets.
    Anything that does not check out is dropped silently."""
    by_id = {str(sh.shape_id): sh for sh in slide.shapes}
    seen = {(r["issue_type"], str(r["shape_id"]), r.get("locator"))
            for r in existing if r["slide_index"] == s_idx}
    out: list[dict] = []

    def emit(**kw):
        rec = make_record(slide_index=s_idx, action="flagged",
                          severity="warning", confidence="medium", **kw)
        key = (rec.issue_type, rec.shape_id, rec.locator)
        if key not in seen:
            seen.add(key)
            out.append(rec.to_dict())

    for obs in observations:
        shapes = [by_id.get(str(i)) for i in obs.get("shape_ids", [])]
        shapes = [s for s in shapes if s is not None
                  and None not in (s.left, s.top, s.width, s.height)
                  and not getattr(s, "rotation", 0)]
        if len(shapes) < 3:
            continue
        note = f"Design copilot: {obs.get('rationale', '').strip()[:160]}"
        action = obs.get("action")

        if action in ("distribute_row", "distribute_col"):
            row = action == "distribute_row"
            run = (lambda s: s.left) if row else (lambda s: s.top)
            size = (lambda s: s.width) if row else (lambda s: s.height)
            shapes = sorted(shapes, key=run)
            gaps = [run(shapes[i + 1]) - (run(shapes[i]) + size(shapes[i]))
                    for i in range(len(shapes) - 1)]
            if any(g <= 0 for g in gaps):
                continue  # overlapping: not a distribution case
            if max(gaps) - min(gaps) <= TOL_EMU:
                continue  # already even to the eye
            ids = ",".join(str(s.shape_id) for s in shapes)
            emit(shape_id=shapes[0].shape_id, module="margin_alignment",
                 issue_type="margin_alignment.uneven_spacing",
                 locator=f"dist-{'row' if row else 'col'}:{ids}",
                 property="spPr.xfrm.off",
                 old_value=", ".join(str(g) for g in gaps),
                 new_value=sum(gaps) // len(gaps),
                 profile_rule_id="geometry.alignment.spacing_tolerance_emu",
                 message=f"{note} Distribute evenly; first and last stay put.")

        elif action in ("align_left", "align_top"):
            left = action == "align_left"
            edge = (lambda s: s.left) if left else (lambda s: s.top)
            vals = sorted(edge(s) for s in shapes)
            median = vals[len(vals) // 2]
            for s in shapes:
                off = abs(edge(s) - median)
                if TOL_EMU < off <= WINDOW_EMU:
                    emit(shape_id=s.shape_id, module="margin_alignment",
                         issue_type="margin_alignment.edge_misaligned",
                         property="spPr.xfrm.off.x" if left
                         else "spPr.xfrm.off.y",
                         old_value=edge(s), new_value=int(median),
                         profile_rule_id="geometry.alignment.edge_tolerance_emu",
                         message=f"{note} Snap to the shared "
                                 f"{'left' if left else 'top'} edge.")

        elif action in ("match_widths", "match_heights"):
            widths = action == "match_widths"
            dim = (lambda s: s.width) if widths else (lambda s: s.height)
            vals = sorted(dim(s) for s in shapes)
            median = vals[len(vals) // 2]
            for s in shapes:
                if abs(dim(s) - median) <= TOL_EMU:
                    continue
                target_w = int(median) if widths else s.width
                target_h = s.height if widths else int(median)
                emit(shape_id=s.shape_id, module="shape_size",
                     issue_type="shape_size.size_mismatch",
                     property="spPr.xfrm.ext",
                     old_value=f"{s.width}x{s.height}",
                     new_value=f"{target_w}x{target_h}",
                     profile_rule_id="shape_size.size_tolerance_emu",
                     message=f"{note} Match sibling "
                             f"{'widths' if widths else 'heights'}.")
    return out


def run_copilot(deck_bytes: bytes, thumbs: dict[int, bytes],
                manifest: dict) -> tuple[list[dict], int]:
    """Review up to MAX_SLIDES rendered slides; returns (new_records,
    slides_reviewed). API failures on individual slides are skipped so one
    bad call never sinks the run."""
    prs = Presentation(io.BytesIO(deck_bytes))
    existing = manifest.get("records") or []
    new_records: list[dict] = []
    reviewed = 0
    for s_idx, slide in enumerate(prs.slides):
        if reviewed >= MAX_SLIDES:
            break
        png = thumbs.get(s_idx)
        if png is None:
            continue
        inv = inventory(slide, prs.slide_width, prs.slide_height)
        if len(inv) < 3:
            continue
        try:
            observations = _ask_vision(png, inv)
        except Exception:
            continue
        reviewed += 1
        new_records.extend(
            synthesize(slide, s_idx, observations, existing + new_records))
    return new_records, reviewed
