"""One door to the model, for every pass that needs judgment rather than
arithmetic.

Three passes ask a model something now - the design copilot, the component
review, and layout matching when applying a master - and each of them wants the
same shape of answer: a picture, some structured facts about it, a closed
question, and JSON back that conforms to a schema. Without a seam that is three
copies of client construction, schema plumbing, refusal handling and "what do we
do when there is no key", and they drift.

The provider lives HERE and nowhere else (design lead, 24/08/2026: use Gemini).
A pass names its schema and its question; it never imports an SDK, never sees a
model id, and never knows which vendor answered. Swapping providers is then this
file, and the tests of the passes above stay honest because they stub `ask_json`
rather than an SDK.

WHAT THE MODEL IS ASKED FOR, EVERYWHERE: a judgment from a closed vocabulary,
naming things by ids it was handed. Never a coordinate, never a measurement,
never a number that reaches the deck. Ask a model for EMU and you get plausible
EMU; ask it which of two lines was intended and you get the answer geometry
cannot compute. Every caller re-verifies what comes back against the real
geometry before it becomes a record.
"""

import json
import os

from .config import LLM_MODEL, LLM_PROVIDER


class LLMUnavailable(RuntimeError):
    """No usable credentials, or no SDK for the configured provider. Raised
    rather than returned so a caller cannot mistake it for "the model looked and
    found nothing" - the two mean very different things to a designer."""


def api_configured() -> bool:
    """Whether a judgment pass can run at all. Checked by the routes so they can
    say WHY the button did nothing instead of returning an empty review."""
    if LLM_PROVIDER == "gemini":
        return bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY"))
    if LLM_PROVIDER == "anthropic":
        return bool(os.getenv("ANTHROPIC_API_KEY"))
    return False


def _gemini(system: str, prompt: str, images: list[bytes], schema: dict,
            max_tokens: int) -> dict:
    try:
        from google import genai
        from google.genai import types
    except ImportError as exc:                     # pragma: no cover
        raise LLMUnavailable(
            "the google-genai package is not installed on this host") from exc

    key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not key:
        raise LLMUnavailable("GEMINI_API_KEY is not set on this host")

    parts = [types.Part.from_bytes(data=png, mime_type="image/png")
             for png in images]
    parts.append(types.Part.from_text(text=prompt))
    response = genai.Client(api_key=key).models.generate_content(
        model=LLM_MODEL,
        contents=[types.Content(role="user", parts=parts)],
        config=types.GenerateContentConfig(
            system_instruction=system,
            # Schema-constrained output, so a malformed answer is the SDK's
            # problem rather than a parse this file has to guess at.
            response_mime_type="application/json",
            response_schema=schema,
            max_output_tokens=max_tokens,
            # Deterministic-ish: the same slide should not group into different
            # components on two runs, or a designer cannot trust the review.
            temperature=0.0,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
        ),
    )
    text = (response.text or "").strip()
    if not text:
        # A blocked or empty answer is not a finding of "nothing wrong".
        raise LLMUnavailable("the model returned no answer for this slide")
    return json.loads(text)


def _anthropic(system: str, prompt: str, images: list[bytes], schema: dict,
               max_tokens: int) -> dict:
    """Kept because two passes shipped against it and a provider switch should
    not be a one-way door."""
    try:
        import anthropic
    except ImportError as exc:                     # pragma: no cover
        raise LLMUnavailable(
            "the anthropic package is not installed on this host") from exc
    import base64

    if not os.getenv("ANTHROPIC_API_KEY"):
        raise LLMUnavailable("ANTHROPIC_API_KEY is not set on this host")
    content = [{"type": "image",
                "source": {"type": "base64", "media_type": "image/png",
                           "data": base64.b64encode(png).decode()}}
               for png in images]
    content.append({"type": "text", "text": prompt})
    response = anthropic.Anthropic().messages.create(
        model=LLM_MODEL, max_tokens=max_tokens,
        thinking={"type": "adaptive"},
        output_config={"format": {"type": "json_schema", "schema": schema}},
        system=system, messages=[{"role": "user", "content": content}])
    if response.stop_reason == "refusal":
        raise LLMUnavailable("the model declined to answer for this slide")
    return json.loads("".join(b.text for b in response.content
                              if b.type == "text"))


_PROVIDERS = {"gemini": _gemini, "anthropic": _anthropic}


def ask_json(*, system: str, prompt: str, schema: dict,
             images: list[bytes] | None = None,
             max_tokens: int = 8192) -> dict:
    """Ask the configured model a closed question and get JSON matching
    `schema`.

    Raises LLMUnavailable when it cannot be asked. Callers treat that as "skip
    this slide", never as "this slide is clean"."""
    provider = _PROVIDERS.get(LLM_PROVIDER)
    if provider is None:
        raise LLMUnavailable(f"'{LLM_PROVIDER}' is not a provider this build "
                             f"knows (try gemini or anthropic)")
    return provider(system, prompt, images or [], schema, max_tokens)
