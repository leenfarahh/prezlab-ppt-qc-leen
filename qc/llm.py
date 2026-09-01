"""One door to the model, for every pass that needs judgment rather than
arithmetic.

Three passes ask a model something now - the design copilot, the component
review, and layout matching when applying a master - and each of them wants the
same shape of answer: a picture, some structured facts about it, a closed
question, and JSON back that conforms to a schema. Without a seam that is three
copies of client construction, schema plumbing, refusal handling and "what do we
do when there is no key", and they drift.

GEMINI IS THE ONLY MODEL THIS BUILD TALKS TO (design lead, 31/08/2026). There
was a second provider behind a `QC_LLM_PROVIDER` switch until then, and it went
because a switch nobody flips is a second code path nobody tests: the Anthropic
branch had its own schema dialect, its own refusal shape and its own key check,
all of it exercised by exactly zero runs of the real tool. A pass names its
schema and its question; it never imports an SDK and never sees a model id.
The tests of the passes above stay honest because they stub `ask_json` rather
than an SDK.

WHAT THE MODEL IS ASKED FOR, EVERYWHERE: a judgment from a closed vocabulary,
naming things by ids it was handed. Never a coordinate, never a measurement,
never a number that reaches the deck. Ask a model for EMU and you get plausible
EMU; ask it which of two lines was intended and you get the answer geometry
cannot compute. Every caller re-verifies what comes back against the real
geometry before it becomes a record.
"""

import json
import os
import re

from .config import (LLM_ATTEMPTS, LLM_CONCURRENCY, LLM_MODEL, LLM_MODEL_NOTE,
                     LLM_RETRY_MAX_DELAY_SEC, LLM_TIMEOUT_MS)


class LLMUnavailable(RuntimeError):
    """No usable credentials, or no SDK for the configured provider. Raised
    rather than returned so a caller cannot mistake it for "the model looked and
    found nothing" - the two mean very different things to a designer."""


def api_configured() -> bool:
    """Whether a judgment pass can run at all. Checked by the routes so they can
    say WHY the button did nothing instead of returning an empty review."""
    return bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY"))


def configuration_note() -> str:
    """What is wrong with this host's model settings, or "" when nothing is.

    Separate from api_configured() because a key can be perfectly present while
    QC_LLM_MODEL names a model this endpoint will not answer on, and the failure
    that produces is a 404 in every pass at once with nothing naming the
    variable responsible (qc.config.LLM_MODEL_NOTE)."""
    return LLM_MODEL_NOTE


# Keys that are valid JSON Schema and that the Gemini endpoint rejects outright:
#
#     400 INVALID_ARGUMENT ... Unknown name "additional_properties" at
#     'generation_config.response_schema': Cannot find field.
#
# The SDK does not filter them - it snake_cases them and posts them - so an
# otherwise valid JSON Schema 400s here. Stripped in the ADAPTER rather than in
# the eight schemas that declare them, because the schemas are the passes' own
# statement of what they will accept: `additionalProperties: False` says the
# pass rejects invented keys, and that statement is worth keeping in the source
# even on an endpoint that ignores it. This file is where a vendor's quirks are
# supposed to live.
_UNSUPPORTED_BY_GEMINI = ("additionalProperties", "additional_properties",
                          "$schema", "$defs", "definitions", "$ref")


def _for_gemini(schema):
    """`schema` with the keys this endpoint cannot parse removed, recursively.

    A copy, not an in-place edit: the caller's schema is a module-level constant
    that outlives the call, and stripping it once would silently rewrite what
    the pass says it accepts for every later reader of that constant.
    """
    if isinstance(schema, dict):
        return {k: _for_gemini(v) for k, v in schema.items()
                if k not in _UNSUPPORTED_BY_GEMINI}
    if isinstance(schema, list):
        return [_for_gemini(v) for v in schema]
    return schema


def _finish_reason(response) -> str:
    """Why the model stopped, as a bare uppercase name, or "" if it did not say.

    Read defensively through getattr rather than off a known shape: the field
    is an enum on a candidate the SDK may not populate at all (a blocked reply
    has no candidates), and a reader that assumes the happy shape turns "the
    answer was cut short" into an AttributeError three frames from anything
    that explains it.
    """
    candidates = getattr(response, "candidates", None) or []
    if not candidates:
        return ""
    reason = getattr(candidates[0], "finish_reason", None)
    if reason is None:
        return ""
    # An enum stringifies as "FinishReason.MAX_TOKENS"; a plain string does not.
    return str(getattr(reason, "name", reason)).rsplit(".", 1)[-1].upper()


def _quota_message(exc, code) -> str:
    """A 429 said in terms of what to DO about it, or "" if it is not one.

    "the model could not be reached (ClientError 429)" is true and it is the
    wrong advice: it reads like an outage to wait out. A 429 here is almost
    never that. Two quite different things arrive as one status code and they
    need opposite responses:

      limit: 0   the configured model has NO allowance on this project's tier -
                 a pro/preview model on a free key, say. Waiting changes
                 nothing, ever. It needs billing enabled or a different model.
      limit: n   a real rate limit, briefly hit. Waiting IS the answer, and the
                 API says how long.

    The distinction is read out of the error body rather than guessed, and the
    retry hint is quoted from the API rather than invented. Everything the
    designer needs to act is in the sentence, because the sentence is all they
    see (30/08/2026: a run came back "could not be reached" against a key whose
    free-tier quota for gemini-3.1-pro was zero, which reads as a network fault
    and is a billing setting).
    """
    if code != 429 and "429" not in str(code or ""):
        return ""
    body = str(exc)
    if "429" not in body and "RESOURCE_EXHAUSTED" not in body:
        return ""
    retry = ""
    match = re.search(r"[Pp]lease retry in ([0-9.]+)s", body)
    if match:
        try:
            retry = f" The API asks for {round(float(match.group(1)))}s."
        except ValueError:
            retry = ""
    if "limit: 0" in body:
        return (f"'{LLM_MODEL}' has no quota at all on this API key's tier "
                f"(the provider reports a limit of zero), so waiting will not "
                f"help. Enable billing on the project, or set QC_LLM_MODEL to a "
                f"model the key can call.")

    # A DAILY allowance is not a rate limit, and telling someone to "run the
    # pass again shortly" when they have spent the day's requests is advice
    # that cannot work. The two arrive as the same 429 with the same
    # RESOURCE_EXHAUSTED status and are told apart only by the quotaId in the
    # body: ...PerDayPerProjectPerModel-FreeTier is the day's ration,
    # ...PerMinute... is the blip the SDK's retries are for.
    #
    # Found 01/09/2026 on the free tier, where gemini-3.5-flash allows twenty
    # requests A DAY: the chat box, the copilot and the component review all
    # reported "run the pass again shortly" for hours, and every retry spent
    # nothing because there was nothing left to spend. The API's own
    # retry-in-25s hint is the wait until the next per-minute window, not until
    # the quota resets, so quoting it here is worse than saying nothing.
    per_day = re.search(r"quotaId['\"]?:\s*['\"]?([A-Za-z]*PerDay[A-Za-z]*)",
                        body)
    if per_day or "PerDay" in body:
        allowance = re.search(r"limit:\s*(\d+)", body)
        cap = f" of {allowance.group(1)} requests" if allowance else ""
        return (f"this API key's DAILY allowance{cap} for '{LLM_MODEL}' is "
                f"spent, so every pass that asks a model will fail until it "
                f"resets (midnight Pacific). This is not a fault in the deck "
                f"and not something retrying fixes: use a key with billing "
                f"enabled, or set QC_LLM_MODEL to a model with allowance left.")

    return (f"the model's rate limit was reached and three attempts did not "
            f"clear it.{retry} Nothing is wrong with the deck; run the pass "
            f"again shortly, or lower QC_LLM_CONCURRENCY.")


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

    # THE CLIENT MUST BE HELD IN A NAME UNTIL THE CALL RETURNS. `genai.Client`
    # closes its httpx connection pool in __del__, and
    # `genai.Client(api_key=key).models.generate_content(...)` keeps no
    # reference to the Client - only to `.models`, which holds the api client
    # rather than its owner. CPython therefore collects the Client the instant
    # the attribute is read, __del__ closes the pool, and the request that was
    # about to go out dies with
    #
    #     RuntimeError: Cannot send a request, as the client has been closed
    #
    # It reads like a race and is not one: 3 out of 3, every time (27/08/2026).
    # Every judgment pass in this package came through here, so every one of
    # them was failing and reporting it as its own polite nothing - the layout
    # proposal said "no layout could be proposed", the ask box said "that could
    # not be answered (RuntimeError)". One local name is the whole fix.
    #
    # Not cached across calls: a module-level client would be a global that
    # outlives a key change, and these calls take seconds each, so a fresh
    # connection pool per call costs nothing worth having.
    # ONE ATTEMPT USED TO BE ALL A PASS GOT. A 429 on slide 7 of 20 is not a
    # judgment that slide 7 is fine, but that is exactly how it read: the pass
    # caught the exception, moved on, and the deck came back reviewed. The SDK
    # retries the transient statuses (408, 429, 5xx) itself once told to, which
    # is better than a hand-rolled loop here because it honours Retry-After and
    # jitters the backoff. The timeout is the other half: without one, a
    # half-open connection pins this thread until the process dies.
    http = types.HttpOptions(
        timeout=LLM_TIMEOUT_MS,
        retry_options=types.HttpRetryOptions(
            attempts=LLM_ATTEMPTS, max_delay=LLM_RETRY_MAX_DELAY_SEC),
    )
    client = genai.Client(api_key=key, http_options=http)
    try:
        response = client.models.generate_content(
            model=LLM_MODEL,
            contents=[types.Content(role="user", parts=parts)],
            config=types.GenerateContentConfig(
                system_instruction=system,
                # Schema-constrained output, so a malformed answer is the SDK's
                # problem rather than a parse this file has to guess at.
                response_mime_type="application/json",
                response_schema=_for_gemini(schema),
                max_output_tokens=max_tokens,
                # Slide images are dense documents - 10pt legal lines, hairline
                # rules, two near-identical navies side by side - and at the
                # default resolution the small type is what gets lost. The
                # small type is where the formatting defects are.
                media_resolution=types.MediaResolution.MEDIA_RESOLUTION_HIGH,
                # NO temperature. This asked for 0.0 until 26/08/2026, on the
                # reasoning that the same slide should not group two ways on
                # two runs. Google's Gemini 3 guidance is the opposite: leave
                # it at the default of 1.0, because a lower value makes these
                # models loop and degrades exactly the multi-step reasoning
                # these passes need (ai.google.dev/gemini-api/docs/gemini-3).
                # Run-to-run stability was never really the sampler's job here
                # anyway - it comes from the closed vocabularies, from ids that
                # must exist in the inventory the model was handed, and from
                # code re-verifying every answer against real geometry before
                # it becomes a record.
                thinking_config=types.ThinkingConfig(thinking_level="high"),
                # No tools are passed here and none ever will be: every pass
                # asks for a judgment from a closed vocabulary and gets JSON
                # back. The SDK nevertheless defaults automatic function
                # calling ON when the field is unset, which sends each call
                # through its AFC loop - a deep copy of the config per
                # iteration, and a warning logged once per process telling us
                # to use Chat.send_message instead. Saying so explicitly takes
                # the direct path.
                automatic_function_calling=(
                    types.AutomaticFunctionCallingConfig(disable=True)),
            ),
        )
    except LLMUnavailable:
        raise
    except Exception as exc:
        # Every transport failure arrives here having already exhausted the
        # SDK's retries, so this is the outage, not the blip. It is still
        # "could not ask", never "asked and heard nothing" - and the type and
        # status are carried into the message because the last time a call
        # failed silently in this file it took two weeks to find (27/08/2026).
        code = getattr(exc, "code", None)
        quota = _quota_message(exc, code)
        if quota:
            raise LLMUnavailable(quota) from exc
        if code == 404:
            # A 404 from this endpoint means the MODEL ID does not exist, not
            # that the network is down, and the two need very different things
            # doing about them. qc.config's guard cannot catch this one: it only
            # knows another vendor's model id when it sees one, and
            # "gemini-3.1-flash" is Gemini-shaped, plausible, and not a model
            # (31/08/2026 - it was set on the LAN box and every pass in the tool
            # reported "the model could not be reached", which reads as an
            # outage and sent a designer looking at their wifi).
            raise LLMUnavailable(
                f"there is no model called {LLM_MODEL!r}. That is the "
                f"QC_LLM_MODEL setting in the .env file at the project root, "
                f"and the id has to match one the API key can list exactly - "
                f"'gemini-3.1-flash' looks right and does not exist, where "
                f"'gemini-3.5-flash' and 'gemini-3.1-flash-lite' do.") from exc
        raise LLMUnavailable(
            f"the model could not be reached "
            f"({type(exc).__name__}{f' {code}' if code else ''})") from exc

    # TRUNCATION IS NOT A SHORT ANSWER. `response_schema` constrains the SHAPE
    # of a reply, not that one arrived whole: thinking tokens are drawn from
    # the same max_output_tokens budget, so a hard slide can spend the budget
    # reasoning and get cut mid-object. That lands here as valid-looking text
    # that json.loads rejects - and a JSONDecodeError is not LLMUnavailable, so
    # before this guard it escaped the "skip this slide" contract every caller
    # is written against and came out of the route as a 500.
    if _finish_reason(response) == "MAX_TOKENS":
        raise LLMUnavailable(
            "the model ran out of output budget before it finished answering")

    text = (response.text or "").strip()
    if not text:
        # A blocked or empty answer is not a finding of "nothing wrong".
        raise LLMUnavailable("the model returned no answer for this slide")
    try:
        return json.loads(text)
    except ValueError as exc:
        raise LLMUnavailable(
            f"the model's answer was not the JSON it was asked for "
            f"({exc})") from exc


def ask_in_parallel(work: list, ask) -> list:
    """`ask(item)` for every item, a few at a time, results in INPUT ORDER.

    The passes above ask one closed question per slide, or per distinct slide
    structure, and no answer depends on another. Asking them in a row meant a
    twenty-slide review took twenty round trips end to end, nearly all of it
    spent waiting on a socket.

    Order is restored before returning, so nothing downstream can tell that the
    calls overlapped: records come out in slide order and a run is reproducible
    in the only sense that matters here.

    AN EXCEPTION COMES BACK AS AN EXCEPTION, in its slot, rather than being
    raised. Every caller of these passes is written to skip the slide it could
    not ask about and keep the rest - that contract is the whole reason
    LLMUnavailable exists - and a pool that raised on the first failure would
    turn one bad call back into a lost run.

    Threads rather than asyncio: the SDK calls are blocking, the routes are
    sync, and the work per item is one HTTPS request that releases the GIL.
    """
    if not work:
        return []
    if len(work) == 1 or LLM_CONCURRENCY <= 1:
        out = []
        for item in work:
            try:
                out.append(ask(item))
            except Exception as exc:
                out.append(exc)
        return out

    from concurrent.futures import ThreadPoolExecutor

    def _guarded(item):
        try:
            return ask(item)
        except Exception as exc:
            return exc

    with ThreadPoolExecutor(max_workers=min(LLM_CONCURRENCY,
                                            len(work))) as pool:
        return list(pool.map(_guarded, work))


def ask_json(*, system: str, prompt: str, schema: dict,
             images: list[bytes] | None = None,
             max_tokens: int = 8192) -> dict:
    """Ask the configured model a closed question and get JSON matching
    `schema`.

    Raises LLMUnavailable when it cannot be asked. Callers treat that as "skip
    this slide", never as "this slide is clean"."""
    return _gemini(system, prompt, images or [], schema, max_tokens)
