"""The model seam. Every pass that needs judgment asks through here.

What is protected is the contract the passes rely on, not the vendor's wire
format: a JSON answer matching a schema, and an EXCEPTION rather than an empty
answer when the model cannot be asked. Those two mean opposite things to a
designer reading "0 findings", and collapsing them is the failure mode this
file exists to prevent.

The second provider went on 31/08/2026 and the tests that pinned the switch
went with it. What replaced them is `test_only_gemini_is_reachable`: the point
is no longer that a build can be pointed at a vendor, it is that it cannot.
"""

import json

import pytest

import qc.llm as llm

SCHEMA = {"type": "object", "additionalProperties": False,
          "required": ["ok"], "properties": {"ok": {"type": "boolean"}}}


def test_gemini_is_the_only_model():
    """Design lead, 31/08/2026. One vendor, chosen in config and nowhere else.

    Asserts the REPO's default, not LLM_MODEL. LLM_MODEL is whatever this host's
    .env says, and a host is entitled to pin a different tier - the LAN box runs
    3.5-flash because 3.1-pro has no quota on its key. A test that read the live
    value would fail on the machine the tool actually runs on, which is the
    least useful place for it to fail."""
    from qc.config import DEFAULT_LLM_MODEL

    assert DEFAULT_LLM_MODEL.startswith("gemini-3.1-pro"), (
        "the vision passes are calibrated on Gemini 3.1 Pro (26/08/2026)")


def test_only_gemini_is_reachable():
    """No second provider, and no switch that could select one.

    The Anthropic branch was removed rather than left wired, so the guard is
    that the names are gone: a re-added branch has to come back through this
    test and state its case."""
    assert not hasattr(llm, "_PROVIDERS")
    assert not hasattr(llm, "_anthropic")
    assert not hasattr(llm, "LLM_PROVIDER")

    import qc.config as cfg

    assert not hasattr(cfg, "LLM_PROVIDER")
    assert not hasattr(cfg, "ASSIST_MODEL"), (
        "triage went through qc.llm with everything else on 31/08/2026")


def test_a_stale_model_id_is_named_rather_than_posted(monkeypatch):
    """A host whose .env still carries the other vendor's model id would 404 in
    every pass at once. The note says which variable to unset."""
    import importlib

    monkeypatch.setenv("QC_LLM_MODEL", "claude-opus-5")
    import qc.config as cfg

    importlib.reload(cfg)
    try:
        assert "QC_LLM_MODEL" in cfg.LLM_MODEL_NOTE
        assert "not a Gemini model" in cfg.LLM_MODEL_NOTE
    finally:
        monkeypatch.delenv("QC_LLM_MODEL", raising=False)
        importlib.reload(cfg)


def test_a_missing_key_raises_rather_than_returning_nothing(monkeypatch):
    """The distinction the whole file turns on. An empty dict would be read by
    every caller as "the model looked and found nothing"."""
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    with pytest.raises(llm.LLMUnavailable):
        llm.ask_json(system="s", prompt="p", schema=SCHEMA)


def test_an_empty_answer_is_not_a_clean_slide(monkeypatch):
    """A blocked or truncated response comes back as no text. Treating that as
    an empty finding list would report a broken call as a clean deck."""
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")

    class _Models:
        def generate_content(self, **kwargs):
            return type("R", (), {"text": ""})()

    class _Client:
        def __init__(self, **kwargs):
            self.models = _Models()

    from google import genai

    monkeypatch.setattr(genai, "Client", _Client)
    with pytest.raises(llm.LLMUnavailable):
        llm.ask_json(system="s", prompt="p", schema=SCHEMA)


def test_the_client_is_still_alive_when_the_request_goes_out(monkeypatch):
    """The bug that silently broke every judgment pass (27/08/2026).

    `genai.Client` closes its connection pool in __del__, and
    `genai.Client(key).models.generate_content(...)` holds no reference to the
    Client - so CPython collected it before the request was sent and every call
    died with "Cannot send a request, as the client has been closed". Each pass
    caught that and reported its own polite nothing, so the tool looked like it
    was working and finding little.

    Asserted the way the SDK actually fails: the stub client marks itself
    closed on __del__ and refuses to send, exactly as httpx does.
    """
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")

    class _Models:
        def __init__(self, owner):
            self._owner = owner

        def generate_content(self, **kwargs):
            if self._owner["closed"]:
                raise RuntimeError(
                    "Cannot send a request, as the client has been closed.")
            return type("R", (), {"text": json.dumps({"ok": True})})()

    class _Client:
        def __init__(self, **kwargs):
            # A plain dict rather than self, so the accessor cannot keep the
            # Client alive on the stub's behalf and hide the very thing this
            # test exists to catch.
            self._state = {"closed": False}
            self.models = _Models(self._state)

        def __del__(self):
            self._state["closed"] = True

    from google import genai

    monkeypatch.setattr(genai, "Client", _Client)
    assert llm.ask_json(system="s", prompt="p", schema=SCHEMA) == {"ok": True}


def test_a_good_answer_comes_back_parsed(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    seen = {}

    class _Models:
        def generate_content(self, **kwargs):
            seen.update(kwargs)
            return type("R", (), {"text": json.dumps({"ok": True})})()

    class _Client:
        def __init__(self, **kwargs):
            self.models = _Models()

    from google import genai

    monkeypatch.setattr(genai, "Client", _Client)
    out = llm.ask_json(system="sys", prompt="ask", schema=SCHEMA,
                       images=[b"png"])
    assert out == {"ok": True}

    cfg = seen["config"]
    assert cfg.system_instruction == "sys"
    assert cfg.response_mime_type == "application/json"
    # The schema still constrains the reply, minus the keys this endpoint
    # cannot parse: `additionalProperties` is valid JSON Schema and 400s here
    # (see _UNSUPPORTED_BY_GEMINI).
    assert cfg.response_schema == {
        "type": "object", "required": ["ok"],
        "properties": {"ok": {"type": "boolean"}}}
    assert SCHEMA["additionalProperties"] is False, (
        "the caller's schema is a shared constant and must not be edited in "
        "place on its way to the endpoint")
    assert cfg.temperature is None, (
        "Gemini 3 wants temperature left at its default; setting it low "
        "makes these models loop and degrades the reasoning these passes "
        "need")
    assert cfg.media_resolution == "MEDIA_RESOLUTION_HIGH", (
        "slides are dense documents; the small type is where the "
        "formatting defects are")
    # the image goes before the text, as the prompts describe
    parts = seen["contents"][0].parts
    assert parts[0].inline_data is not None and parts[-1].text == "ask"


def test_api_configured_reads_the_gemini_keys_only(monkeypatch):
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    monkeypatch.setenv("ANTHROPIC_API_KEY", "irrelevant-here")
    assert llm.api_configured() is False, \
        "another vendor's key configures nothing in this build"

    monkeypatch.setenv("GOOGLE_API_KEY", "k")
    assert llm.api_configured() is True

    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    monkeypatch.setenv("GEMINI_API_KEY", "k")
    assert llm.api_configured() is True


# ------------------------------------------------------ asking in parallel
#
# Each judgment pass asks one closed question per slide (or per distinct slide
# structure) and no answer depends on another, but they were asked in a row: a
# twenty-slide review was twenty round trips end to end, nearly all of it spent
# waiting on a socket (30/08/2026).


def test_answers_come_back_in_the_order_they_were_asked():
    """Records are emitted in slide order and a run has to be reproducible, so
    nothing downstream may be able to tell the calls overlapped."""
    import time

    from qc.llm import ask_in_parallel

    def _slow_for_early_items(n):
        # The first item is the slowest, so a pool returning completion order
        # rather than input order would put it last.
        time.sleep(0.05 if n == 0 else 0.0)
        return {"n": n}

    out = ask_in_parallel(list(range(8)), _slow_for_early_items)
    assert out == [{"n": i} for i in range(8)]


def test_one_failure_is_returned_in_its_slot_not_raised():
    """Every caller is written to skip the slide it could not ask about and
    keep the rest - that contract is why LLMUnavailable exists. A pool that
    raised on the first failure would turn one bad call back into a lost run."""
    from qc.llm import LLMUnavailable, ask_in_parallel

    def _fails_on_two(n):
        if n == 2:
            raise LLMUnavailable("the model could not be reached")
        return {"n": n}

    out = ask_in_parallel([0, 1, 2, 3], _fails_on_two)
    assert out[0] == {"n": 0} and out[3] == {"n": 3}
    assert isinstance(out[2], LLMUnavailable), (
        "the failure arrives as a value in its own slot, so the caller can skip "
        "exactly that slide")
    assert not isinstance(out[1], Exception)


def test_nothing_to_ask_makes_no_calls():
    from qc.llm import ask_in_parallel

    assert ask_in_parallel([], lambda x: pytest.fail("asked with no work")) == []


def test_the_calls_actually_overlap():
    import threading
    import time

    from qc.llm import ask_in_parallel

    peak = 0
    live = 0
    lock = threading.Lock()

    def _watch(_n):
        nonlocal peak, live
        with lock:
            live += 1
            peak = max(peak, live)
        time.sleep(0.05)
        with lock:
            live -= 1
        return True

    ask_in_parallel(list(range(8)), _watch)
    assert peak > 1, "the whole point is that they do not wait for each other"


# ------------------------------------------------------ a 429 that explains
#
# "the model could not be reached (ClientError 429)" is true and it is the wrong
# advice: it reads like an outage to wait out. Two different things arrive as
# that one status and they need opposite responses - a model with NO allowance
# on this tier (waiting never helps; it is a billing setting) and a rate limit
# briefly hit (waiting is exactly the answer). Real body, 30/08/2026.

_ZERO_QUOTA = (
    "429 RESOURCE_EXHAUSTED. {'error': {'code': 429, 'message': 'You exceeded "
    "your current quota, please check your plan and billing details. "
    "* Quota exceeded for metric: "
    "generativelanguage.googleapis.com/generate_content_free_tier_requests, "
    "limit: 0, model: gemini-3.1-pro\nPlease retry in 45.08927866s.'}}")

_RATE_LIMIT = (
    "429 RESOURCE_EXHAUSTED. {'error': {'code': 429, 'message': "
    "'Quota exceeded for metric: generate_content_free_tier_requests, "
    "limit: 15, model: gemini-3.5-flash\nPlease retry in 12.5s.'}}")


class _Err(Exception):
    def __init__(self, body):
        super().__init__(body)
        self.code = 429


def test_a_model_with_no_allowance_says_so_rather_than_looking_like_an_outage():
    from qc.llm import _quota_message

    msg = _quota_message(_Err(_ZERO_QUOTA), 429)
    assert "no quota at all" in msg
    assert "waiting will not help" in msg, (
        "a limit of zero never clears; telling a designer to retry sends them "
        "back to the same wall")
    assert "QC_LLM_MODEL" in msg, "and it names the setting that fixes it"


def test_an_ordinary_rate_limit_says_to_wait_and_how_long():
    from qc.llm import _quota_message

    msg = _quota_message(_Err(_RATE_LIMIT), 429)
    assert "rate limit" in msg
    assert "13s" in msg or "12s" in msg, "the wait is quoted from the API"
    assert "Nothing is wrong with the deck" in msg
    assert "no quota at all" not in msg


def test_anything_that_is_not_a_429_is_left_alone():
    from qc.llm import _quota_message

    assert _quota_message(_Err("500 INTERNAL"), 500) == ""
    assert _quota_message(Exception("timed out"), None) == ""
