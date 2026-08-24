"""The provider seam. Three passes ask a model for judgment; none of them
knows which vendor answered.

What is protected here is the contract the passes rely on, not any vendor's
wire format: a JSON answer matching a schema, and an EXCEPTION rather than an
empty answer when the model cannot be asked. Those two mean opposite things to
a designer reading "0 findings", and collapsing them is the failure mode this
file exists to prevent.
"""

import json

import pytest

import qc.llm as llm

SCHEMA = {"type": "object", "additionalProperties": False,
          "required": ["ok"], "properties": {"ok": {"type": "boolean"}}}


def test_gemini_is_the_default_provider():
    """Design lead, 24/08/2026. The switch lives in config and nowhere else."""
    from qc.config import LLM_MODEL, LLM_PROVIDER

    assert LLM_PROVIDER == "gemini"
    assert LLM_MODEL.startswith("gemini")


def test_the_model_choice_follows_the_provider(monkeypatch):
    """Pointing at a provider without also naming a model must not leave a
    Gemini build asking for a Claude model id."""
    import importlib

    monkeypatch.setenv("QC_LLM_PROVIDER", "anthropic")
    monkeypatch.delenv("QC_LLM_MODEL", raising=False)
    import qc.config as cfg

    importlib.reload(cfg)
    try:
        assert cfg.LLM_PROVIDER == "anthropic"
        assert cfg.LLM_MODEL.startswith("claude")
    finally:
        monkeypatch.delenv("QC_LLM_PROVIDER", raising=False)
        importlib.reload(cfg)


def test_an_unknown_provider_is_refused_not_guessed(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "nope")
    with pytest.raises(llm.LLMUnavailable) as exc:
        llm.ask_json(system="s", prompt="p", schema=SCHEMA)
    assert "nope" in str(exc.value)


def test_a_missing_key_raises_rather_than_returning_nothing(monkeypatch):
    """The distinction the whole file turns on. An empty dict would be read by
    every caller as "the model looked and found nothing"."""
    monkeypatch.setattr(llm, "LLM_PROVIDER", "gemini")
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    with pytest.raises(llm.LLMUnavailable):
        llm.ask_json(system="s", prompt="p", schema=SCHEMA)


def test_an_empty_answer_is_not_a_clean_slide(monkeypatch):
    """A blocked or truncated response comes back as no text. Treating that as
    an empty finding list would report a broken call as a clean deck."""
    monkeypatch.setattr(llm, "LLM_PROVIDER", "gemini")
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


def test_a_good_answer_comes_back_parsed(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "gemini")
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
    assert cfg.response_schema == SCHEMA, "the schema has to constrain the reply"
    assert cfg.temperature == 0.0, \
        "the same slide must not group two ways on two runs"
    # the image goes before the text, as the prompts describe
    parts = seen["contents"][0].parts
    assert parts[0].inline_data is not None and parts[-1].text == "ask"


def test_api_configured_answers_per_provider(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "gemini")
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    monkeypatch.setenv("ANTHROPIC_API_KEY", "irrelevant-here")
    assert llm.api_configured() is False, \
        "a Claude key does not configure a Gemini build"

    monkeypatch.setenv("GOOGLE_API_KEY", "k")
    assert llm.api_configured() is True

    monkeypatch.setattr(llm, "LLM_PROVIDER", "anthropic")
    assert llm.api_configured() is True
