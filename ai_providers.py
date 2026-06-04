"""LLM provider abstraction for the AI-enrichment phase.

A small, uniform interface over several chat-completion backends — Anthropic
Claude, OpenAI, Google Gemini — so the enrichment orchestrator in
:mod:`ai_enrich` can call any of them the same way. Each provider's SDK is
imported **lazily inside the provider**, so importing this module never pulls in
an AI dependency; non-AI use of pdf2excel is completely unaffected.

The design mirrors the rest of the project: flat module, lazy imports, and
clear, actionable error messages (like ``web_extract.fetch_rendered``).

Public API:
    PROVIDERS                                  # {name: ProviderClass}
    list_providers() -> list[str]
    get_provider(name, api_key=None) -> ChatProvider
    ChatProvider.complete(system, user, *, model, temperature, max_tokens,
                          schema) -> str        # returns the assistant text

Every ``complete`` returns the raw assistant text. When ``schema`` is supplied
the provider requests strict JSON output where its SDK supports it; the caller
(:mod:`ai_enrich`) still parses and validates defensively, so a provider that
cannot enforce a schema only needs to honour ``response_mime_type``-style JSON.
"""

from __future__ import annotations

import os
import re
from typing import Dict, List, Optional, Type


class ProviderError(RuntimeError):
    """Raised for a missing SDK, a missing API key, or a failed provider call."""


class RateLimitError(ProviderError):
    """A 429 / quota-exceeded error, carrying the server's retry-after when known.

    A subclass of :class:`ProviderError`, so callers that only catch
    ``ProviderError`` still work; :mod:`ai_enrich` catches it specifically to
    retry with backoff.
    """

    def __init__(self, message: str, retry_after: Optional[float] = None) -> None:
        super().__init__(message)
        self.retry_after = retry_after


# Pull a retry delay out of an error: an HTTP `Retry-After` header, an SDK
# `retry_after`, "retry in 12.4s", or a gRPC "retry_delay { seconds: 12 }".
_RETRY_RES = (
    re.compile(r"retry[_ ]?(?:delay|after|in)\D{0,16}?(\d+(?:\.\d+)?)", re.I),
    re.compile(r"seconds:\s*(\d+)", re.I),
)


def _parse_retry_after(exc) -> Optional[float]:
    resp = getattr(exc, "response", None)
    if resp is not None:
        try:
            ra = resp.headers.get("retry-after")
            if ra:
                return float(ra)
        except Exception:  # noqa: BLE001 - header parsing is best-effort
            pass
    ra = getattr(exc, "retry_after", None)
    if isinstance(ra, (int, float)):
        return float(ra)
    for rx in _RETRY_RES:
        m = rx.search(str(exc))
        if m:
            try:
                return float(m.group(1))
            except ValueError:
                pass
    return None


def _is_rate_limit(exc) -> bool:
    code = getattr(exc, "status_code", None) or getattr(exc, "code", None)
    if code == 429:
        return True
    name = type(exc).__name__.lower()
    if any(k in name for k in ("ratelimit", "resourceexhausted", "toomanyrequests")):
        return True
    s = str(exc).lower()
    return ("429" in s or "quota" in s or "rate limit" in s
            or "resource exhausted" in s)


def _raise_provider_error(label: str, exc: Exception) -> None:
    """Raise a :class:`RateLimitError` for 429s, else a plain :class:`ProviderError`."""
    if _is_rate_limit(exc):
        raise RateLimitError(f"{label}: {exc}",
                             retry_after=_parse_retry_after(exc)) from exc
    raise ProviderError(f"{label}: {exc}") from exc


class ChatProvider:
    """Base class for a named chat-completion backend.

    Subclasses set the class attributes below and implement :meth:`complete`.
    The API key is resolved once at construction: an explicit ``api_key`` wins,
    otherwise the provider's environment variable is used.
    """

    name: str = ""
    env_key: str = ""           # environment variable that may hold the API key
    default_model: str = ""
    models: List[str] = []      # suggested model ids (UI dropdown; free text ok)
    requires_key: bool = True   # False for local backends (e.g. Ollama)

    def __init__(self, api_key: Optional[str] = None) -> None:
        self.api_key = (api_key or os.environ.get(self.env_key) or "").strip()

    # -- helpers ----------------------------------------------------------- #
    def available_models(self) -> List[str]:
        return list(self.models)

    def list_models(self) -> List[str]:
        """Live model ids from the provider's API.

        Override per provider to query the API (so the picker never goes stale as
        models are retired). Falls back to the static :attr:`models` list.
        """
        return list(self.models)

    def has_key(self) -> bool:
        return bool(self.api_key)

    def require_key(self) -> str:
        if not self.api_key:
            raise ProviderError(
                f"No API key for {self.name}. Enter it in the AI Configuration "
                f"tab, or set the {self.env_key} environment variable."
            )
        return self.api_key

    # -- interface --------------------------------------------------------- #
    def complete(
        self,
        system: str,
        user: str,
        *,
        model: str = "",
        temperature: float = 0.0,
        max_tokens: int = 8000,
        schema: Optional[dict] = None,
    ) -> str:
        """Return the assistant's text for ``system`` + ``user``.

        ``schema`` (a JSON Schema dict) requests strict JSON output where the
        provider supports it. Implementations raise :class:`ProviderError` for a
        missing SDK / key or an API failure.
        """
        raise NotImplementedError


# --------------------------------------------------------------------------- #
# Anthropic Claude
# --------------------------------------------------------------------------- #

class AnthropicProvider(ChatProvider):
    """Claude via the official ``anthropic`` SDK (Messages API)."""

    name = "claude"
    env_key = "ANTHROPIC_API_KEY"
    default_model = "claude-opus-4-8"
    models = [
        "claude-opus-4-8",      # most capable (default)
        "claude-sonnet-4-6",    # balanced — good default for high volume
        "claude-haiku-4-5",     # fastest / cheapest
        "claude-opus-4-7",
    ]

    @staticmethod
    def _rejects_sampling(model: str) -> bool:
        """Opus 4.7/4.8 remove ``temperature``/``top_p``/``top_k`` (400 if sent)."""
        return model.startswith(("claude-opus-4-8", "claude-opus-4-7"))

    def complete(self, system, user, *, model="", temperature=0.0,
                 max_tokens=8000, schema=None) -> str:
        try:
            import anthropic
        except ImportError:
            raise ProviderError(
                "The 'anthropic' package is not installed. Install it with: "
                "pip install anthropic"
            ) from None

        key = self.require_key()
        model = model or self.default_model
        client = anthropic.Anthropic(api_key=key)

        kwargs = dict(
            model=model,
            max_tokens=max_tokens,
            # Cache the large, static instruction prompt across the many batches
            # of one document (prefix match — the volatile clauses go in `user`).
            system=[{
                "type": "text",
                "text": system,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{"role": "user", "content": user}],
            # Classification/extraction: no thinking needed — faster and cheaper.
            thinking={"type": "disabled"},
        )
        # Sampling params 400 on Opus 4.7/4.8; only send them where accepted.
        if temperature is not None and not self._rejects_sampling(model):
            kwargs["temperature"] = temperature
        # Strict JSON via structured outputs (Opus 4.8 / Sonnet 4.6 / Haiku 4.5).
        if schema is not None:
            kwargs["output_config"] = {
                "format": {"type": "json_schema", "schema": schema}
            }

        try:
            resp = client.messages.create(**kwargs)
        except TypeError:
            # An older SDK may not accept `output_config`; drop it and rely on the
            # prompt's explicit JSON contract (the caller parses defensively).
            kwargs.pop("output_config", None)
            try:
                resp = client.messages.create(**kwargs)
            except anthropic.APIError as exc:
                _raise_provider_error("Claude API error", exc)
        except anthropic.APIError as exc:
            _raise_provider_error("Claude API error", exc)

        return "".join(
            b.text for b in resp.content if getattr(b, "type", "") == "text"
        )

    def list_models(self):
        try:
            import anthropic
        except ImportError:
            raise ProviderError(
                "The 'anthropic' package is not installed. Install it with: "
                "pip install anthropic"
            ) from None
        try:
            client = anthropic.Anthropic(api_key=self.require_key())
            ids = [m.id for m in client.models.list()]
        except anthropic.APIError as exc:
            _raise_provider_error("Claude API error", exc)
        return ids or list(self.models)


# --------------------------------------------------------------------------- #
# OpenAI
# --------------------------------------------------------------------------- #

class OpenAIProvider(ChatProvider):
    """OpenAI chat completions via the official ``openai`` SDK (v1+)."""

    name = "openai"
    env_key = "OPENAI_API_KEY"
    default_model = "gpt-4o"
    models = ["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"]

    def complete(self, system, user, *, model="", temperature=0.0,
                 max_tokens=8000, schema=None) -> str:
        try:
            from openai import OpenAI
        except ImportError:
            raise ProviderError(
                "The 'openai' package is not installed. Install it with: "
                "pip install openai"
            ) from None

        key = self.require_key()
        model = model or self.default_model
        client = OpenAI(api_key=key)

        if schema is not None:
            response_format = {
                "type": "json_schema",
                "json_schema": {
                    "name": "enrichment", "schema": schema, "strict": True,
                },
            }
        else:
            response_format = {"type": "json_object"}

        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                temperature=temperature,
                max_tokens=max_tokens,
                response_format=response_format,
            )
        except Exception as exc:  # noqa: BLE001 - openai error types vary by version
            _raise_provider_error("OpenAI API error", exc)

        return resp.choices[0].message.content or ""

    def list_models(self):
        try:
            from openai import OpenAI
        except ImportError:
            raise ProviderError(
                "The 'openai' package is not installed. Install it with: "
                "pip install openai"
            ) from None
        try:
            client = OpenAI(api_key=self.require_key())
            ids = sorted(m.id for m in client.models.list().data)
        except Exception as exc:  # noqa: BLE001 - openai error types vary by version
            _raise_provider_error("OpenAI API error", exc)
        # Keep chat-capable models (filter out embeddings/tts/whisper/etc.).
        chat = [i for i in ids if i.startswith(("gpt-", "o1", "o3", "o4", "chatgpt"))]
        return chat or ids or list(self.models)


# --------------------------------------------------------------------------- #
# Google Gemini
# --------------------------------------------------------------------------- #

class GeminiProvider(ChatProvider):
    """Gemini via the ``google-generativeai`` SDK."""

    name = "gemini"
    env_key = "GOOGLE_API_KEY"
    # Gemini 1.5 models were retired — default to current models, and use the
    # live "List models" action (list_models) when a key's availability differs.
    default_model = "gemini-2.0-flash"
    models = ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-2.5-pro"]

    def complete(self, system, user, *, model="", temperature=0.0,
                 max_tokens=8000, schema=None) -> str:
        try:
            import google.generativeai as genai
        except ImportError:
            raise ProviderError(
                "The 'google-generativeai' package is not installed. Install it "
                "with: pip install google-generativeai"
            ) from None

        key = self.require_key()
        model = model or self.default_model
        genai.configure(api_key=key)

        gen_config = {
            "temperature": temperature,
            "max_output_tokens": max_tokens,
            # JSON mode; the prompt also states the contract and the caller parses
            # defensively, so we don't depend on Gemini's response_schema support.
            "response_mime_type": "application/json",
        }
        model_obj = genai.GenerativeModel(
            model_name=model,
            system_instruction=system,
            generation_config=gen_config,
        )
        try:
            resp = model_obj.generate_content(user)
        except Exception as exc:  # noqa: BLE001 - genai error types vary by version
            _raise_provider_error("Gemini API error", exc)

        return getattr(resp, "text", "") or ""

    def list_models(self):
        try:
            import google.generativeai as genai
        except ImportError:
            raise ProviderError(
                "The 'google-generativeai' package is not installed. Install it "
                "with: pip install google-generativeai"
            ) from None
        try:
            genai.configure(api_key=self.require_key())
            ids = []
            for m in genai.list_models():
                methods = getattr(m, "supported_generation_methods", []) or []
                if "generateContent" in methods:
                    ids.append(m.name.split("/")[-1])  # "models/x" -> "x"
        except Exception as exc:  # noqa: BLE001 - genai error types vary by version
            _raise_provider_error("Gemini API error", exc)
        return ids or list(self.models)


# --------------------------------------------------------------------------- #
# Ollama (local models — no API key, no cost, no rate limits)
# --------------------------------------------------------------------------- #

class OllamaProvider(ChatProvider):
    """Local models via a running Ollama server (``pip install ollama``).

    No API key is needed — only a host (default ``http://localhost:11434``). The
    ``api_key`` slot is repurposed as an optional host override, and the live
    "List models" action shows exactly which models you've ``ollama pull``ed.
    """

    name = "ollama"
    env_key = "OLLAMA_HOST"          # holds a host, not a key
    requires_key = False
    default_model = "gemma4"
    models = ["gemma4", "gemma3", "llama3.1", "llama3.2", "qwen2.5", "mistral"]

    def __init__(self, api_key=None):
        super().__init__(api_key=api_key)   # api_key == optional host override
        self.host = self.api_key or "http://localhost:11434"

    def _client(self):
        try:
            import ollama
        except ImportError:
            raise ProviderError(
                "The 'ollama' package is not installed. Install it with: "
                "pip install ollama  — then run the Ollama app/server and "
                "`ollama pull <model>`."
            ) from None
        return ollama.Client(host=self.host)

    def complete(self, system, user, *, model="", temperature=0.0,
                 max_tokens=8000, schema=None):
        model = model or self.default_model
        client = self._client()
        try:
            resp = client.chat(
                model=model,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                # Force JSON when a schema is wanted; the prompt + defensive parser
                # do the rest (works across Ollama versions).
                format="json" if schema is not None else None,
                options={"temperature": temperature, "num_predict": max_tokens},
            )
        except Exception as exc:  # noqa: BLE001 - surface a setup-oriented hint
            raise ProviderError(
                f"Ollama error ({self.host}): {exc}. Is the Ollama server running "
                f"and the model pulled (`ollama pull {model}`)?"
            ) from exc
        try:
            return resp["message"]["content"]
        except (TypeError, KeyError, AttributeError):
            return getattr(getattr(resp, "message", None), "content", "") or ""

    def list_models(self):
        try:
            data = self._client().list()
        except ProviderError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise ProviderError(
                f"Ollama error ({self.host}): {exc}. Is the server running?"
            ) from exc
        raw = data.get("models") if isinstance(data, dict) else getattr(data, "models", [])
        out = []
        for m in raw or []:
            if isinstance(m, dict):
                name = m.get("model") or m.get("name")
            else:
                name = getattr(m, "model", None) or getattr(m, "name", None)
            if name:
                out.append(name)
        return out or list(self.models)


# --------------------------------------------------------------------------- #
# Registry
# --------------------------------------------------------------------------- #

PROVIDERS: Dict[str, Type[ChatProvider]] = {
    "claude": AnthropicProvider,
    "openai": OpenAIProvider,
    "gemini": GeminiProvider,
    "ollama": OllamaProvider,
}


def list_providers() -> List[str]:
    """Names of the registered providers, in display order."""
    return list(PROVIDERS)


def get_provider(name: str, api_key: Optional[str] = None) -> ChatProvider:
    """Instantiate a provider by name (``"claude"``/``"openai"``/``"gemini"``)."""
    try:
        cls = PROVIDERS[name]
    except KeyError:
        raise ProviderError(
            f"Unknown provider: {name!r}. Choose one of: {', '.join(PROVIDERS)}."
        ) from None
    return cls(api_key=api_key)
