"""AI-enrichment phase: classify clauses and fill Standard Assessment cols E–I.

This is phase 2 of the pipeline. Phase 1 (``router`` → ``standard_export``)
turns a PDF/URL into *items* — dicts ``{clause_id, title, text, classification}``
— and writes columns A–E. This module takes those items and, by calling a chosen
LLM (via :mod:`ai_providers`), fills:

    E  Classification           "Requirement" | "Information"  (AI; regex fallback)
    F  Requirement              organization-perspective requirement (Standard Format)
    G  Detailed Description      "<ReqID> | Verification: <v> | <Direct/Derived> — <trace>\\n<description>"
    H  Change in Requirement     concrete change the org must implement
    I  Requirement Classification one of a configurable controlled vocab (e.g. Product/Process/Other)

When a clause is ``Information`` (and ``fill_only_requirements`` is on, the
default) columns F–I are left blank — matching the requested behaviour.

Strategy: one **consolidated structured call per batch** of clauses (cheaper and
keeps E–I mutually consistent), parsed defensively with a JSON-repair retry and a
per-row regex fallback. A ``dry_run`` mode synthesises results with no network,
so the whole pipeline + UI + Excel write can be exercised without keys or cost.

Public API:
    EnrichConfig                       # all knobs (provider/model/prompts/...)
    DEFAULT_PROMPTS                    # editable per-field instruction defaults
    build_system_prompt(cfg) -> str    # the composed system prompt
    enrich(items, cfg, progress=None) -> list[dict]
    estimate_cost(items, cfg) -> dict
"""

from __future__ import annotations

import concurrent.futures as cf
import hashlib
import json
import logging
import math
import os
import re
import time
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from ai_providers import ProviderError, RateLimitError, get_provider
from standard_export import _REQUIREMENT_RE  # reuse the obligation-language regex

log = logging.getLogger("pdf2excel.ai")

# Enriched item keys this module adds (consumed by standard_export columns F–I).
ENRICHED_KEYS = (
    "requirement",
    "detailed_description",
    "change_in_requirement",
    "requirement_classification",
)

# Fixed enums for the strict-JSON contract.
_VERIFICATION = ["Inspection", "Test", "Demonstration", "Audit", "Design Review"]

# Default controlled vocabulary for column I (feeds the template's NIST cascade).
DEFAULT_VOCAB_I = ["Product", "Process", "Other"]


# --------------------------------------------------------------------------- #
# Default prompts (every field is independently editable via EnrichConfig)
# --------------------------------------------------------------------------- #

# "base" is the user-supplied Column-F analyst prompt, verbatim through STEP 4 +
# RULES. STEP 5 (output) is replaced at runtime by a strict-JSON contract that
# also carries the G/H/I field instructions below, so one call fills all of E–I.
DEFAULT_PROMPTS: Dict[str, str] = {
    "base": """ROLE
You are a compliance-to-requirements analyst. I will give you one or more
texts (clauses) from a standard, law, or regulation. Analyze EACH text
independently and produce exactly ONE classified output per text.

STEP 1 — CLASSIFY each text as exactly one of:
  • "Requirement" — the text imposes, directly or by clear implication, a
    binding, testable obligation that an organization can be made to satisfy.
    Cues: "shall", "must", "is required to", or a specific control / action /
    constraint.
  • "Information" — the text is a definition, scope statement, principle,
    aspirational / "endeavor to" wording, or background that does NOT imply any
    specific organizational obligation. Cues: "endeavor to", "should consider",
    aspirational goals, "obtain trust"; or it defines a term / scope / category.
  Pick ONE class per text. Do NOT split a text into separate requirement and
  information parts.

STEP 2 — SINGLE-OUTPUT RULE:
  Produce exactly ONE outcome per text:
  • Requirement → write ONE single requirement (no atomic a/b/c breakdown; if
    the text bundles several controls, capture them in one consolidated
    requirement).
  • Information → output Information only, with no requirement.

STEP 3 — IF REQUIREMENT, write it from the ORGANIZATION'S perspective:
  • Actor is always the organization (organization / controller / system),
    never the State, legislature, or regulator.
  • DIRECT  — the text already obligates the organization → restate it.
  • DERIVED — the text obligates a regulator/State, or is a principle, but
    clearly implies an organizational duty → translate it into that obligation
    and give a one-line trace.
  • FAITHFULNESS: the requirement must be traceable to the text. Do NOT add
    controls, thresholds, timelines, or scope the source does not support. If
    the text supports only a general duty, write a general but still testable one.
  • STANDARD FORMAT (testable): "The [organization/controller/system] shall
    [action] [object] [condition]."
  • Assign a sequential Requirement ID (REQ-001, REQ-002, …).
  • State a Verification Method: Inspection, Test, Demonstration, Audit, or
    Design Review.

STEP 4 — IF INFORMATION:
  • Mark "Information" and give a one-line reason it is not a testable
    requirement. No Requirement ID, no requirement text.

RULES:
  • De-duplicate identical texts.
  • Do not invent obligations not present or clearly implied by the source;
    justify every Derived requirement in the trace note.
  • Maintain a continuous Requirement ID sequence across all texts.""",

    # Column G — what complying entails (the composed cell also prepends ReqID /
    # Verification / Direct-Derived / trace; see _compose_detailed).
    "detailed_description": (
        "a 2-4 sentence plain-language description of what complying with this "
        "requirement entails (who/what is in scope and what evidence shows "
        "compliance). Faithful to the source — add no obligations it does not "
        'support. Use "" when Information.'
    ),

    # Column H — the gap-oriented action.
    "change_in_requirement": (
        "one sentence stating the concrete change or action the organization must "
        "implement to satisfy this requirement (gap-oriented; do not assume the "
        'current state). Use "" when Information.'
    ),

    # Column I — controlled-vocab classification feeding the NIST cascade.
    "requirement_classification": (
        "classify the requirement into exactly ONE of: {vocab}. Pick the single "
        'best fit. Use "" when Information.'
    ),
}


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #

@dataclass
class EnrichConfig:
    """All knobs for an enrichment run (every field is UI-/CLI-settable)."""

    provider: str = "claude"
    model: str = ""                       # "" → the provider's default model
    api_key: Optional[str] = None         # "" / None → provider env var
    temperature: float = 0.0
    max_tokens: int = 8000
    batch_size: int = 12                  # clauses per API call
    workers: int = 4                      # parallel batches
    retries: int = 5                      # retries on a 429 / rate-limit
    retry_base_delay: float = 2.0         # backoff base (s) when no server hint
    dry_run: bool = False                 # no network — synthesise placeholders
    fill_only_requirements: bool = True   # blank F–I for Information rows
    vocab_I: List[str] = field(default_factory=lambda: list(DEFAULT_VOCAB_I))
    prompts: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_PROMPTS))
    use_cache: bool = False
    cache_path: Optional[str] = None      # default: ~/.pdf2excel_cache.json


# --------------------------------------------------------------------------- #
# Prompt / schema construction
# --------------------------------------------------------------------------- #

def build_system_prompt(cfg: EnrichConfig) -> str:
    """Compose the full system prompt: the base analyst prompt + JSON contract."""
    p = cfg.prompts
    vocab = ", ".join(cfg.vocab_I) or "Other"
    verifications = ", ".join(_VERIFICATION)
    req_class_instr = p["requirement_classification"].replace("{vocab}", vocab)
    contract = (
        "STEP 5 — OUTPUT (STRICT JSON, MACHINE-READ):\n"
        "Do NOT output a table. Return ONLY a single JSON object of the form:\n"
        '  {"results": [ <one object per input text, in the SAME ORDER>, ... ]}\n'
        "Each object MUST have exactly these keys:\n"
        '  - "index": integer — the 1-based number of the input text.\n'
        '  - "classification": "Requirement" or "Information" (STEP 1).\n'
        '  - "req_id": e.g. "REQ-001" when Requirement, else "".\n'
        '  - "requirement": the organization-perspective requirement in STANDARD '
        'FORMAT when Requirement, else "" (STEP 3).\n'
        '  - "direct_derived": "Direct" or "Derived" when Requirement, else "".\n'
        f'  - "verification": one of {verifications} when Requirement, else "".\n'
        '  - "reason_trace": the one-line Derived trace (STEP 3), or the one-line '
        "reason a text is Information (STEP 4).\n"
        f'  - "detailed_description": {p["detailed_description"]}\n'
        f'  - "change_in_requirement": {p["change_in_requirement"]}\n'
        f'  - "requirement_classification": {req_class_instr}\n'
        "Output JSON only — no prose, no markdown, no code fences."
    )
    return p["base"].strip() + "\n\n" + contract


def _build_schema(vocab_I: List[str]) -> dict:
    """Strict JSON Schema for structured-output providers (Claude/OpenAI)."""
    item = {
        "type": "object",
        "properties": {
            "index": {"type": "integer"},
            "classification": {"type": "string",
                               "enum": ["Requirement", "Information"]},
            "req_id": {"type": "string"},
            "requirement": {"type": "string"},
            "direct_derived": {"type": "string",
                               "enum": ["Direct", "Derived", ""]},
            "verification": {"type": "string", "enum": _VERIFICATION + [""]},
            "reason_trace": {"type": "string"},
            "detailed_description": {"type": "string"},
            "change_in_requirement": {"type": "string"},
            "requirement_classification": {"type": "string",
                                           "enum": list(vocab_I) + [""]},
        },
        "required": [
            "index", "classification", "req_id", "requirement", "direct_derived",
            "verification", "reason_trace", "detailed_description",
            "change_in_requirement", "requirement_classification",
        ],
        "additionalProperties": False,
    }
    return {
        "type": "object",
        "properties": {"results": {"type": "array", "items": item}},
        "required": ["results"],
        "additionalProperties": False,
    }


def _format_batch(batch: List[dict]) -> str:
    """Render a batch of clauses as a numbered list for the user turn."""
    lines = []
    for i, it in enumerate(batch, start=1):
        text = (it.get("text") or "").strip()
        title = (it.get("title") or "").strip()
        prefix = f"[{i}]"
        lines.append(f"{prefix} {('(' + title + ') ') if title else ''}{text}")
    return (
        "TEXTS TO ANALYZE (numbered; analyze each independently):\n\n"
        + "\n\n".join(lines)
    )


# --------------------------------------------------------------------------- #
# JSON parsing / alignment (defensive)
# --------------------------------------------------------------------------- #

def _loads_loose(raw: str):
    """Parse JSON, tolerating code fences and surrounding prose."""
    if not raw:
        return None
    s = raw.strip()
    if s.startswith("```"):
        s = s.strip("`")
        # drop a leading "json" language tag if present
        nl = s.find("\n")
        if nl != -1 and s[:nl].strip().lower() in ("json", ""):
            s = s[nl + 1:]
    try:
        return json.loads(s)
    except (ValueError, TypeError):
        pass
    # Fall back to the first balanced { … } or [ … ] span.
    for opener, closer in (("{", "}"), ("[", "]")):
        start = s.find(opener)
        end = s.rfind(closer)
        if start != -1 and end > start:
            try:
                return json.loads(s[start:end + 1])
            except (ValueError, TypeError):
                continue
    return None


def _extract_results(data):
    """Pull the per-clause result list out of whatever JSON shape came back."""
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        if isinstance(data.get("results"), list):
            return data["results"]
        if "classification" in data:        # a single bare result object
            return [data]
    return None


def _align(results, n: int) -> List[dict]:
    """Return exactly ``n`` result dicts, keyed by ``index`` then positionally."""
    out: List[Optional[dict]] = [None] * n
    leftover: List[dict] = []
    for r in results or []:
        if not isinstance(r, dict):
            continue
        idx = r.get("index")
        if isinstance(idx, int) and 1 <= idx <= n and out[idx - 1] is None:
            out[idx - 1] = r
        else:
            leftover.append(r)
    # Fill any gaps positionally from leftovers, then with {} (regex fallback).
    li = 0
    for i in range(n):
        if out[i] is None:
            out[i] = leftover[li] if li < len(leftover) else {}
            li += 1 if li < len(leftover) else 0
    return [r or {} for r in out]


# --------------------------------------------------------------------------- #
# Result → enriched item mapping
# --------------------------------------------------------------------------- #

def _regex_classify(text: str) -> str:
    return "Requirement" if _REQUIREMENT_RE.search(text or "") else "Information"


def _norm_class(val, text: str) -> str:
    v = (val or "").strip().lower()
    if v.startswith("require"):
        return "Requirement"
    if v.startswith("inform"):
        return "Information"
    return _regex_classify(text)          # unusable value → deterministic fallback


def _validate_vocab(val, vocab: List[str]) -> str:
    v = (val or "").strip()
    for opt in vocab:
        if opt.lower() == v.lower():
            return opt
    # A Requirement should carry a class; default to the last vocab item ("Other").
    return vocab[-1] if vocab else v


def _compose_detailed(req_id: str, r: dict) -> str:
    """Fold ReqID / Verification / Direct-Derived / trace into column G."""
    head_bits: List[str] = []
    if req_id:
        head_bits.append(req_id)
    ver = (r.get("verification") or "").strip()
    if ver:
        head_bits.append(f"Verification: {ver}")
    dd = (r.get("direct_derived") or "").strip()
    trace = (r.get("reason_trace") or "").strip()
    if dd and trace:
        head_bits.append(f"{dd} — {trace}")
    elif dd:
        head_bits.append(dd)
    elif trace:
        head_bits.append(trace)
    head = " | ".join(head_bits)
    desc = (r.get("detailed_description") or "").strip()
    if head and desc:
        return f"{head}\n{desc}"
    return head or desc


def _finalize(items: List[dict], results: List[dict],
              cfg: EnrichConfig) -> List[dict]:
    """Map aligned results onto items, with a global REQ-NNN sequence."""
    out: List[dict] = []
    req_counter = 0
    for it, r in zip(items, results):
        r = r or {}
        cls = _norm_class(r.get("classification"), it.get("text", ""))
        e = dict(it)
        e["classification"] = cls
        if cls == "Information" and cfg.fill_only_requirements:
            e["requirement"] = ""
            e["detailed_description"] = ""
            e["change_in_requirement"] = ""
            e["requirement_classification"] = ""
        else:
            req_counter += 1
            req_id = f"REQ-{req_counter:03d}"
            e["requirement"] = (r.get("requirement") or "").strip()
            e["detailed_description"] = _compose_detailed(req_id, r)
            e["change_in_requirement"] = (r.get("change_in_requirement") or "").strip()
            e["requirement_classification"] = _validate_vocab(
                r.get("requirement_classification"), cfg.vocab_I
            )
        out.append(e)
    return out


_REQ_ID_RE = re.compile(r"^REQ-\d+")


def renumber_requirements(items) -> int:
    """Re-sequence the leading ``REQ-NNN`` of each Requirement's detailed
    description into one continuous run across ``items`` (in order).

    Useful after enriching a *subset* of rows and merging the results back into
    the full list, so requirement ids stay continuous and unique. Returns the
    number of requirements. Rows whose description was hand-edited to not start
    with ``REQ-`` are left untouched.
    """
    n = 0
    for it in items:
        if it.get("classification") != "Requirement":
            continue
        n += 1
        desc = it.get("detailed_description") or ""
        if _REQ_ID_RE.match(desc):
            it["detailed_description"] = _REQ_ID_RE.sub(f"REQ-{n:03d}", desc, count=1)
    return n


# --------------------------------------------------------------------------- #
# Provider calls
# --------------------------------------------------------------------------- #

def _complete_with_retry(provider, system: str, user: str,
                         cfg: EnrichConfig, schema: dict) -> str:
    """``provider.complete`` with retry/backoff on 429s (honours retry-after)."""
    attempt = 0
    while True:
        try:
            return provider.complete(
                system, user, model=cfg.model, temperature=cfg.temperature,
                max_tokens=cfg.max_tokens, schema=schema,
            )
        except RateLimitError as exc:
            attempt += 1
            if attempt > max(0, cfg.retries):
                raise ProviderError(
                    f"{exc}\n\nThe rate limit / quota did not clear after "
                    f"{cfg.retries} retries. Try: lower Workers to 1 and reduce "
                    "Batch size; choose a model with available quota (use 'List "
                    "models'); enable billing on your provider account; or switch "
                    "provider. (Dry-run needs no quota.)"
                ) from exc
            delay = exc.retry_after or cfg.retry_base_delay * (2 ** (attempt - 1))
            delay = min(float(delay) + 0.5, 60.0)
            log.warning("%s rate limited; retry %d/%d in %.1fs",
                        provider.name, attempt, cfg.retries, delay)
            time.sleep(delay)


def _call_batch(provider, system: str, schema: dict,
                batch: List[dict], cfg: EnrichConfig) -> List[dict]:
    """Call the provider for one batch; parse, repair-retry, align to len(batch)."""
    user = _format_batch(batch)
    raw = _complete_with_retry(provider, system, user, cfg, schema)
    results = _extract_results(_loads_loose(raw))
    if results is None:
        # One repair retry with an explicit JSON-only nudge.
        raw = _complete_with_retry(
            provider, system,
            user + "\n\nReturn ONLY the JSON object — no prose, no code fences.",
            cfg, schema)
        results = _extract_results(_loads_loose(raw))
        if results is None:
            log.warning("batch JSON unparseable; falling back to regex for %d "
                        "clauses", len(batch))
    return _align(results, len(batch))


def _dry_result(it: dict, cfg: EnrichConfig) -> dict:
    """Synthesise a result for dry-run (no network), reusing the regex classifier."""
    text = (it.get("text") or "").strip()
    if _regex_classify(text) == "Information":
        return {"classification": "Information",
                "reason_trace": "dry-run: no obligation language detected."}
    return {
        "classification": "Requirement",
        "requirement": f"[DRY-RUN] The organization shall comply with: {text[:120]}",
        "direct_derived": "Derived",
        "verification": "Audit",
        "reason_trace": "dry-run placeholder (no AI call).",
        "detailed_description": f"Dry-run description for: {text[:80]}",
        "change_in_requirement": "[DRY-RUN] Implement controls to satisfy this requirement.",
        "requirement_classification": cfg.vocab_I[0] if cfg.vocab_I else "",
    }


# --------------------------------------------------------------------------- #
# Simple optional on-disk cache (per clause)
# --------------------------------------------------------------------------- #

def _cache_file(cfg: EnrichConfig) -> str:
    return cfg.cache_path or os.path.join(
        os.path.expanduser("~"), ".pdf2excel_cache.json"
    )


def _cache_key(cfg: EnrichConfig, system: str, text: str) -> str:
    h = hashlib.sha256()
    h.update((cfg.provider + "\x00" + (cfg.model or "") + "\x00").encode("utf-8"))
    h.update((system + "\x00" + (text or "")).encode("utf-8"))
    return h.hexdigest()


def _load_cache(cfg: EnrichConfig) -> dict:
    try:
        with open(_cache_file(cfg), encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_cache(cfg: EnrichConfig, cache: dict) -> None:
    try:
        with open(_cache_file(cfg), "w", encoding="utf-8") as fh:
            json.dump(cache, fh)
    except OSError as exc:  # cache is best-effort
        log.warning("could not write cache: %s", exc)


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #

ProgressFn = Callable[[int, int, str], None]


def enrich(items, cfg: EnrichConfig,
           progress: Optional[ProgressFn] = None) -> List[dict]:
    """Enrich ``items`` (adds E–I fields) using ``cfg``. Returns new item dicts.

    ``progress(done, total, message)`` is called as batches complete. Raises
    :class:`ai_providers.ProviderError` if a (non-dry-run) provider call fails —
    the caller surfaces the message so the user can fix the key/model and re-run.
    """
    items = list(items)
    total = len(items)

    def report(done: int, msg: str = "") -> None:
        if progress:
            progress(done, total, msg)

    if total == 0:
        report(0, "no items")
        return []

    if cfg.dry_run:
        results = [_dry_result(it, cfg) for it in items]
        report(total, "dry-run complete")
        return _finalize(items, results, cfg)

    provider = get_provider(cfg.provider, cfg.api_key)
    system = build_system_prompt(cfg)
    schema = _build_schema(cfg.vocab_I)

    results: List[Optional[dict]] = [None] * total
    cache = _load_cache(cfg) if cfg.use_cache else None
    pending: List[int] = []
    if cache is not None:
        for i, it in enumerate(items):
            hit = cache.get(_cache_key(cfg, system, it.get("text", "")))
            if isinstance(hit, dict):
                results[i] = hit
            else:
                pending.append(i)
    else:
        pending = list(range(total))

    done = total - len(pending)
    report(done, "checked cache" if cache is not None else "starting")

    batches = [pending[b:b + cfg.batch_size]
               for b in range(0, len(pending), cfg.batch_size)]

    def run_one(idx_list):
        batch_items = [items[i] for i in idx_list]
        return idx_list, _call_batch(provider, system, schema, batch_items, cfg)

    if cfg.workers and cfg.workers > 1 and len(batches) > 1:
        with cf.ThreadPoolExecutor(max_workers=cfg.workers) as ex:
            futures = [ex.submit(run_one, b) for b in batches]
            for fut in cf.as_completed(futures):
                idx_list, batch_results = fut.result()   # propagates ProviderError
                for j, i in enumerate(idx_list):
                    results[i] = batch_results[j]
                done += len(idx_list)
                report(done, "enriching")
    else:
        for b in batches:
            idx_list, batch_results = run_one(b)
            for j, i in enumerate(idx_list):
                results[i] = batch_results[j]
            done += len(idx_list)
            report(done, "enriching")

    if cache is not None:
        for i in pending:
            if isinstance(results[i], dict):
                cache[_cache_key(cfg, system, items[i].get("text", ""))] = results[i]
        _save_cache(cfg, cache)

    return _finalize(items, [r or {} for r in results], cfg)


# --------------------------------------------------------------------------- #
# Cost estimation (rough, pre-run)
# --------------------------------------------------------------------------- #

# Approximate USD per 1M tokens (input, output). Editable; not authoritative.
_PRICES: Dict[str, tuple] = {
    "claude-opus-4-8": (5.0, 25.0),
    "claude-opus-4-7": (5.0, 25.0),
    "claude-sonnet-4-6": (3.0, 15.0),
    "claude-haiku-4-5": (1.0, 5.0),
    "gpt-4o": (2.5, 10.0),
    "gpt-4o-mini": (0.15, 0.6),
    "gpt-4.1": (2.0, 8.0),
    "gpt-4.1-mini": (0.4, 1.6),
    "gemini-1.5-pro": (1.25, 5.0),
    "gemini-1.5-flash": (0.075, 0.30),
    "gemini-2.0-flash": (0.10, 0.40),
}

_DEFAULT_PRICE = (3.0, 15.0)


def estimate_cost(items, cfg: EnrichConfig) -> dict:
    """Crude pre-run cost estimate (≈4 chars/token). Returns tokens + USD."""
    items = list(items)
    n = len(items)
    if n == 0:
        return {"input_tokens": 0, "output_tokens": 0, "usd": 0.0, "n_batches": 0}

    sys_tokens = max(1, len(build_system_prompt(cfg)) // 4)
    n_batches = math.ceil(n / max(1, cfg.batch_size))
    body_tokens = sum(len(it.get("text", "")) for it in items) // 4
    # The system prompt is (re)sent per batch; prompt caching makes repeats ~0.1x,
    # but we estimate full cost for a conservative figure.
    input_tokens = sys_tokens * n_batches + body_tokens + n_batches * 40
    output_tokens = sum(max(60, len(it.get("text", "")) // 3) for it in items)

    model = cfg.model or {
        "claude": "claude-opus-4-8", "openai": "gpt-4o",
        "gemini": "gemini-2.0-flash", "ollama": "llama3.1",
    }.get(cfg.provider, "")
    if cfg.provider == "ollama":
        in_price, out_price = 0.0, 0.0          # local — no API cost
    else:
        in_price, out_price = _PRICES.get(model, _DEFAULT_PRICE)
    usd = input_tokens / 1e6 * in_price + output_tokens / 1e6 * out_price
    return {
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "usd": round(usd, 4),
        "n_batches": n_batches,
        "model": model,
    }
