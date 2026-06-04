"""Offline tests for ai_enrich — no network, via fake providers.

Run with pytest (``pytest tests/``) or directly (``python tests/test_ai_enrich.py``).
"""

import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ai_enrich as ae          # noqa: E402
import ai_providers as ap       # noqa: E402


class EchoProvider(ap.ChatProvider):
    """A deterministic fake: classifies by keyword, returns well-formed JSON."""

    name = "echo"
    env_key = "X"
    default_model = "echo"
    models = ["echo"]
    calls = 0

    def __init__(self, api_key=None):
        super().__init__(api_key="echo")

    def complete(self, system, user, *, model="", temperature=0.0,
                 max_tokens=8000, schema=None):
        type(self).calls += 1
        results = []
        for m in re.finditer(r"^\[(\d+)\]\s*(.*)$", user, flags=re.M):
            idx, text = int(m.group(1)), m.group(2)
            if re.search(r"\b(shall|must)\b", text, re.I):
                results.append({
                    "index": idx, "classification": "Requirement",
                    "req_id": f"REQ-{idx:03d}",
                    "requirement": f"The organization shall handle: {text[:40]}",
                    "direct_derived": "Direct", "verification": "Audit",
                    "reason_trace": "", "detailed_description": "what it entails",
                    "change_in_requirement": "implement it",
                    "requirement_classification": "Process",
                })
            else:
                results.append({
                    "index": idx, "classification": "Information",
                    "req_id": "", "requirement": "", "direct_derived": "",
                    "verification": "", "reason_trace": "background",
                    "detailed_description": "", "change_in_requirement": "",
                    "requirement_classification": "",
                })
        return json.dumps({"results": results})


class FencedProvider(EchoProvider):
    """Wraps EchoProvider output in code fences, shuffles, drops one index."""

    name = "fenced"

    def complete(self, system, user, **kw):
        data = json.loads(super().complete(system, user, **kw))
        data["results"] = [r for r in data["results"] if r["index"] != 3]  # drop #3
        data["results"].reverse()                                          # shuffle
        return "```json\n" + json.dumps(data) + "\n```"


class GarbageProvider(ap.ChatProvider):
    name = "garbage"
    env_key = "X"
    default_model = "g"
    models = ["g"]

    def __init__(self, api_key=None):
        super().__init__(api_key="g")

    def complete(self, system, user, **kw):
        return "Sorry, I can't help with that."   # never valid JSON


class FlakyProvider(EchoProvider):
    """Raises 429 a few times, then succeeds — exercises retry/backoff."""

    name = "flaky"
    fail_times = 2
    seen = 0

    def complete(self, system, user, **kw):
        if type(self).seen < type(self).fail_times:
            type(self).seen += 1
            raise ap.RateLimitError("429 quota exceeded", retry_after=0.01)
        return super().complete(system, user, **kw)


class AlwaysLimited(ap.ChatProvider):
    """Always 429 (like a free-tier limit: 0) — exercises the give-up message."""

    name = "limited"
    env_key = "X"
    default_model = "x"
    models = ["x"]

    def __init__(self, api_key=None):
        super().__init__(api_key="x")

    def complete(self, system, user, *, model="", temperature=0.0,
                 max_tokens=8000, schema=None):
        raise ap.RateLimitError("429 limit: 0 for model x", retry_after=0.01)


ap.PROVIDERS["echo"] = EchoProvider
ap.PROVIDERS["fenced"] = FencedProvider
ap.PROVIDERS["garbage"] = GarbageProvider
ap.PROVIDERS["flaky"] = FlakyProvider
ap.PROVIDERS["limited"] = AlwaysLimited

_SAMPLE = [
    "Operators shall keep records for five years.",   # 1 Requirement
    "This Regulation defines its scope.",             # 2 Information
    "The controller must encrypt personal data.",     # 3 Requirement
    "The following definitions apply.",               # 4 Information
    "Systems shall log all access events.",           # 5 Requirement
]


def _items():
    return [{"clause_id": f"C{i}", "title": "", "text": t, "classification": ""}
            for i, t in enumerate(_SAMPLE, start=1)]


def test_happy_path_batching_and_numbering():
    EchoProvider.calls = 0
    cfg = ae.EnrichConfig(provider="echo", batch_size=2, workers=1)
    out = ae.enrich(_items(), cfg)
    assert len(out) == 5
    assert EchoProvider.calls == 3            # 5 items / batch 2 -> 3 calls
    assert [o["classification"] for o in out] == \
        ["Requirement", "Information", "Requirement", "Information", "Requirement"]
    # Global REQ numbering across batches (model's per-batch ids are overridden).
    req_ids = [o["detailed_description"].split(" | ")[0]
               for o in out if o["classification"] == "Requirement"]
    assert req_ids == ["REQ-001", "REQ-002", "REQ-003"], req_ids
    # Information rows blank F–I; Requirement rows carry vocab + composed G.
    assert out[1]["requirement"] == "" and out[1]["requirement_classification"] == ""
    assert out[0]["requirement_classification"] == "Process"
    assert "Verification: Audit" in out[0]["detailed_description"]


def test_fenced_misaligned_with_dropped_row_falls_back():
    cfg = ae.EnrichConfig(provider="fenced", batch_size=10, workers=1)
    out = ae.enrich(_items(), cfg)
    # Row 3 ("must encrypt") was dropped from the response → regex fallback,
    # which still classifies it Requirement and assigns a continuous REQ id.
    assert out[2]["classification"] == "Requirement"
    assert out[2]["detailed_description"].startswith("REQ-")
    # The other rows still map correctly despite reversed order.
    assert out[0]["classification"] == "Requirement"
    assert out[1]["classification"] == "Information"


def test_garbage_response_falls_back_to_regex():
    cfg = ae.EnrichConfig(provider="garbage", batch_size=10, workers=1)
    out = ae.enrich(_items(), cfg)
    # Unparseable both times → keyword fallback for every row.
    assert [o["classification"] for o in out] == \
        ["Requirement", "Information", "Requirement", "Information", "Requirement"]
    # Fallback requirements get an id + class but no model-authored text.
    assert out[0]["detailed_description"].startswith("REQ-001")
    assert out[0]["requirement"] == ""
    assert out[0]["requirement_classification"] == "Other"  # vocab default (last)


def test_dry_run_needs_no_provider():
    cfg = ae.EnrichConfig(provider="does-not-exist", dry_run=True)
    out = ae.enrich(_items(), cfg)
    assert [o["classification"] for o in out] == \
        ["Requirement", "Information", "Requirement", "Information", "Requirement"]
    assert out[0]["requirement"].startswith("[DRY-RUN]")
    assert out[1]["detailed_description"] == ""   # Information blank


def test_information_fill_toggle():
    cfg = ae.EnrichConfig(provider="echo", fill_only_requirements=False, workers=1)
    out = ae.enrich(_items(), cfg)
    # With the toggle off we still map the model's (blank) values, not forced blanks.
    assert out[1]["classification"] == "Information"


def test_response_cache_avoids_recalls():
    import tempfile
    cache = os.path.join(tempfile.mkdtemp(prefix="pdf2excel_test_"), "cache.json")
    EchoProvider.calls = 0
    cfg = ae.EnrichConfig(provider="echo", batch_size=2, workers=1,
                          use_cache=True, cache_path=cache)
    out1 = ae.enrich(_items(), cfg)
    first = EchoProvider.calls
    out2 = ae.enrich(_items(), cfg)             # identical input -> all cached
    assert first == 3 and EchoProvider.calls - first == 0
    assert [o["classification"] for o in out1] == [o["classification"] for o in out2]


def test_rate_limit_retries_then_succeeds():
    FlakyProvider.seen = 0
    orig_sleep = ae.time.sleep
    ae.time.sleep = lambda *_a, **_k: None          # don't actually wait
    try:
        cfg = ae.EnrichConfig(provider="flaky", batch_size=10, workers=1, retries=5)
        out = ae.enrich(_items(), cfg)
        assert len(out) == 5
        assert FlakyProvider.seen == 2              # failed twice, then succeeded
        assert out[0]["classification"] == "Requirement"
    finally:
        ae.time.sleep = orig_sleep


def test_rate_limit_persistent_raises_actionable_error():
    orig_sleep = ae.time.sleep
    ae.time.sleep = lambda *_a, **_k: None
    try:
        cfg = ae.EnrichConfig(provider="limited", workers=1, retries=2)
        raised = None
        try:
            ae.enrich(_items(), cfg)
        except ae.ProviderError as exc:
            raised = exc
        assert raised is not None, "persistent 429 should raise"
        msg = str(raised).lower()
        assert "retries" in msg or "switch provider" in msg or "quota" in msg
    finally:
        ae.time.sleep = orig_sleep


def test_renumber_requirements_makes_ids_continuous():
    items = [
        {"classification": "Requirement",
         "detailed_description": "REQ-005 | Verification: Audit\nx"},
        {"classification": "Information", "detailed_description": ""},
        {"classification": "Requirement",
         "detailed_description": "REQ-001 | Direct\ny"},
        {"classification": "Requirement",
         "detailed_description": "hand-edited, no id"},
    ]
    n = ae.renumber_requirements(items)
    assert n == 3                                    # all Requirement rows counted
    assert items[0]["detailed_description"].startswith("REQ-001")
    assert items[2]["detailed_description"].startswith("REQ-002")
    assert items[1]["detailed_description"] == ""    # Information untouched
    assert items[3]["detailed_description"] == "hand-edited, no id"  # no prefix → left


def test_subset_enrich_then_merge_and_renumber():
    """Mimic the GUI's 'enrich a subset, merge back' flow."""
    base = _items()                                  # 5 clauses (R, I, R, I, R)
    working = [dict(it) for it in base]              # working copy (cols E–I blank)
    targets = [2, 4]                                 # enrich clauses #3 and #5 only
    sub = [base[i] for i in targets]
    cfg = ae.EnrichConfig(provider="echo", workers=1)
    enriched_sub = ae.enrich(sub, cfg)
    for pos, i in enumerate(targets):
        working[i] = enriched_sub[pos]
    n = ae.renumber_requirements(working)
    # Only the two targeted (both Requirement) rows are filled and numbered 1,2.
    assert working[2]["detailed_description"].startswith("REQ-001")
    assert working[4]["detailed_description"].startswith("REQ-002")
    assert working[0].get("requirement", "") == ""   # untouched row still blank
    assert n == 2


def test_ollama_provider_needs_no_key_and_is_free():
    g = ap.get_provider("ollama")
    assert g.requires_key is False
    assert g.host                                   # defaults to localhost
    est = ae.estimate_cost(_items(),
                           ae.EnrichConfig(provider="ollama", model="llama3.1"))
    assert est["usd"] == 0.0 and est["input_tokens"] > 0


def test_estimate_cost_positive():
    est = ae.estimate_cost(_items(), ae.EnrichConfig(provider="claude",
                                                     model="claude-sonnet-4-6"))
    assert est["usd"] > 0 and est["input_tokens"] > 0 and est["n_batches"] >= 1


def _run_all():
    fns = [v for k, v in sorted(globals().items())
           if k.startswith("test_") and callable(v)]
    for fn in fns:
        fn()
        print(f"  PASS {fn.__name__}")
    print(f"ai_enrich: {len(fns)} tests passed")


if __name__ == "__main__":
    _run_all()
