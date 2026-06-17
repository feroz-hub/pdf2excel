"""Advanced Tkinter desktop GUI for pdf2excel — extract first, then AI-fill.

A thin shell over the library: phase 1 is :func:`router.convert` (source →
Standard Assessment, columns A–E); phase 2 is :func:`ai_enrich.enrich` +
:func:`standard_export.write_standard_assessment` (an LLM fills columns E–I).
No extraction or AI logic lives here — only orchestration and presentation.

Three tabs:
  1. Input & Extract  — choose source / output / mode / format, then Extract to
     preview the clauses (and write the base A–E workbook).
  2. AI Configuration — provider & model, API keys, batch / temperature / workers,
     toggles, the column-I vocabulary, and a fully editable prompt per column.
  3. Run & Results    — run the AI fill (progress + live log), edit any E–I cell,
     then Export the finished workbook.

Conversion and enrichment run on background threads; results arrive via a queue
polled on the Tk main loop, so the window stays responsive.

Run:  python pdf_to_excel_gui.py
"""

from __future__ import annotations

import os

# macOS ships a deprecated system Tk (8.5). Silence its load-time deprecation
# notice before tkinter initialises Tk. For a genuinely current Tk (8.6), run
# from a Python built against Tcl/Tk 8.6 (a conda or python.org build) — see the
# "Desktop GUI" note in the README.
os.environ.setdefault("TK_SILENCE_DEPRECATION", "1")

import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import config
from ai_enrich import (
    DEFAULT_PROMPTS,
    DEFAULT_VOCAB_I,
    EnrichConfig,
    enrich,
    estimate_cost,
    renumber_requirements,
)
from ai_providers import PROVIDERS, get_provider
from router import QualityGateError, _standard_writer_kwargs, convert
from standard_export import write_standard_assessment
from web_extract import clean_url

_PREVIEW_LIMIT = 100
_EXTRACT_PREVIEW_LIMIT = 100

FORMAT_MAP = {
    "Default": "default",
    "Default Excel": "default",
    "Standard Assessment": "standard",
}
MODE_MAP = {
    "Auto": "auto",
    "Prose": "prose",
    "Tables": "tables",
    "Standard / Structured": "standard",
    "NIST 800-53": "nist80053",
}
PROFILE_MAP = {
    "Auto": "auto",
    "Auto-detect": "auto",
    "Generic": "generic",
    "NIST 800-53": "nist80053",
    "Control Catalog": "control-catalog",
    "Control catalog": "control-catalog",
    "ISO-like": "iso",
    "Legal / Article": "legal",
    "Legal/article": "legal",
    "PCI": "pci",
    "PCI DSS": "pci",
    "CIS": "cis",
    "CIS Controls": "cis",
}
OCR_MAP = {"Off": "off", "Detect only": "detect", "off": "off", "detect": "detect"}

# Backward-compatible aliases used in combobox values
_FORMAT_LABELS = FORMAT_MAP
_MODE_LABELS = MODE_MAP
_PROFILE_LABELS = PROFILE_MAP
_OCR_LABELS = OCR_MAP
_RENDER_LABELS = {"Auto": "auto", "Always": "always", "Never": "never"}
_PROVIDER_LABELS = {"Claude (Anthropic)": "claude", "OpenAI": "openai",
                    "Gemini (Google)": "gemini", "Ollama (local)": "ollama"}
_PROVIDER_REVERSE = {v: k for k, v in _PROVIDER_LABELS.items()}

# Results-grid columns -> enriched-item field (clause is read-only context).
_RESULT_FIELDS = {
    "E Classification": "classification",
    "F Requirement": "requirement",
    "G Detailed Description": "detailed_description",
    "H Change in Requirement": "change_in_requirement",
    "I Req. Classification": "requirement_classification",
}
# The prompt editors shown on the AI tab (label -> DEFAULT_PROMPTS key).
_PROMPT_TABS = {
    "Base (Classify + Requirement → E, F)": "base",
    "Detailed Description (G)": "detailed_description",
    "Change in Requirement (H)": "change_in_requirement",
    "Requirement Classification (I)": "requirement_classification",
}


def _map_dropdown(label: str, mapping: dict, default: str) -> str:
    """Map a GUI combobox label to a router value (tolerant of spacing/case)."""
    text = (label or "").strip()
    if text in mapping:
        return mapping[text]
    lower = text.lower()
    for key, val in mapping.items():
        if key.lower() == lower:
            return val
    if "standard assessment" in lower or lower == "standard":
        return "standard"
    if "default" in lower:
        return "default"
    return default


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("pdf2excel — PDF/Web → Excel → AI")
        self.geometry("1100x820")
        self.minsize(920, 680)

        # Cross-phase state.
        self._items = []          # base items from phase 1 (A–E)
        self._export_items = []   # rows actually written to Standard Assessment
        self._meta = {}           # resolved Standard Assessment metadata
        self._enriched = []       # items after AI fill (E–I)
        self._prompt_texts = {}   # prompt key -> Text widget
        self._key_vars = {}       # provider -> StringVar for its API key
        self._quality_gate_ok = True

        self._events: "queue.Queue[tuple]" = queue.Queue()
        self._build_ui()
        self._load_settings()
        self.after(100, self._poll_events)

    # ===================================================================== #
    # UI construction
    # ===================================================================== #
    def _build_ui(self) -> None:
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.nb = nb

        self.tab_extract = ttk.Frame(nb, padding=10)
        self.tab_ai = ttk.Frame(nb, padding=10)
        self.tab_run = ttk.Frame(nb, padding=10)
        nb.add(self.tab_extract, text="1 · Input & Extract")
        nb.add(self.tab_ai, text="2 · AI Configuration")
        nb.add(self.tab_run, text="3 · Run & Results")

        self._build_extract_tab(self.tab_extract)
        self._build_ai_tab(self.tab_ai)
        self._build_run_tab(self.tab_run)

    # -- Tab 1: Input & Extract ------------------------------------------- #
    def _build_extract_tab(self, frm: ttk.Frame) -> None:
        pad = {"padx": 6, "pady": 3}
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(11, weight=1)

        ttk.Label(frm, text="PDF file or URL:").grid(row=0, column=0, sticky="w", **pad)
        self.pdf_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.pdf_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self._browse_pdf).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Output Excel:").grid(row=1, column=0, sticky="w", **pad)
        self.out_var = tk.StringVar()
        out_entry = ttk.Entry(frm, textvariable=self.out_var)
        out_entry.grid(row=1, column=1, sticky="ew", **pad)
        out_entry.bind("<FocusOut>", lambda _e: self._sync_review_path())
        ttk.Button(frm, text="Save as…", command=self._browse_out).grid(row=1, column=2, **pad)

        opts = ttk.Frame(frm)
        opts.grid(row=2, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Label(opts, text="Mode:").pack(side="left")
        self.mode_var = tk.StringVar(value="Auto")
        ttk.Combobox(opts, textvariable=self.mode_var,
                     values=["Auto", "Prose", "Tables", "Standard / Structured"],
                     state="readonly", width=16).pack(side="left", padx=(4, 12))
        ttk.Label(opts, text="Format:").pack(side="left")
        self.fmt_var = tk.StringVar(value="Standard Assessment")
        ttk.Combobox(opts, textvariable=self.fmt_var,
                     values=["Default Excel", "Standard Assessment"],
                     state="readonly", width=18).pack(side="left", padx=(4, 12))
        ttk.Label(opts, text="Profile:").pack(side="left")
        self.profile_var = tk.StringVar(value="Auto")
        ttk.Combobox(opts, textvariable=self.profile_var,
                     values=["Auto", "Generic", "NIST 800-53", "Control Catalog",
                             "ISO-like", "Legal / Article", "PCI", "CIS"],
                     state="readonly", width=14).pack(side="left", padx=4)

        meta1 = ttk.LabelFrame(frm, text="Standard metadata", padding=6)
        meta1.grid(row=3, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Label(meta1, text="Standard ID:").grid(row=0, column=0, sticky="w", padx=4)
        self.std_id_var = tk.StringVar(value="MLSR")
        ttk.Entry(meta1, textvariable=self.std_id_var, width=14).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(meta1, text="Title:").grid(row=0, column=2, sticky="w", padx=4)
        self.std_title_var = tk.StringVar()
        ttk.Entry(meta1, textvariable=self.std_title_var, width=24).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Label(meta1, text="Edition:").grid(row=0, column=4, sticky="w", padx=4)
        self.std_edition_var = tk.StringVar()
        ttk.Entry(meta1, textvariable=self.std_edition_var, width=12).grid(row=0, column=5, sticky="w", padx=4)
        meta1.columnconfigure(3, weight=1)

        meta2 = ttk.Frame(meta1)
        meta2.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(6, 0))
        ttk.Label(meta2, text="Document ID:").pack(side="left")
        self.doc_id_var = tk.StringVar()
        ttk.Entry(meta2, textvariable=self.doc_id_var, width=16).pack(side="left", padx=(4, 12))
        ttk.Label(meta2, text="Name:").pack(side="left")
        self.doc_name_var = tk.StringVar()
        ttk.Entry(meta2, textvariable=self.doc_name_var, width=20).pack(side="left", padx=(4, 12))
        ttk.Label(meta2, text="Revision:").pack(side="left")
        self.doc_rev_var = tk.StringVar()
        ttk.Entry(meta2, textvariable=self.doc_rev_var, width=10).pack(side="left", padx=4)

        struct = ttk.LabelFrame(frm, text="Structured extraction & review", padding=6)
        struct.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)
        self.gen_review_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(struct, text="Generate review workbook",
                        variable=self.gen_review_var,
                        command=self._sync_review_path).grid(row=0, column=0, sticky="w", padx=4)
        ttk.Label(struct, text="Review output:").grid(row=0, column=1, sticky="e", padx=(8, 4))
        self.review_var = tk.StringVar()
        ttk.Entry(struct, textvariable=self.review_var, width=36).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(struct, text="Browse…", command=self._browse_review).grid(row=0, column=3, padx=4)
        struct.columnconfigure(2, weight=1)

        row1 = ttk.Frame(struct)
        row1.grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self.show_issues_var = tk.BooleanVar(value=True)
        self.force_export_var = tk.BooleanVar(value=False)
        self.export_low_conf_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row1, text="Show Extraction_Issues sheet",
                        variable=self.show_issues_var).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(row1, text="Force export if quality gate fails",
                        variable=self.force_export_var).pack(side="left", padx=10)
        ttk.Checkbutton(row1, text="Export low-confidence rows",
                        variable=self.export_low_conf_var).pack(side="left", padx=10)

        incl = ttk.LabelFrame(frm, text="Include page sections", padding=6)
        incl.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)
        self.include_front_var = tk.BooleanVar(value=False)
        self.include_toc_var = tk.BooleanVar(value=False)
        self.include_appendix_var = tk.BooleanVar(value=False)
        self.include_refs_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(incl, text="Front matter", variable=self.include_front_var).pack(side="left", padx=8)
        ttk.Checkbutton(incl, text="TOC", variable=self.include_toc_var).pack(side="left", padx=8)
        ttk.Checkbutton(incl, text="Appendix", variable=self.include_appendix_var).pack(side="left", padx=8)
        ttk.Checkbutton(incl, text="References", variable=self.include_refs_var).pack(side="left", padx=8)

        tune = ttk.Frame(frm)
        tune.grid(row=6, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Label(tune, text="OCR mode:").pack(side="left")
        self.ocr_display_var = tk.StringVar(value="Off")
        ttk.Combobox(tune, textvariable=self.ocr_display_var, values=list(_OCR_LABELS),
                     state="readonly", width=10).pack(side="left", padx=(4, 12))
        ttk.Label(tune, text="Min confidence:").pack(side="left")
        self.min_conf_var = tk.StringVar(value="0.0")
        ttk.Entry(tune, textvariable=self.min_conf_var, width=6).pack(side="left", padx=(4, 12))
        ttk.Label(tune, text="Gap factor:").pack(side="left")
        self.gap_var = tk.DoubleVar(value=1.6)
        ttk.Scale(tune, from_=1.1, to=3.0, variable=self.gap_var, orient="horizontal",
                  length=120, command=lambda _=None: self._gap_label.config(
                      text=f"{self.gap_var.get():.2f}")).pack(side="left", padx=4)
        self._gap_label = ttk.Label(tune, text="1.60")
        self._gap_label.pack(side="left")
        ttk.Label(tune, text="Render JS:").pack(side="left", padx=(12, 4))
        self.render_var = tk.StringVar(value="Auto")
        ttk.Combobox(tune, textvariable=self.render_var, values=list(_RENDER_LABELS),
                     state="readonly", width=7).pack(side="left", padx=4)
        self.insecure_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(tune, text="Allow insecure TLS",
                        variable=self.insecure_var).pack(side="left", padx=(12, 0))
        self.skip_cover_var = tk.BooleanVar(value=True)

        act = ttk.Frame(frm)
        act.grid(row=7, column=0, columnspan=3, sticky="w", **pad)
        self.extract_btn = ttk.Button(act, text="Run Extraction",
                                      command=self._on_extract)
        self.extract_btn.pack(side="left")
        self.extract_prog = ttk.Progressbar(act, mode="indeterminate", length=220)
        self.extract_prog.pack(side="left", padx=12)

        self.extract_status = ttk.Label(
            frm, text="Choose a PDF/URL, set Standard Assessment + Profile Auto, then Run Extraction.",
            foreground="#444", wraplength=900)
        self.extract_status.grid(row=8, column=0, columnspan=3, sticky="w", **pad)

        prev = ttk.LabelFrame(frm, text="Extract preview (first rows)", padding=4)
        prev.grid(row=11, column=0, columnspan=3, sticky="nsew", **pad)
        prev.rowconfigure(0, weight=1)
        prev.columnconfigure(0, weight=1)
        cols = ("clause", "title", "class", "text", "conf", "issues")
        self.extract_tree = ttk.Treeview(prev, columns=cols, show="headings", height=8)
        headings = {
            "clause": "Clause ID", "title": "Title", "class": "Classification",
            "text": "Text Preview", "conf": "Confidence", "issues": "Issues",
        }
        widths = {"clause": 90, "title": 140, "class": 95, "text": 380, "conf": 70, "issues": 160}
        for c in cols:
            self.extract_tree.heading(c, text=headings[c])
            self.extract_tree.column(c, width=widths[c], anchor="w", stretch=(c == "text"))
        self.extract_tree.grid(row=0, column=0, sticky="nsew")
        ys = ttk.Scrollbar(prev, orient="vertical", command=self.extract_tree.yview)
        ys.grid(row=0, column=1, sticky="ns")
        self.extract_tree.configure(yscrollcommand=ys.set)
        self.extract_tree.tag_configure("warning", background="#fff3cd")

    # -- Tab 2: AI Configuration ------------------------------------------ #
    def _build_ai_tab(self, frm: ttk.Frame) -> None:
        pad = {"padx": 6, "pady": 3}
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(2, weight=1)

        # -- provider / model / key --
        top = ttk.LabelFrame(frm, text="Provider, model & API key", padding=8)
        top.grid(row=0, column=0, sticky="ew", **pad)
        top.columnconfigure(5, weight=1)

        ttk.Label(top, text="Provider:").grid(row=0, column=0, sticky="w", **pad)
        self.provider_var = tk.StringVar(value="Claude (Anthropic)")
        prov_cb = ttk.Combobox(top, textvariable=self.provider_var,
                               values=list(_PROVIDER_LABELS), state="readonly", width=18)
        prov_cb.grid(row=0, column=1, sticky="w", **pad)
        prov_cb.bind("<<ComboboxSelected>>", lambda _e: self._on_provider_change())

        ttk.Label(top, text="Model:").grid(row=0, column=2, sticky="w", **pad)
        self.model_var = tk.StringVar()
        self.model_cb = ttk.Combobox(top, textvariable=self.model_var, width=24)
        self.model_cb.grid(row=0, column=3, sticky="w", **pad)
        ttk.Button(top, text="List models", command=self._on_list_models).grid(
            row=0, column=4, sticky="w", **pad)

        self.key_label = ttk.Label(top, text="API key:")
        self.key_label.grid(row=1, column=0, sticky="w", **pad)
        self.key_var = tk.StringVar()
        self.key_entry = ttk.Entry(top, textvariable=self.key_var, show="•", width=46)
        self.key_entry.grid(row=1, column=1, columnspan=3, sticky="ew", **pad)
        self.key_status = ttk.Label(top, text="", foreground="#888")
        self.key_status.grid(row=1, column=4, sticky="w", **pad)
        ttk.Button(top, text="Save settings", command=self._on_save_settings).grid(
            row=1, column=5, sticky="e", **pad)

        # -- numeric / toggle options --
        mid = ttk.LabelFrame(frm, text="Options", padding=8)
        mid.grid(row=1, column=0, sticky="ew", **pad)
        self.batch_var = tk.IntVar(value=12)
        self.workers_var = tk.IntVar(value=4)
        self.maxtok_var = tk.IntVar(value=8000)
        self.temp_var = tk.DoubleVar(value=0.0)
        self.retries_var = tk.IntVar(value=5)
        ttk.Label(mid, text="Batch size:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Spinbox(mid, from_=1, to=50, textvariable=self.batch_var, width=6).grid(
            row=0, column=1, sticky="w", **pad)
        ttk.Label(mid, text="Workers:").grid(row=0, column=2, sticky="w", **pad)
        ttk.Spinbox(mid, from_=1, to=16, textvariable=self.workers_var, width=6).grid(
            row=0, column=3, sticky="w", **pad)
        ttk.Label(mid, text="Max tokens:").grid(row=0, column=4, sticky="w", **pad)
        ttk.Spinbox(mid, from_=512, to=64000, increment=512, textvariable=self.maxtok_var,
                    width=8).grid(row=0, column=5, sticky="w", **pad)
        ttk.Label(mid, text="Temperature:").grid(row=0, column=6, sticky="w", **pad)
        ttk.Spinbox(mid, from_=0.0, to=1.0, increment=0.1, textvariable=self.temp_var,
                    width=6, format="%.1f").grid(row=0, column=7, sticky="w", **pad)
        ttk.Label(mid, text="Retries:").grid(row=0, column=8, sticky="w", **pad)
        ttk.Spinbox(mid, from_=0, to=10, textvariable=self.retries_var, width=5).grid(
            row=0, column=9, sticky="w", **pad)

        self.dryrun_var = tk.BooleanVar(value=False)
        self.req_only_var = tk.BooleanVar(value=True)
        self.cache_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(mid, text="Dry-run (no API calls)",
                        variable=self.dryrun_var).grid(row=1, column=0, columnspan=2, sticky="w", **pad)
        ttk.Checkbutton(mid, text="Fill F–I only for Requirements",
                        variable=self.req_only_var).grid(row=1, column=2, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(mid, text="Use response cache",
                        variable=self.cache_var).grid(row=1, column=5, columnspan=2, sticky="w", **pad)
        ttk.Label(mid, text="Column I vocabulary:").grid(row=2, column=0, sticky="w", **pad)
        self.vocab_var = tk.StringVar(value=", ".join(DEFAULT_VOCAB_I))
        ttk.Entry(mid, textvariable=self.vocab_var, width=40).grid(
            row=2, column=1, columnspan=5, sticky="w", **pad)

        # -- per-column prompt editors --
        pf = ttk.LabelFrame(frm, text="Prompts (editable per column)", padding=6)
        pf.grid(row=2, column=0, sticky="nsew", **pad)
        pf.rowconfigure(0, weight=1)
        pf.columnconfigure(0, weight=1)
        pnb = ttk.Notebook(pf)
        pnb.grid(row=0, column=0, sticky="nsew")
        for label, key in _PROMPT_TABS.items():
            sub = ttk.Frame(pnb)
            txt = scrolledtext.ScrolledText(sub, wrap="word", height=8, width=90)
            txt.pack(fill="both", expand=True)
            txt.insert("1.0", DEFAULT_PROMPTS[key])
            self._prompt_texts[key] = txt
            pnb.add(sub, text=label)
        prow = ttk.Frame(pf)
        prow.grid(row=1, column=0, sticky="w")
        ttk.Button(prow, text="Reset prompts to defaults",
                   command=self._reset_prompts).pack(side="left", pady=4)
        ttk.Button(prow, text="Estimate cost",
                   command=self._on_estimate_cost).pack(side="left", padx=10)
        self.cost_label = ttk.Label(prow, text="", foreground="#06c")
        self.cost_label.pack(side="left", padx=6)

        self._on_provider_change()

    # -- Tab 3: Run & Results --------------------------------------------- #
    def _build_run_tab(self, frm: ttk.Frame) -> None:
        pad = {"padx": 6, "pady": 4}
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(4, weight=1)

        act = ttk.Frame(frm)
        act.grid(row=0, column=0, sticky="ew", **pad)
        self.generate_btn = ttk.Button(act, text="② Generate with AI",
                                       command=self._on_generate, state="disabled")
        self.generate_btn.pack(side="left")
        ttk.Label(act, text="Rows:").pack(side="left", padx=(12, 2))
        self.rows_var = tk.StringVar(value="All")
        ttk.Combobox(act, textvariable=self.rows_var,
                     values=["All", "Selected", "First N"], state="readonly",
                     width=9).pack(side="left")
        self.rows_n_var = tk.StringVar(value="5")
        ttk.Entry(act, textvariable=self.rows_n_var, width=5).pack(side="left", padx=4)
        self.run_prog = ttk.Progressbar(act, mode="determinate", length=170)
        self.run_prog.pack(side="left", padx=10)
        self.run_pct = ttk.Label(act, text="")
        self.run_pct.pack(side="left")
        self.export_btn = ttk.Button(act, text="③ Export to Excel",
                                     command=self._on_export, state="disabled")
        self.export_btn.pack(side="right")
        self.open_btn = ttk.Button(act, text="Open file",
                                   command=self._open_output, state="disabled")
        self.open_btn.pack(side="right", padx=8)

        self.run_status = ttk.Label(frm, text="Extract first (tab 1), then generate.",
                                    foreground="#444")
        self.run_status.grid(row=1, column=0, sticky="w", **pad)

        ext_prev = ttk.LabelFrame(frm, text="Exported items preview", padding=4)
        ext_prev.grid(row=2, column=0, sticky="nsew", **pad)
        ext_prev.rowconfigure(0, weight=1)
        ext_prev.columnconfigure(0, weight=1)
        pcols = ("clause", "title", "class", "text", "conf", "issues")
        self.run_preview_tree = ttk.Treeview(ext_prev, columns=pcols, show="headings", height=6)
        for c, h, w in (
            ("clause", "Clause ID", 90), ("title", "Title", 130),
            ("class", "Classification", 90), ("text", "Text Preview", 340),
            ("conf", "Confidence", 70), ("issues", "Issues", 150),
        ):
            self.run_preview_tree.heading(c, text=h)
            self.run_preview_tree.column(c, width=w, anchor="w", stretch=(c == "text"))
        self.run_preview_tree.grid(row=0, column=0, sticky="nsew")
        pys = ttk.Scrollbar(ext_prev, orient="vertical", command=self.run_preview_tree.yview)
        pys.grid(row=0, column=1, sticky="ns")
        self.run_preview_tree.configure(yscrollcommand=pys.set)
        self.run_preview_tree.tag_configure("warning", background="#fff3cd")

        ttk.Label(frm, text="Log:").grid(row=3, column=0, sticky="w", padx=6)
        self.log = scrolledtext.ScrolledText(frm, height=5, wrap="word", state="disabled")
        self.log.grid(row=3, column=0, sticky="ew", padx=6, pady=(20, 4))

        res = ttk.LabelFrame(frm, text="AI enrichment results (E–I)", padding=4)
        res.grid(row=4, column=0, sticky="nsew", **pad)
        res.rowconfigure(0, weight=1)
        res.columnconfigure(0, weight=1)
        cols = ("clause", *_RESULT_FIELDS)
        self.result_tree = ttk.Treeview(res, columns=cols, show="headings",
                                        height=12, selectmode="extended")
        widths = {"clause": 90, "E Classification": 95, "F Requirement": 280,
                  "G Detailed Description": 250, "H Change in Requirement": 200,
                  "I Req. Classification": 110}
        for c in cols:
            self.result_tree.heading(c, text=c)
            self.result_tree.column(c, width=widths.get(c, 150), anchor="w")
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        ys = ttk.Scrollbar(res, orient="vertical", command=self.result_tree.yview)
        ys.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=ys.set)
        self.result_tree.tag_configure("warning", background="#ffebeb")
        self.result_tree.tag_configure("low_conf", background="#fffde6")
        self.result_tree.bind("<Double-1>", self._on_result_edit)
        ttk.Label(frm, text="Double-click an E–I cell to edit. For Rows = Selected, "
                  "multi-select rows in this grid before Generate.",
                  foreground="#888").grid(row=5, column=0, sticky="w", padx=6)

    # ===================================================================== #
    # Settings persistence
    # ===================================================================== #
    def _current_provider(self) -> str:
        return _PROVIDER_LABELS.get(self.provider_var.get(), "claude")

    def _on_provider_change(self) -> None:
        prov = self._current_provider()
        models = get_provider(prov).available_models()
        self.model_cb["values"] = models
        if self.model_var.get() not in models:
            self.model_var.set(models[0] if models else "")
        # Show the saved/env key (or, for local backends, the optional host).
        self.key_var.set(self._key_vars.get(prov, tk.StringVar()).get()
                         if prov in self._key_vars else "")
        needs_key = getattr(PROVIDERS.get(prov), "requires_key", True)
        self.key_label.config(text="API key:" if needs_key else "Host (optional):")
        self.key_entry.config(show="•" if needs_key else "")
        if not needs_key:
            self.key_status.config(text="local — no key needed (host optional)")
        else:
            present = bool(config.get_api_key(prov))
            self.key_status.config(
                text="✓ key on file" if (present or self.key_var.get()) else "no key set")

    def _on_list_models(self) -> None:
        """Query the selected provider's API for the models its key can use."""
        prov = self._current_provider()
        key = self.key_var.get().strip() or config.get_api_key(prov)
        if getattr(PROVIDERS.get(prov), "requires_key", True) and not key:
            messagebox.showerror("pdf2excel", f"Enter an API key for {prov} first.")
            return
        self.key_status.config(text="listing models…")
        threading.Thread(target=self._list_models_worker, args=(prov, key),
                         daemon=True).start()

    def _list_models_worker(self, prov, key) -> None:
        try:
            models = get_provider(prov, key).list_models()
            self._events.put(("models_done", models))
        except Exception as exc:  # noqa: BLE001
            self._events.put(("models_err", str(exc)))

    def _load_settings(self) -> None:
        data = config.load()
        prov = data.get("provider", "claude")
        self.provider_var.set(_PROVIDER_REVERSE.get(prov, "Claude (Anthropic)"))
        if data.get("model"):
            self.model_var.set(data["model"])
        for key, default in (("batch_size", 12), ("workers", 4),
                             ("max_tokens", 8000)):
            getattr(self, {"batch_size": "batch_var", "workers": "workers_var",
                           "max_tokens": "maxtok_var"}[key]).set(data.get(key, default))
        self.temp_var.set(data.get("temperature", 0.0))
        self.retries_var.set(data.get("retries", 5))
        self.req_only_var.set(data.get("fill_only_requirements", True))
        self.cache_var.set(data.get("use_cache", False))
        if data.get("vocab_I"):
            self.vocab_var.set(", ".join(data["vocab_I"]))
        for key, txt in self._prompt_texts.items():
            saved = (data.get("prompts") or {}).get(key)
            if saved:
                txt.delete("1.0", "end")
                txt.insert("1.0", saved)
        # Pre-fill key vars from config/env for each provider.
        for p in PROVIDERS:
            self._key_vars[p] = tk.StringVar(value=config.get_api_key(p))
        self._on_provider_change()
        self.key_var.set(self._key_vars[self._current_provider()].get())

    def _collect_prompts(self) -> dict:
        return {k: t.get("1.0", "end").strip() for k, t in self._prompt_texts.items()}

    def _on_save_settings(self) -> None:
        prov = self._current_provider()
        # Remember the typed key for this provider.
        self._key_vars.setdefault(prov, tk.StringVar()).set(self.key_var.get().strip())
        data = config.load()
        data.update(
            provider=prov,
            model=self.model_var.get().strip(),
            batch_size=int(self.batch_var.get()),
            workers=int(self.workers_var.get()),
            retries=int(self.retries_var.get()),
            max_tokens=int(self.maxtok_var.get()),
            temperature=float(self.temp_var.get()),
            fill_only_requirements=bool(self.req_only_var.get()),
            use_cache=bool(self.cache_var.get()),
            vocab_I=self._vocab_list(),
            prompts=self._collect_prompts(),
        )
        keys = data.setdefault("api_keys", {})
        for p, var in self._key_vars.items():
            v = var.get().strip()
            if v:
                keys[p] = v
        try:
            config.save(data)
        except RuntimeError as exc:
            messagebox.showerror("pdf2excel", str(exc))
            return
        self.key_status.config(text="✓ key on file" if self.key_var.get() else "no key set")
        messagebox.showinfo("pdf2excel", f"Settings saved to {config.config_path()}")

    def _reset_prompts(self) -> None:
        for key, txt in self._prompt_texts.items():
            txt.delete("1.0", "end")
            txt.insert("1.0", DEFAULT_PROMPTS[key])

    def _vocab_list(self) -> list:
        return [v.strip() for v in self.vocab_var.get().split(",") if v.strip()] \
            or list(DEFAULT_VOCAB_I)

    # ===================================================================== #
    # File pickers
    # ===================================================================== #
    def _default_review_path(self, out_path: str) -> str:
        base, ext = os.path.splitext(out_path)
        return f"{base}_review{ext or '.xlsx'}"

    def _sync_review_path(self) -> None:
        if not self.gen_review_var.get():
            return
        out = self.out_var.get().strip()
        if out and not self.review_var.get().strip():
            self.review_var.set(self._default_review_path(out))

    def _browse_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose a PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.pdf_var.set(path)
            if not self.out_var.get():
                self.out_var.set(os.path.splitext(path)[0] + ".xlsx")
            self._sync_review_path()

    def _browse_out(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Excel as", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.out_var.set(path)
            self._sync_review_path()

    def _browse_review(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Review sheet as", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.review_var.set(path)

    # ===================================================================== #
    # Phase 1: extract  (background thread)
    # ===================================================================== #
    def _parse_min_confidence(self) -> float | None:
        raw = self.min_conf_var.get().strip()
        try:
            val = float(raw)
        except ValueError:
            return None
        if val < 0.0 or val > 1.0:
            return None
        return val

    def _collect_extract_kwargs(self) -> dict | None:
        """Build router.convert kwargs from GUI fields; return None if validation fails."""
        min_conf = self._parse_min_confidence()
        if min_conf is None:
            messagebox.showerror("pdf2excel", "Minimum confidence must be a number between 0.0 and 1.0.")
            return None

        out = self.out_var.get().strip()
        review_output = None
        if self.gen_review_var.get():
            review_output = self.review_var.get().strip() or (
                self._default_review_path(out) if out else ""
            )
            if not review_output:
                messagebox.showerror("pdf2excel", "Set an output Excel path or review workbook path.")
                return None

        fmt = _map_dropdown(self.fmt_var.get(), FORMAT_MAP, "standard")
        mode = _map_dropdown(self.mode_var.get(), MODE_MAP, "auto")
        profile = _map_dropdown(self.profile_var.get(), PROFILE_MAP, "auto")

        # Standard Assessment must use structured pipeline — never prose/table dump.
        if fmt == "standard" and mode in ("prose", "tables"):
            mode = "auto"

        kwargs = dict(
            mode=mode,
            fmt=fmt,
            profile=profile,
            gap_factor=float(self.gap_var.get()),
            standard_id=self.std_id_var.get().strip() or "MLSR",
            standard_title=self.std_title_var.get().strip(),
            standard_edition=self.std_edition_var.get().strip(),
            document_id=self.doc_id_var.get().strip(),
            document_name=self.doc_name_var.get().strip(),
            document_revision=self.doc_rev_var.get().strip(),
            insecure=bool(self.insecure_var.get()),
            render=_RENDER_LABELS.get(self.render_var.get(), "auto"),
            ocr_mode=_map_dropdown(self.ocr_display_var.get(), OCR_MAP, "off"),
            min_confidence=min_conf,
            review_output=review_output,
            show_issues=bool(self.show_issues_var.get()),
            force_export=bool(self.force_export_var.get()),
            export_low_confidence=bool(self.export_low_conf_var.get()),
            include_front_matter=bool(self.include_front_var.get()),
            include_toc=bool(self.include_toc_var.get()),
            include_appendix=bool(self.include_appendix_var.get()),
            include_references=bool(self.include_refs_var.get()),
            skip_cover=bool(self.skip_cover_var.get()),
            raise_on_quality_gate=False,
        )
        debug_labels = {
            "source": self.pdf_var.get().strip(),
            "output": out,
            "format_label": self.fmt_var.get(),
            "mode_label": self.mode_var.get(),
            "profile_label": self.profile_var.get(),
        }
        return kwargs, debug_labels

    def _log_extract_options(self, out: str, kwargs: dict, labels: dict) -> None:
        lines = [
            "Extraction options:",
            f"  Source: {labels.get('source', '')}",
            f"  Output: {out}",
            f"  Format: {labels.get('format_label', '')} -> {kwargs.get('fmt')}",
            f"  Mode: {labels.get('mode_label', '')} -> {kwargs.get('mode')}",
            f"  Profile: {labels.get('profile_label', '')} -> {kwargs.get('profile')}",
            f"  Review output: {kwargs.get('review_output') or '(none)'}",
            f"  Show issues: {kwargs.get('show_issues')}",
            f"  Force export: {kwargs.get('force_export')}",
            f"  Export low confidence: {kwargs.get('export_low_confidence')}",
            f"  Include front matter: {kwargs.get('include_front_matter')}",
            f"  Include TOC: {kwargs.get('include_toc')}",
            f"  Include appendix: {kwargs.get('include_appendix')}",
            f"  Include references: {kwargs.get('include_references')}",
            f"  OCR mode: {kwargs.get('ocr_mode')}",
            f"  Minimum confidence: {kwargs.get('min_confidence')}",
        ]
        for line in lines:
            self._log(line)

    def _on_extract(self) -> None:
        src = self.pdf_var.get().strip()
        out = self.out_var.get().strip()
        is_url = "://" in src
        if is_url:
            try:
                src = clean_url(src)
                self.pdf_var.set(src)
            except ValueError as exc:
                messagebox.showerror("pdf2excel", str(exc))
                return
        elif not src or not os.path.isfile(src):
            messagebox.showerror("pdf2excel", "Choose a valid PDF file or enter a URL.")
            return
        if not out:
            base = "" if is_url else os.path.splitext(os.path.basename(src))[0]
            out = (base or "web_export") + ".xlsx"
            self.out_var.set(out)
        self._sync_review_path()

        kwargs = self._collect_extract_kwargs()
        if kwargs is None:
            return
        convert_kwargs, debug_labels = kwargs

        self._log_clear()
        self._log_extract_options(out, convert_kwargs, debug_labels)
        self.extract_btn.config(state="disabled")
        self.extract_prog.start(12)
        self.extract_status.config(text="Extracting…")
        threading.Thread(target=self._extract_worker, args=(src, out, convert_kwargs),
                         daemon=True).start()

    def _extract_worker(self, src, out, kwargs) -> None:
        try:
            convert_kwargs = {k: v for k, v in kwargs.items() if not k.startswith("_")}
            result = convert(src, out, **convert_kwargs)
            self._events.put(("extract_done", result))
        except QualityGateError as exc:
            if exc.result is not None:
                self._events.put(("extract_qg_failed", exc.result, exc.failures))
            else:
                self._events.put(("extract_err", str(exc)))
        except Exception as exc:  # noqa: BLE001
            self._events.put(("extract_err", str(exc)))

    # ===================================================================== #
    # Phase 2: enrich  (background thread)
    # ===================================================================== #
    def _gather_config(self) -> EnrichConfig:
        prov = self._current_provider()
        return EnrichConfig(
            provider=prov,
            model=self.model_var.get().strip(),
            api_key=self.key_var.get().strip() or config.get_api_key(prov),
            temperature=float(self.temp_var.get()),
            max_tokens=int(self.maxtok_var.get()),
            batch_size=int(self.batch_var.get()),
            workers=int(self.workers_var.get()),
            retries=int(self.retries_var.get()),
            dry_run=bool(self.dryrun_var.get()),
            fill_only_requirements=bool(self.req_only_var.get()),
            vocab_I=self._vocab_list(),
            prompts=self._collect_prompts(),
            use_cache=bool(self.cache_var.get()),
        )

    def _selected_indices(self) -> list:
        """Which item indices to enrich, per the Rows selector."""
        n = len(self._items)
        mode = self.rows_var.get()
        if mode == "Selected":
            idxs = sorted(self.result_tree.index(r)
                          for r in self.result_tree.selection())
            return [i for i in idxs if 0 <= i < n]
        if mode == "First N":
            try:
                k = int(self.rows_n_var.get())
            except (TypeError, ValueError):
                k = 0
            return list(range(min(max(k, 0), n)))
        return list(range(n))   # All

    def _on_generate(self) -> None:
        if not self._items:
            messagebox.showinfo("pdf2excel", "Run extraction first (tab 1).")
            return
        if not getattr(self, "_quality_gate_ok", True):
            messagebox.showerror(
                "pdf2excel",
                "Quality gate did not pass. Fix extraction issues or enable "
                "Force export before running AI enrichment.",
            )
            return
        targets = self._selected_indices()
        if not targets:
            messagebox.showinfo(
                "pdf2excel",
                "No rows to enrich. Pick rows in the grid and set Rows = Selected, "
                "or choose Rows = All / First N.")
            return
        cfg = self._gather_config()
        needs_key = getattr(PROVIDERS.get(cfg.provider), "requires_key", True)
        if not cfg.dry_run and needs_key and not cfg.api_key:
            messagebox.showerror(
                "pdf2excel",
                f"No API key for {cfg.provider}. Enter one on the AI Configuration "
                f"tab (or enable Dry-run).")
            self.nb.select(self.tab_ai)
            return
        self._enrich_targets = targets
        sublist = [self._items[i] for i in targets]
        self.generate_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.run_prog.config(value=0, maximum=len(sublist))
        self._log_clear()
        self._log(f"Enriching {len(sublist)} of {len(self._items)} clauses via "
                  f"{cfg.provider}{' (dry-run)' if cfg.dry_run else ''}…")
        threading.Thread(target=self._enrich_worker, args=(cfg, sublist),
                         daemon=True).start()

    def _enrich_worker(self, cfg, sublist) -> None:
        def progress(done, total, msg):
            self._events.put(("enrich_progress", done, total, msg))
        try:
            enriched = enrich(sublist, cfg, progress=progress)
            self._events.put(("enrich_done", enriched))
        except Exception as exc:  # noqa: BLE001
            self._events.put(("enrich_err", str(exc)))

    # ===================================================================== #
    # Phase 3: export
    # ===================================================================== #
    def _on_export(self) -> None:
        items = self._enriched or self._items
        if not items:
            return
        out = self.out_var.get().strip()
        meta = dict(self._meta)
        meta.setdefault("standard_id", self.std_id_var.get().strip() or "MLSR")
        # After AI enrichment, export all rows; otherwise only validated export rows.
        export_items = None
        if self._enriched and self._enriched != self._items:
            export_items = None
        elif self._export_items:
            export_items = self._export_items
        try:
            kwargs = _standard_writer_kwargs(**meta)
            write_standard_assessment(
                items, out, export_items=export_items, **kwargs,
                show_issues=bool(self.show_issues_var.get()),
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("pdf2excel", f"Export failed:\n{exc}")
            return
        self.open_btn.config(state="normal")
        n_req = sum(1 for it in items if it.get("classification") == "Requirement")
        self.run_status.config(text=f"Exported {len(items)} rows "
                               f"({n_req} requirements) → {out}")
        messagebox.showinfo("pdf2excel", f"Saved {len(items)} rows to:\n{out}")

    def _open_output(self) -> None:
        out = self.out_var.get().strip()
        if not out or not os.path.exists(out):
            return
        try:
            if os.name == "nt":
                os.startfile(out)  # noqa: S606
            else:
                import subprocess
                opener = "open" if sys_is_mac() else "xdg-open"
                subprocess.Popen([opener, out])
        except Exception:  # noqa: BLE001 - opening is best-effort
            pass

    # ===================================================================== #
    # Cost estimate
    # ===================================================================== #
    def _on_estimate_cost(self) -> None:
        if not self._items:
            self.cost_label.config(text="Extract first to estimate.")
            return
        est = estimate_cost(self._items, self._gather_config())
        self.cost_label.config(
            text=f"~{est['input_tokens']:,} in / {est['output_tokens']:,} out tokens "
                 f"· {est['n_batches']} batches · ≈ ${est['usd']:.4f} "
                 f"({est.get('model') or 'default'})")

    # ===================================================================== #
    # Event pump
    # ===================================================================== #
    def _poll_events(self) -> None:
        try:
            while True:
                self._handle_event(self._events.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _handle_event(self, event: tuple) -> None:
        kind = event[0]
        if kind == "extract_done":
            self.extract_prog.stop()
            self.extract_btn.config(state="normal")
            self._after_extract(event[1])
        elif kind == "extract_qg_failed":
            self.extract_prog.stop()
            self.extract_btn.config(state="normal")
            result, failures = event[1], event[2]
            self._after_extract(result, quality_gate_failed=True, gate_failures=failures)
        elif kind == "extract_err":
            self.extract_prog.stop()
            self.extract_btn.config(state="normal")
            self.extract_status.config(text=f"Error: {event[1]}")
            messagebox.showerror("pdf2excel", f"Extraction failed:\n{event[1]}")
        elif kind == "enrich_progress":
            _, done, total, msg = event
            self.run_prog.config(value=done, maximum=max(total, 1))
            self.run_pct.config(text=f"{done}/{total}")
            if msg:
                self.run_status.config(text=msg)
        elif kind == "enrich_done":
            self._after_enrich(event[1])
        elif kind == "enrich_err":
            self.generate_btn.config(state="normal")
            self._log(f"ERROR: {event[1]}")
            self.run_status.config(text=f"Error: {event[1]}")
            messagebox.showerror("pdf2excel", f"AI enrichment failed:\n{event[1]}")
        elif kind == "models_done":
            models = event[1]
            self.model_cb["values"] = models
            if models and self.model_var.get() not in models:
                self.model_var.set(models[0])
            self.key_status.config(text=f"✓ {len(models)} models")
        elif kind == "models_err":
            self.key_status.config(text="model list failed")
            messagebox.showerror("pdf2excel", f"Could not list models:\n{event[1]}")

    # ===================================================================== #
    # Result rendering
    # ===================================================================== #
    def _populate_extract_preview(self, items, tree: ttk.Treeview) -> None:
        tree.delete(*tree.get_children())
        for it in items[:_EXTRACT_PREVIEW_LIMIT]:
            text = (it.get("text") or "")[:120]
            if len(it.get("text") or "") > 120:
                text += "…"
            issues = it.get("issues") or []
            issue_str = ", ".join(issues[:3])
            if len(issues) > 3:
                issue_str += "…"
            tags = ("warning",) if issues else ()
            tree.insert("", "end", values=(
                it.get("clause_id", ""),
                it.get("title", ""),
                it.get("classification", ""),
                text,
                f"{it.get('confidence', 1.0):.2f}",
                issue_str,
            ), tags=tags)

    def _format_extract_summary(self, result, gate_failures=None) -> str:
        meta = result.meta or {}
        profile = result.profile or meta.get("profile", "")
        lines = [
            "Extraction completed.",
            f"Output: {result.out_path}",
        ]
        review = result.review_output_path or meta.get("review_output")
        if review:
            lines.append(f"Review workbook: {review}")
        lines.extend([
            f"Mode used: {result.mode}",
            f"Format: {result.fmt}",
            f"Profile: {profile or 'n/a'}",
            f"Items exported: {result.n_items}",
            f"Requirements: {result.n_requirements or meta.get('n_requirements', 0)}",
            f"Information: {result.n_information or meta.get('n_information', 0)}",
            f"Issues: {result.issues_count or meta.get('issues_count', 0)}",
            f"Rejected blocks: {result.rejected_count or meta.get('rejected_count', 0)}",
            f"Quality gate passed: {result.quality_gate_passed}",
            f"Warnings: {len(result.warnings)}",
        ])
        if gate_failures:
            lines.append("Quality gate: FAILED")
            for f in gate_failures[:8]:
                lines.append(f"  • {f}")
        elif not result.quality_gate_passed:
            lines.append("Quality gate: FAILED (no rows exported)")
        else:
            lines.append("Quality gate: passed")
        if self.force_export_var.get() and result.n_items:
            lines.append("Note: Force export was enabled.")
        return "\n".join(lines)

    def _after_extract(self, result, quality_gate_failed: bool = False,
                       gate_failures=None) -> None:
        self._items = list(result.items)
        self._export_items = list(result.items)
        self._meta = dict(result.meta)
        self._enriched = [dict(it) for it in self._items]
        self._quality_gate_ok = result.quality_gate_passed and not quality_gate_failed

        self._populate_extract_preview(self._export_items, self.extract_tree)
        self._populate_extract_preview(self._export_items, self.run_preview_tree)

        # Preview sanity check for NIST-like output
        if self._export_items and result.fmt == "standard":
            first = self._export_items[0]
            if len(self._export_items) > 3000:
                self._log(f"WARNING: {len(self._export_items)} rows exported — likely wrong extractor.")
            elif first.get("clause_id") != "AC-1" and "800-53" in (self.pdf_var.get() or ""):
                self._log(
                    f"WARNING: first row is {first.get('clause_id')!r}, expected 'AC-1' for NIST PDF."
                )

        summary = self._format_extract_summary(result, gate_failures)
        self._log_clear()
        self._log(summary)

        if getattr(result, "has_scanned_pages", False):
            messagebox.showwarning(
                "Scanned Pages Detected",
                "Scanned or image-only pages were detected. "
                "Enable OCR mode 'Detect only' and re-run if needed.",
            )

        if quality_gate_failed or not result.quality_gate_passed:
            review = result.review_output_path or (result.meta or {}).get("review_output")
            extra = f"\n\nReview workbook:\n{review}" if review else ""
            messagebox.showerror(
                "Quality Gate Failed",
                "Quality gate failed. Bad rows were not exported.\n"
                "Check the review workbook and Extraction_Issues sheet."
                + extra,
            )
            self.generate_btn.config(state="disabled")
        elif self.force_export_var.get() and result.n_items:
            messagebox.showwarning(
                "Force Export",
                "Force export enabled. Output may contain low-quality rows.",
            )

        if result.fmt == "standard" and self._quality_gate_ok:
            self.generate_btn.config(state="normal")
            self.export_btn.config(state="normal")
        elif result.fmt == "standard":
            self.export_btn.config(state="normal")

        short = (
            f"{'Quality gate failed — ' if not self._quality_gate_ok else ''}"
            f"{result.n_items} rows exported · "
            f"{result.issues_count} issue rows · "
            f"{len(result.warnings)} warnings"
        )
        self.extract_status.config(text=short)
        self.run_status.config(
            text=f"{'Ready for AI fill.' if self._quality_gate_ok else 'Fix extraction issues before AI fill.'} "
                 f"See log for details.")
        if result.fmt == "standard":
            self._populate_results(self._enriched)
        else:
            self.generate_btn.config(state="disabled")
        self.open_btn.config(state="normal")

    def _after_enrich(self, enriched_sub) -> None:
        # Merge the enriched subset back into the full working copy, then keep
        # the REQ-NNN sequence continuous across all rows enriched so far.
        if len(self._enriched) != len(self._items):
            self._enriched = [dict(it) for it in self._items]
        targets = getattr(self, "_enrich_targets", list(range(len(enriched_sub))))
        for pos, i in enumerate(targets):
            if pos < len(enriched_sub) and 0 <= i < len(self._enriched):
                self._enriched[i] = enriched_sub[pos]
        renumber_requirements(self._enriched)

        self.generate_btn.config(state="normal")
        self.export_btn.config(state="normal")
        n_done = len(enriched_sub)
        run_req = sum(1 for it in enriched_sub
                      if it.get("classification") == "Requirement")
        run_info = n_done - run_req
        self.run_pct.config(text="done")
        note = " (Information rows leave F–I blank by design.)" if run_info else ""
        self._log(f"Done — enriched {n_done} row(s) this run: {run_req} requirement(s), "
                  f"{run_info} information.{note} Review/edit below, then Export.")
        self.run_status.config(
            text=f"AI fill: {n_done} row(s) · {run_req} requirement(s), "
                 f"{run_info} information.{note}")
        self._populate_results(self._enriched)
        # Highlight + scroll to the rows just enriched so they're easy to spot.
        kids = self.result_tree.get_children()
        sel = [kids[i] for i in targets if 0 <= i < len(kids)]
        if sel:
            self.result_tree.selection_set(sel)
            self.result_tree.see(sel[0])

    def _populate_results(self, items) -> None:
        self.result_tree.delete(*self.result_tree.get_children())
        for it in items[:_PREVIEW_LIMIT]:
            vals = [it.get("clause_id", "")]
            vals += [it.get(field, "") for field in _RESULT_FIELDS.values()]
            
            tags = []
            if it.get("issues"):
                tags.append("warning")
            elif it.get("confidence", 1.0) < 0.5:
                tags.append("low_conf")
                
            self.result_tree.insert("", "end", values=vals, tags=tags)

    def _on_result_edit(self, event) -> None:
        """Double-click an E–I cell → inline edit; commit back to the item."""
        if not self._enriched:
            return
        region = self.result_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.result_tree.identify_row(event.y)
        col_id = self.result_tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_idx = int(col_id[1:]) - 1          # "#1" -> 0 (clause, read-only)
        if col_idx == 0:
            return
        row_idx = self.result_tree.index(row_id)
        if row_idx >= len(self._enriched):
            return
        field = list(_RESULT_FIELDS.values())[col_idx - 1]

        x, y, w, h = self.result_tree.bbox(row_id, col_id)
        var = tk.StringVar(value=str(self._enriched[row_idx].get(field, "")))
        entry = ttk.Entry(self.result_tree, textvariable=var)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()

        def commit(_e=None):
            self._enriched[row_idx][field] = var.get()
            self.result_tree.set(row_id, col_id, var.get())
            entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", lambda _e: entry.destroy())

    # ===================================================================== #
    # Logging
    # ===================================================================== #
    def _log(self, msg: str) -> None:
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _log_clear(self) -> None:
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")


def sys_is_mac() -> bool:
    import sys
    return sys.platform == "darwin"


def main() -> int:
    App().mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
