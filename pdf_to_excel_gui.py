"""A small Tkinter desktop GUI for pdf2excel.

Pick a PDF, choose an output location, a **Mode** (Auto / Prose / Tables) and
the paragraph-break sensitivity, then click Extract. The actual work is done by
:func:`router.convert`, so this file stays a thin shell with no extraction
logic of its own. Conversion runs on a background thread to keep the window
responsive, and the result is previewed in a table:

  * prose  -> para_id | page | type | section | text
  * tables -> a summary of the sheets written (sheet | rows | cols)

Run:
    python pdf_to_excel_gui.py
"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from router import convert
from web_extract import clean_url

# How many prose rows to show in the preview (the full set is still written).
_PREVIEW_LIMIT = 300

_MODE_LABELS = {"Auto": "auto", "Prose": "prose", "Tables": "tables"}
_FORMAT_LABELS = {"Default": "default", "Standard Assessment": "standard"}
_RENDER_LABELS = {"Auto": "auto", "Always": "always", "Never": "never"}


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("pdf2excel — PDF to Excel")
        self.geometry("820x520")
        self.minsize(700, 440)

        self._events: "queue.Queue[tuple]" = queue.Queue()
        self._build_widgets()
        self.after(100, self._poll_events)

    # -- UI ---------------------------------------------------------------- #
    def _build_widgets(self) -> None:
        pad = {"padx": 8, "pady": 5}
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(8, weight=1)

        # Input: a local PDF path OR a URL (PDF or HTML page)
        ttk.Label(frm, text="PDF file or URL:").grid(row=0, column=0, sticky="w", **pad)
        self.pdf_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.pdf_var).grid(
            row=0, column=1, sticky="ew", **pad
        )
        ttk.Button(frm, text="Browse…", command=self._browse_pdf).grid(
            row=0, column=2, **pad
        )

        # Output xlsx
        ttk.Label(frm, text="Output Excel:").grid(row=1, column=0, sticky="w", **pad)
        self.out_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.out_var).grid(
            row=1, column=1, sticky="ew", **pad
        )
        ttk.Button(frm, text="Save as…", command=self._browse_out).grid(
            row=1, column=2, **pad
        )

        # Mode + gap factor on one row
        opts = ttk.Frame(frm)
        opts.grid(row=2, column=0, columnspan=3, sticky="w", **pad)

        ttk.Label(opts, text="Mode:").pack(side="left")
        self.mode_var = tk.StringVar(value="Auto")
        ttk.Combobox(
            opts,
            textvariable=self.mode_var,
            values=list(_MODE_LABELS.keys()),
            state="readonly",
            width=10,
        ).pack(side="left", padx=(4, 20))

        ttk.Label(opts, text="Gap factor:").pack(side="left")
        self.gap_var = tk.DoubleVar(value=1.6)
        ttk.Scale(
            opts,
            from_=1.1,
            to=3.0,
            variable=self.gap_var,
            orient="horizontal",
            length=200,
            command=lambda _=None: self._gap_label.config(
                text=f"{self.gap_var.get():.2f}"
            ),
        ).pack(side="left", padx=4)
        self._gap_label = ttk.Label(opts, text="1.60")
        self._gap_label.pack(side="left")
        ttk.Label(opts, text="(prose only)", foreground="#888").pack(
            side="left", padx=6
        )

        # Output format + Standard ID on the next row
        opts2 = ttk.Frame(frm)
        opts2.grid(row=3, column=0, columnspan=3, sticky="w", **pad)

        ttk.Label(opts2, text="Output format:").pack(side="left")
        self.fmt_var = tk.StringVar(value="Default")
        ttk.Combobox(
            opts2,
            textvariable=self.fmt_var,
            values=list(_FORMAT_LABELS.keys()),
            state="readonly",
            width=20,
        ).pack(side="left", padx=(4, 20))

        ttk.Label(opts2, text="Standard ID:").pack(side="left")
        self.std_id_var = tk.StringVar(value="MLSR")
        ttk.Entry(opts2, textvariable=self.std_id_var, width=16).pack(
            side="left", padx=4
        )
        ttk.Label(opts2, text="(Standard Assessment only)", foreground="#888").pack(
            side="left", padx=6
        )

        # URL-fetch options: JS render mode + insecure-TLS toggle.
        opts3 = ttk.Frame(frm)
        opts3.grid(row=4, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(opts3, text="Render JavaScript:").pack(side="left")
        self.render_var = tk.StringVar(value="Auto")
        ttk.Combobox(
            opts3,
            textvariable=self.render_var,
            values=list(_RENDER_LABELS.keys()),
            state="readonly",
            width=8,
        ).pack(side="left", padx=(4, 20))
        self.insecure_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opts3,
            text="Allow insecure TLS (skip certificate verification)",
            variable=self.insecure_var,
        ).pack(side="left")
        ttk.Label(
            opts3,
            text="— disables MITM protection; for trusted public sources only",
            foreground="#a00",
        ).pack(side="left", padx=6)

        # Action button
        self.run_btn = ttk.Button(frm, text="Extract", command=self._on_extract)
        self.run_btn.grid(row=5, column=0, sticky="w", **pad)

        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew", **pad)

        self.status = ttk.Label(frm, text="Ready.", foreground="#444")
        self.status.grid(row=7, column=0, columnspan=3, sticky="w", **pad)

        # Preview table (columns reconfigured per result mode)
        prev = ttk.Frame(frm)
        prev.grid(row=8, column=0, columnspan=3, sticky="nsew", **pad)
        prev.rowconfigure(0, weight=1)
        prev.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(prev, show="headings", height=10)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(prev, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=yscroll.set)

    # -- File pickers ------------------------------------------------------ #
    def _browse_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose a PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self.pdf_var.set(path)
            if not self.out_var.get():
                self.out_var.set(os.path.splitext(path)[0] + ".xlsx")

    def _browse_out(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Excel as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.out_var.set(path)

    # -- Extraction (threaded) -------------------------------------------- #
    def _on_extract(self) -> None:
        pdf = self.pdf_var.get().strip()
        out = self.out_var.get().strip()
        is_url = "://" in pdf
        if is_url:
            # Normalize a pasted URL (strip wrappers / trailing " [host]") up front.
            try:
                pdf = clean_url(pdf)
                self.pdf_var.set(pdf)
            except ValueError as exc:
                messagebox.showerror("pdf2excel", str(exc))
                return
        elif not pdf or not os.path.isfile(pdf):
            messagebox.showerror(
                "pdf2excel", "Please choose a valid PDF file or enter a URL."
            )
            return
        if not out:
            # URLs rarely have a useful basename; fall back to a fixed name.
            base = os.path.splitext(os.path.basename(pdf))[0] if not is_url else ""
            out = (base or "web_export") + ".xlsx"
            self.out_var.set(out)

        mode = _MODE_LABELS.get(self.mode_var.get(), "auto")
        fmt = _FORMAT_LABELS.get(self.fmt_var.get(), "default")
        std_id = self.std_id_var.get().strip() or "MLSR"
        insecure = bool(self.insecure_var.get())
        render = _RENDER_LABELS.get(self.render_var.get(), "auto")
        self.run_btn.config(state="disabled")
        self.progress.start(12)
        self.status.config(text="Working…")

        worker = threading.Thread(
            target=self._worker,
            args=(pdf, out, mode, float(self.gap_var.get()), fmt, std_id,
                  insecure, render),
            daemon=True,
        )
        worker.start()

    def _worker(self, pdf, out, mode, gap, fmt, std_id, insecure, render) -> None:
        try:
            result = convert(
                pdf, out, mode=mode, gap_factor=gap,
                fmt=fmt, standard_id=std_id, insecure=insecure, render=render,
            )
            self._events.put(("done", result))
        except Exception as exc:  # noqa: BLE001
            self._events.put(("error", str(exc)))

    def _poll_events(self) -> None:
        try:
            while True:
                self._handle_event(self._events.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _handle_event(self, event: tuple) -> None:
        self.progress.stop()
        self.run_btn.config(state="normal")
        kind = event[0]
        if kind == "error":
            msg = event[1]
            self.status.config(text=f"Error: {msg}")
            messagebox.showerror("pdf2excel", f"Conversion failed:\n{msg}")
            return

        result = event[1]
        if result.mode == "prose":
            self._show_prose(result)
        else:
            self._show_tables(result)

    # -- Preview rendering ------------------------------------------------- #
    def _reset_tree(self, columns, widths) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns
        for col, width in zip(columns, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor="w", stretch=(col == columns[-1]))

    def _show_prose(self, result) -> None:
        self._reset_tree(
            ("para_id", "page", "type", "section", "text"),
            (60, 50, 70, 200, 520),
        )
        for p in result.paragraphs[:_PREVIEW_LIMIT]:
            self.tree.insert(
                "", "end",
                values=(p.para_id, p.page, p.type, p.section, p.text),
            )
        total = len(result.paragraphs)
        shown = min(total, _PREVIEW_LIMIT)
        more = f" (showing first {shown})" if total > shown else ""
        tag = self._fmt_tag(result)
        self.status.config(
            text=f"[prose{tag}] {total} paragraphs{more} → {result.out_path}"
        )
        messagebox.showinfo(
            "pdf2excel",
            f"Prose mode{tag}: {total} paragraphs"
            f"{self._rows_note(result)}.\n\nSaved to:\n{result.out_path}",
        )

    def _show_tables(self, result) -> None:
        self._reset_tree(("sheet", "rows", "cols"), (260, 80, 80))
        for name, rows, cols in result.sheets:
            self.tree.insert("", "end", values=(name, rows, cols))
        n_tables = max(len(result.sheets) - 1, 0)  # minus slides_text
        tag = self._fmt_tag(result)
        self.status.config(
            text=f"[tables{tag}] {n_tables} table sheets + slides_text "
            f"→ {result.out_path}"
        )
        messagebox.showinfo(
            "pdf2excel",
            f"Tables mode{tag}: {n_tables} table sheet(s) + slides_text"
            f"{self._rows_note(result)}.\n\nSaved to:\n{result.out_path}",
        )

    @staticmethod
    def _fmt_tag(result) -> str:
        return " → Standard Assessment" if result.fmt == "standard" else ""

    @staticmethod
    def _rows_note(result) -> str:
        return f"; {result.n_items} rows written" if result.fmt == "standard" else ""


def main() -> int:
    App().mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
