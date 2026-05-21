"""
icms_gui.py
~~~~~~~~~~~
Tkinter GUI wrapper for the ICMS weekly snapshot ETL pipeline.

Calls icms_etl_core.py functions directly.  Does not invoke
update_icms_snapshot_files.py as a subprocess.

Usage
-----
    python icms_gui.py
"""

from __future__ import annotations

import datetime as dt
import os
import queue
import re
import threading
import traceback
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import icms_etl_core as core
from update_icms_snapshot_files import (
    _parse_date_from_filename,
    load_settings,
    save_settings,
)

# ── Optional dependencies ──────────────────────────────────────────────────────

try:
    from tkcalendar import DateEntry  # type: ignore
    _HAS_DATEPICKER = True
except ImportError:
    DateEntry = None  # type: ignore
    _HAS_DATEPICKER = False

try:
    import keyring as _keyring  # type: ignore
    _HAS_KEYRING = True
except ImportError:
    _keyring = None  # type: ignore
    _HAS_KEYRING = False

_KR_SERVICE = "ICMS_ETL"
_KR_USER = "dhs_snapshot"

# ── Design tokens ─────────────────────────────────────────────────────────────

BG           = "#F5F3EF"   # warm off-white
INPUT_BG     = "#FFFFFF"
ACCENT       = "#4A7C8E"   # muted teal
ACCENT_DARK  = "#3A6474"
FG           = "#2C2C2C"
FG_MUTED     = "#6B6B6B"
RED          = "#C0392B"
AMBER        = "#B07D2A"
GREEN        = "#2E7D32"
BORDER       = "#C8C5BF"
LOG_BG       = "#FAFAF8"

_F = "Segoe UI"
F_BODY   = (_F, 13)
F_LABEL  = (_F, 11)
F_HEADER = (_F, 15)
F_SMALL  = (_F, 10)
F_MONO   = ("Consolas", 11)

PAD   = 20   # outer panel padding
VSEP  = 12   # vertical gap between field rows

_TOTAL_STEPS = 7
_STEP_LABELS = [
    "Opening workbook…",
    "Reading sheets…",
    "Resolving fields…",
    "Writing flat file…",
    "Writing snapshot history…",
    "Writing client/project master…",
    "Writing logs…",
]


# ── GUI-compatible field resolver (called from worker thread) ─────────────────

def _gui_resolve_fields(
    df_columns: List[str],
    aliases: Dict[str, str],
    script_dir: Path,
    gui_q: "queue.Queue[dict]",
    resp_q: "queue.Queue[Optional[str]]",
) -> tuple:
    """Resolve fields using GUI dialogs for required unresolved fields.

    Optional unresolved fields are silently skipped.
    Communicates with the main thread via *gui_q* / *resp_q*.
    Returns (resolved, missing_req, missing_opt, aliases_added).
    """
    resolved, unresolved_req, unresolved_opt = core.resolve_fields(
        df_columns, aliases
    )
    aliases_added = 0
    missing_req: List[str] = []
    missing_opt: List[str] = list(unresolved_opt)

    for field in unresolved_req:
        gui_q.put({
            "type": "alias_needed",
            "canonical": field,
            "columns": df_columns,
        })
        chosen: Optional[str] = resp_q.get()   # blocks until main thread responds
        if chosen is not None:
            resolved[field] = chosen
            if aliases.get(field) != chosen:
                aliases[field] = chosen
                aliases_added += 1
        else:
            missing_req.append(field)

    if aliases_added > 0:
        core.save_aliases(aliases, script_dir)

    return resolved, missing_req, missing_opt, aliases_added


# ── Field-alias modal dialog ──────────────────────────────────────────────────

class _FieldAliasDialog(tk.Toplevel):
    """Modal dialog that lets the user map an unresolved required field."""

    def __init__(
        self,
        parent: tk.Tk,
        canonical: str,
        columns: List[str],
        resp_q: "queue.Queue[Optional[str]]",
    ):
        super().__init__(parent)
        self._canonical = canonical
        self._columns = columns
        self._resp_q = resp_q

        self.title("Map Required Field")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_skip)

        self._build()

        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_reqwidth()
        dh = self.winfo_reqheight()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _build(self):
        outer = ttk.Frame(self, padding=PAD)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Required field not found in workbook:",
                  foreground=FG_MUTED, font=F_LABEL,
                  background=BG).pack(anchor="w")
        ttk.Label(outer, text=self._canonical,
                  foreground=FG, font=F_HEADER,
                  background=BG).pack(anchor="w", pady=(2, 8))

        ttk.Label(
            outer,
            text=(
                "Select a replacement column from the list, or type a name\n"
                "in the box below, then click  Use this column."
            ),
            foreground=FG_MUTED, font=F_LABEL,
            background=BG, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        # Listbox
        lb_frame = ttk.Frame(outer)
        lb_frame.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(lb_frame, orient="vertical")
        self._lb = tk.Listbox(
            lb_frame,
            yscrollcommand=sb.set,
            font=F_LABEL,
            selectmode="single",
            height=9,
            bg=INPUT_BG,
            fg=FG,
            selectbackground=ACCENT,
            selectforeground="white",
            borderwidth=1,
            relief="flat",
            activestyle="none",
        )
        sb.config(command=self._lb.yview)
        sb.pack(side="right", fill="y")
        self._lb.pack(side="left", fill="both", expand=True)
        for col in self._columns:
            self._lb.insert(tk.END, col)
        self._lb.bind("<<ListboxSelect>>", self._on_lb_select)

        # Entry
        ttk.Label(outer, text="Column name to use:",
                  foreground=FG_MUTED, font=F_LABEL,
                  background=BG).pack(anchor="w", pady=(10, 2))
        self._entry_var = tk.StringVar()
        self._entry = ttk.Entry(outer, textvariable=self._entry_var, font=F_BODY)
        self._entry.pack(fill="x")

        # Buttons
        btn_row = ttk.Frame(outer)
        btn_row.pack(fill="x", pady=(14, 0))
        ttk.Button(
            btn_row,
            text="Skip this field",
            command=self._on_skip,
        ).pack(side="left")
        ttk.Button(
            btn_row,
            text="Use this column",
            command=self._on_use,
        ).pack(side="right")

    def _on_lb_select(self, _=None):
        sel = self._lb.curselection()
        if sel:
            self._entry_var.set(self._columns[sel[0]])

    def _on_use(self):
        chosen = self._entry_var.get().strip()
        if not chosen:
            messagebox.showwarning(
                "No column selected",
                "Please select or type a column name.",
                parent=self,
            )
            return
        if chosen not in self._columns:
            messagebox.showwarning(
                "Unrecognised column",
                f"'{chosen}' is not a column in this workbook.\n\n"
                "Please select from the list.",
                parent=self,
            )
            return
        self._resp_q.put(chosen)
        self.destroy()

    def _on_skip(self):
        confirmed = messagebox.askokcancel(
            "Skip Required Field",
            f"'{self._canonical}' is a REQUIRED field.\n\n"
            "Skipping it means this field will be absent from all outputs.\n\n"
            "Continue without this field?",
            parent=self,
        )
        if confirmed:
            self._resp_q.put(None)
            self.destroy()


class _BatchConfirmDialog(tk.Toplevel):
    """Modal dialog to confirm/correct snapshot dates for batch files."""

    def __init__(self, parent: tk.Tk, entries: List[dict]):
        super().__init__(parent)
        self.title("Confirm Batch Files")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self._entries = entries
        self.result: Optional[List[Tuple[Path, str]]] = None
        self._selected_idx: Optional[int] = None

        self._build()
        self._refresh_list()
        self._center(parent)

    def _build(self):
        outer = ttk.Frame(self, padding=PAD)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="Review snapshot dates before batch ingest.",
            font=F_LABEL,
            foreground=FG_MUTED,
            background=BG,
        ).pack(anchor="w", pady=(0, 8))

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True)
        self._list = tk.Listbox(
            list_frame,
            font=F_LABEL,
            height=10,
            selectmode="single",
            bg=INPUT_BG,
            fg=FG,
            selectbackground=ACCENT,
            selectforeground="white",
            relief="flat",
            borderwidth=1,
        )
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self._list.yview)
        self._list.configure(yscrollcommand=sb.set)
        self._list.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._list.bind("<<ListboxSelect>>", self._on_select)

        row = ttk.Frame(outer)
        row.pack(fill="x", pady=(10, 0))
        ttk.Label(
            row,
            text="Corrected date (YYYY-MM-DD):",
            font=F_LABEL,
            foreground=FG_MUTED,
            background=BG,
        ).pack(anchor="w")
        self._date_var = tk.StringVar()
        ttk.Entry(row, textvariable=self._date_var, font=F_BODY).pack(fill="x", pady=(2, 0))

        btn_row = ttk.Frame(outer)
        btn_row.pack(fill="x", pady=(12, 0))
        ttk.Button(btn_row, text="Apply Date", command=self._apply_date).pack(side="left")
        ttk.Button(btn_row, text="Cancel", command=self._on_cancel).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="Confirm Batch", style="Accent.TButton", command=self._on_confirm).pack(side="right")

    def _center(self, parent: tk.Tk):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_reqwidth()
        dh = self.winfo_reqheight()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _sort_entries(self):
        self._entries.sort(key=lambda e: (e["date"] is None, e["date"] or ""))

    def _refresh_list(self):
        self._sort_entries()
        self._list.delete(0, tk.END)
        for i, entry in enumerate(self._entries, 1):
            label = entry["date"] if entry["date"] else "UNPARSED"
            self._list.insert(tk.END, f"{i:2}. {label:<12}  {entry['path'].name}")

    def _on_select(self, _event=None):
        sel = self._list.curselection()
        if not sel:
            self._selected_idx = None
            return
        self._selected_idx = sel[0]
        current = self._entries[self._selected_idx]["date"] or ""
        self._date_var.set(current)

    def _apply_date(self):
        if self._selected_idx is None:
            messagebox.showwarning("No file selected", "Select a file to edit first.", parent=self)
            return
        raw = self._date_var.get().strip()
        try:
            dt.datetime.strptime(raw, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid date", "Use YYYY-MM-DD format.", parent=self)
            return
        self._entries[self._selected_idx]["date"] = raw
        self._refresh_list()

    def _on_confirm(self):
        if any(e["date"] is None for e in self._entries):
            messagebox.showwarning(
                "Missing dates",
                "Please correct all UNPARSED dates before confirming.",
                parent=self,
            )
            return
        self.result = [(e["path"], e["date"]) for e in self._entries]
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


class _BatchDuplicateDialog(tk.Toplevel):
    """Modal dialog to decide how to handle already-integrated files."""

    def __init__(self, parent: tk.Tk, duplicates: List[Tuple[Path, str]]):
        super().__init__(parent)
        self.title("Already Integrated Files Detected")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self._duplicates = duplicates
        self.action: str = "cancel"  # exclude | include | cancel
        self._build()
        self._center(parent)

    def _build(self):
        outer = ttk.Frame(self, padding=PAD)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text=(
                "Some selected files appear to be already integrated\n"
                "(same source_file_name + snapshot_date)."
            ),
            font=F_LABEL,
            foreground=FG_MUTED,
            background=BG,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True)
        lb = tk.Listbox(
            list_frame,
            font=F_LABEL,
            height=min(10, max(3, len(self._duplicates))),
            bg=INPUT_BG,
            fg=FG,
            relief="flat",
            borderwidth=1,
        )
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for p, s in self._duplicates:
            lb.insert(tk.END, f"{s}  |  {p.name}")

        btn = ttk.Frame(outer)
        btn.pack(fill="x", pady=(12, 0))
        ttk.Button(
            btn,
            text="Exclude flagged files (recommended)",
            command=self._on_exclude,
        ).pack(side="left")
        ttk.Button(
            btn,
            text="Include all anyway",
            command=self._on_include,
        ).pack(side="right")
        ttk.Button(
            outer,
            text="Cancel",
            command=self._on_cancel,
        ).pack(anchor="e", pady=(8, 0))

    def _center(self, parent: tk.Tk):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_reqwidth()
        dh = self.winfo_reqheight()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _on_exclude(self):
        self.action = "exclude"
        self.destroy()

    def _on_include(self):
        self.action = "include"
        self.destroy()

    def _on_cancel(self):
        self.action = "cancel"
        self.destroy()


class _BatchServicesDuplicateDialog(tk.Toplevel):
    """Modal dialog for files whose Services Log appears already integrated."""

    def __init__(self, parent: tk.Tk, duplicates: List[Tuple[Path, str]]):
        super().__init__(parent)
        self.title("Services Log Appears Already Integrated")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self._duplicates = duplicates
        self.action: str = "cancel"  # exclude | include | cancel
        self._build()
        self._center(parent)

    def _build(self):
        outer = ttk.Frame(self, padding=PAD)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text=(
                "Some selected files have a Services Log table that appears\n"
                "already integrated in this output folder.\n\n"
                "This warning is table-level (whole Services Log overlap),\n"
                "not row-level duplicates."
            ),
            font=F_LABEL,
            foreground=FG_MUTED,
            background=BG,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True)
        lb = tk.Listbox(
            list_frame,
            font=F_LABEL,
            height=min(10, max(3, len(self._duplicates))),
            bg=INPUT_BG,
            fg=FG,
            relief="flat",
            borderwidth=1,
        )
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for p, s in self._duplicates:
            lb.insert(tk.END, f"{s}  |  {p.name}")

        btn = ttk.Frame(outer)
        btn.pack(fill="x", pady=(12, 0))
        ttk.Button(
            btn,
            text="Exclude flagged files (recommended)",
            command=self._on_exclude,
        ).pack(side="left")
        ttk.Button(
            btn,
            text="Include all anyway",
            command=self._on_include,
        ).pack(side="right")
        ttk.Button(
            outer,
            text="Cancel",
            command=self._on_cancel,
        ).pack(anchor="e", pady=(8, 0))

    def _center(self, parent: tk.Tk):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_reqwidth()
        dh = self.winfo_reqheight()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _on_exclude(self):
        self.action = "exclude"
        self.destroy()

    def _on_include(self):
        self.action = "include"
        self.destroy()

    def _on_cancel(self):
        self.action = "cancel"
        self.destroy()


# ── Main application ──────────────────────────────────────────────────────────

class ICMSApp(tk.Tk):
    """Single-window Tkinter application for the ICMS snapshot ETL pipeline."""

    def __init__(self):
        super().__init__()

        self.title("ICMS Snapshot ETL")
        self.geometry("560x820")
        self.minsize(520, 640)
        self.resizable(False, True)
        self.configure(bg=BG)

        self._script_dir: Path = Path(__file__).parent.resolve()
        self._gui_q: queue.Queue = queue.Queue()
        self._resp_q: queue.Queue = queue.Queue()
        self._etl_results: dict = {}
        self._current_output_dir: Optional[Path] = None
        self._batch_files: List[Path] = []
        self._batch_date_rows: List[dict] = []
        self._last_skipped_duplicates: List[Tuple[Path, str]] = []
        self._last_skipped_services_duplicates: List[Tuple[Path, str]] = []

        self._configure_styles()

        # Build all three panels and stack them
        self._setup_frame  = ttk.Frame(self)
        self._run_frame    = ttk.Frame(self)
        self._results_frame = ttk.Frame(self)
        for frame in (self._setup_frame, self._run_frame, self._results_frame):
            frame.place(relwidth=1, relheight=1)

        self._build_setup_panel()
        self._build_run_panel()
        self._build_results_panel()

        self._load_initial_state()
        self._show_panel(self._setup_frame)
        self.after(100, self._poll_queue)

    # ── Styles ────────────────────────────────────────────────────────────────

    def _configure_styles(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass

        s.configure(".",              background=BG, foreground=FG, font=F_BODY)
        s.configure("TFrame",         background=BG)
        s.configure("TLabel",         background=BG, foreground=FG, font=F_BODY)
        s.configure("TSeparator",     background=BORDER)

        s.configure("TEntry",
                    fieldbackground=INPUT_BG, foreground=FG,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    insertcolor=FG)
        s.map("TEntry",
              bordercolor=[("focus", ACCENT)],
              lightcolor=[("focus", ACCENT)])

        s.configure("TButton",
                    background=BG, foreground=FG, font=F_BODY,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    padding=(8, 5))
        s.map("TButton",
              background=[("active", "#ECEAE5"), ("pressed", "#E0DDD8")])

        s.configure("Accent.TButton",
                    background=ACCENT, foreground="white",
                    bordercolor=ACCENT_DARK, lightcolor=ACCENT, darkcolor=ACCENT_DARK,
                    font=F_BODY, padding=(14, 9))
        s.map("Accent.TButton",
              background=[("disabled", "#A8BEC4"),
                          ("active",   ACCENT_DARK),
                          ("pressed",  "#2E5A68")],
              foreground=[("disabled", "#E0EAED")])

        s.configure("TCheckbutton", background=BG, foreground=FG, font=F_BODY)
        s.map("TCheckbutton", background=[("active", BG)])

        s.configure("Accent.Horizontal.TProgressbar",
                    troughcolor="#DAD7D1", background=ACCENT, thickness=10)

    # ── Panel builders ─────────────────────────────────────────────────────────

    def _build_setup_panel(self):
        f = self._setup_frame

        footer = ttk.Frame(f)
        footer.pack(side="bottom", fill="x")

        # Scrollable container so controls are always reachable on smaller displays.
        setup_container = ttk.Frame(f)
        setup_container.pack(fill="both", expand=True, side="top")

        self._setup_canvas = tk.Canvas(
            setup_container,
            bg=BG,
            highlightthickness=0,
            borderwidth=0,
            relief="flat",
        )
        setup_scroll = ttk.Scrollbar(
            setup_container, orient="vertical", command=self._setup_canvas.yview
        )
        self._setup_canvas.configure(yscrollcommand=setup_scroll.set)

        setup_scroll.pack(side="right", fill="y")
        self._setup_canvas.pack(side="left", fill="both", expand=True)

        setup_content = ttk.Frame(self._setup_canvas)
        self._setup_canvas_window = self._setup_canvas.create_window(
            (0, 0), window=setup_content, anchor="nw"
        )

        setup_content.bind(
            "<Configure>",
            lambda _e: self._setup_canvas.configure(
                scrollregion=self._setup_canvas.bbox("all")
            ),
        )
        self._setup_canvas.bind("<Configure>", self._on_setup_canvas_configure)
        self._setup_canvas.bind("<Enter>", self._bind_setup_mousewheel)
        self._setup_canvas.bind("<Leave>", self._unbind_setup_mousewheel)

        ttk.Label(setup_content, text="ICMS Snapshot ETL", font=F_HEADER).pack(
            anchor="w", padx=PAD, pady=(PAD, 0))
        ttk.Separator(setup_content, orient="horizontal").pack(
            fill="x", padx=PAD, pady=(10, 16))

        body = ttk.Frame(setup_content)
        body.pack(fill="both", expand=True, padx=PAD)

        # ── Input file ──────────────────────────────────────────────────────
        self._mk_label(body, "Input File (.xlsx)")
        fr = ttk.Frame(body)
        fr.pack(fill="x", pady=(3, 0))

        self._file_var = tk.StringVar()
        self._file_var.trace_add("write", self._on_field_changed)
        ttk.Entry(fr, textvariable=self._file_var, state="readonly",
                  font=F_BODY).pack(side="left", fill="x", expand=True)
        ttk.Button(fr, text="Browse…",
                   command=self._on_browse_file).pack(side="left", padx=(6, 0))
        ttk.Button(fr, text="Browse files…",
                   command=self._on_browse_files).pack(side="left", padx=(6, 0))

        self._file_err = self._mk_err(body)
        self._batch_info_lbl = ttk.Label(
            body, text="", foreground=FG_MUTED, font=F_SMALL, background=BG
        )
        self._batch_info_lbl.pack(anchor="w", pady=(2, 0))

        # ── Snapshot date ───────────────────────────────────────────────────
        self._date_lbl = ttk.Label(
            body,
            text="Snapshot Date",
            foreground=FG_MUTED,
            font=F_LABEL,
            background=BG,
        )
        self._date_lbl.pack(anchor="w", pady=(VSEP, 0))
        self._date_row = ttk.Frame(body)
        self._date_row.pack(anchor="w", pady=(3, 0))

        if _HAS_DATEPICKER:
            self._date_entry = DateEntry(
                self._date_row,
                font=F_BODY,
                date_pattern="yyyy-mm-dd",
                background=ACCENT,
                foreground="white",
                headersbackground=ACCENT_DARK,
                headersforeground="white",
                normalbackground=INPUT_BG,
                normalforeground=FG,
                weekendbackground=INPUT_BG,
                weekendforeground=FG,
                othermonthforeground=FG_MUTED,
                width=13,
            )
            self._date_entry.pack(side="left")
            self._date_entry.bind("<<DateEntrySelected>>", self._on_field_changed)
            self._date_entry.bind("<KeyRelease>", self._on_field_changed)
        else:
            self._date_var = tk.StringVar()
            self._date_var.trace_add("write", self._on_field_changed)
            self._date_entry = ttk.Entry(
                self._date_row, textvariable=self._date_var, font=F_BODY, width=15)
            self._date_entry.pack(side="left")
            self._date_entry.insert(0, "YYYY-MM-DD")
            self._date_entry.configure(foreground=FG_MUTED)
            self._date_entry.bind("<FocusIn>",  self._on_date_focus_in)
            self._date_entry.bind("<FocusOut>", self._on_date_focus_out)

        self._batch_date_frame = ttk.Frame(body)
        self._batch_date_hdr = ttk.Label(
            self._batch_date_frame,
            text="Batch Snapshot Dates",
            foreground=FG_MUTED,
            font=F_LABEL,
            background=BG,
        )
        self._batch_date_hdr.pack(anchor="w", pady=(10, 4))
        self._batch_date_list = ttk.Frame(self._batch_date_frame)
        self._batch_date_list.pack(fill="x")

        self._date_err = self._mk_err(body)

        # ── Sheet names ─────────────────────────────────────────────────────
        self._mk_label(body, "Worksheets Queried", top_pad=VSEP)

        self._mk_label(body, "Main participant sheet", top_pad=6)
        self._sheet_var = tk.StringVar(value="Participant Level Summary")
        self._sheet_var.trace_add("write", self._on_field_changed)
        ttk.Entry(body, textvariable=self._sheet_var,
                  font=F_BODY).pack(fill="x", pady=(3, 0))

        self._mk_label(body, "Vacant slots sheet", top_pad=6)
        self._vacant_sheet_var = tk.StringVar(value="Vacant Slots")
        self._vacant_sheet_var.trace_add("write", self._on_field_changed)
        ttk.Entry(body, textvariable=self._vacant_sheet_var,
                  font=F_BODY).pack(fill="x", pady=(3, 0))

        self._mk_label(body, "Exit log sheet", top_pad=6)
        self._exit_sheet_var = tk.StringVar(value="Exit Log")
        self._exit_sheet_var.trace_add("write", self._on_field_changed)
        ttk.Entry(body, textvariable=self._exit_sheet_var,
                  font=F_BODY).pack(fill="x", pady=(3, 0))

        self._mk_label(body, "Services log sheet", top_pad=6)
        self._services_sheet_var = tk.StringVar(value="Services Log")
        self._services_sheet_var.trace_add("write", self._on_field_changed)
        ttk.Entry(body, textvariable=self._services_sheet_var,
                  font=F_BODY).pack(fill="x", pady=(3, 0))

        self._sheet_err = self._mk_err(body)

        # ── Password ────────────────────────────────────────────────────────
        self._mk_label(body, "Password (if file is encrypted)", top_pad=VSEP)
        pw_row = ttk.Frame(body)
        pw_row.pack(fill="x", pady=(3, 0))

        self._pw_var = tk.StringVar()
        ttk.Entry(pw_row, textvariable=self._pw_var,
                  show="*", font=F_BODY, width=22).pack(side="left")

        self._save_pw_var = tk.BooleanVar(value=False)
        if _HAS_KEYRING:
            ttk.Checkbutton(pw_row, text="Save password",
                            variable=self._save_pw_var).pack(
                side="left", padx=(10, 0))
        else:
            ttk.Label(pw_row,
                      text="(keyring unavailable — password saving disabled)",
                      foreground=FG_MUTED, font=F_SMALL,
                      background=BG).pack(side="left", padx=(10, 0))

        self._pw_status_lbl = ttk.Label(body, text="",
                                        foreground=FG_MUTED, font=F_SMALL,
                                        background=BG)
        self._pw_status_lbl.pack(anchor="w", pady=(2, 0))

        # ── Output folder ───────────────────────────────────────────────────
        self._mk_label(body, "Output Folder", top_pad=VSEP)
        fo_row = ttk.Frame(body)
        fo_row.pack(fill="x", pady=(3, 0))

        self._folder_var = tk.StringVar()
        self._folder_var.trace_add("write", self._on_field_changed)
        ttk.Entry(fo_row, textvariable=self._folder_var, state="readonly",
                  font=F_BODY).pack(side="left", fill="x", expand=True)
        ttk.Button(fo_row, text="Browse…",
                   command=self._on_browse_folder).pack(side="left", padx=(6, 0))

        self._folder_err = self._mk_err(body)

        # ── Run behavior toggles ────────────────────────────────────────────
        self._mk_label(body, "Run Behavior", top_pad=VSEP)
        self._publish_dashboard_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            body,
            text="Publish this file as current dashboard flat file",
            variable=self._publish_dashboard_var,
        ).pack(anchor="w", pady=(3, 0))
        ttk.Label(
            body,
            text="Turn this off during historical backfill loads.",
            foreground=FG_MUTED,
            font=F_SMALL,
            background=BG,
        ).pack(anchor="w", pady=(2, 0))

        # ── Run button (fixed footer) ───────────────────────────────────────
        run_frame = ttk.Frame(footer)
        run_frame.pack(fill="x", padx=PAD, pady=(10, PAD))
        self._run_btn = ttk.Button(
            run_frame,
            text="▶   Run",
            style="Accent.TButton",
            command=self._on_run,
            state="disabled",
        )
        self._run_btn.pack(fill="x")

    def _build_run_panel(self):
        f = self._run_frame

        ttk.Label(f, text="Running ETL…", font=F_HEADER).pack(
            anchor="w", padx=PAD, pady=(PAD, 0))
        ttk.Separator(f, orient="horizontal").pack(
            fill="x", padx=PAD, pady=(10, 16))

        pg_frame = ttk.Frame(f)
        pg_frame.pack(fill="x", padx=PAD)

        self._step_lbl = ttk.Label(pg_frame, text="Initialising…",
                                   foreground=FG_MUTED, font=F_LABEL,
                                   background=BG)
        self._step_lbl.pack(anchor="w", pady=(0, 6))

        self._progress = ttk.Progressbar(
            pg_frame,
            style="Accent.Horizontal.TProgressbar",
            mode="determinate",
            maximum=_TOTAL_STEPS,
            value=0,
        )
        self._progress.pack(fill="x")

        ttk.Separator(f, orient="horizontal").pack(
            fill="x", padx=PAD, pady=(16, 8))

        log_frame = ttk.Frame(f)
        log_frame.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))

        self._log_text = scrolledtext.ScrolledText(
            log_frame,
            height=20,
            font=F_MONO,
            bg=LOG_BG,
            fg=FG,
            relief="flat",
            borderwidth=1,
            state="disabled",
            wrap="word",
        )
        self._log_text.pack(fill="both", expand=True)
        self._log_text.tag_config("error", foreground=RED)
        self._log_text.tag_config("warn",  foreground=AMBER)
        self._log_text.tag_config("info",  foreground=FG)

    def _build_results_panel(self):
        f = self._results_frame
        f.columnconfigure(0, weight=1)
        f.rowconfigure(2, weight=1)

        self._status_lbl = ttk.Label(f, text="", font=F_HEADER, background=BG)
        self._status_lbl.grid(row=0, column=0, sticky="w", padx=PAD, pady=(PAD, 0))

        ttk.Separator(f, orient="horizontal").grid(
            row=1, column=0, sticky="ew", padx=PAD, pady=(10, 8))

        sum_frame = ttk.Frame(f)
        sum_frame.grid(row=2, column=0, sticky="nsew", padx=PAD, pady=(0, 4))
        sum_frame.rowconfigure(0, weight=1)
        sum_frame.columnconfigure(0, weight=1)

        self._sum_text = scrolledtext.ScrolledText(
            sum_frame,
            height=10,
            font=F_BODY,
            bg=BG,
            fg=FG,
            relief="flat",
            borderwidth=0,
            state="disabled",
            cursor="arrow",
            wrap="word",
        )
        self._sum_text.grid(row=0, column=0, sticky="nsew")
        self._sum_text.tag_config("lbl",   foreground=FG_MUTED, font=F_LABEL)
        self._sum_text.tag_config("val",   foreground=FG,       font=F_BODY)
        self._sum_text.tag_config("warn",  foreground=AMBER,    font=F_BODY)
        self._sum_text.tag_config("error", foreground=RED,      font=F_BODY)
        self._sum_text.tag_config("link",
                                  foreground=ACCENT,
                                  font=(_F, 13, "underline"))

        btn_row = ttk.Frame(f)
        btn_row.grid(row=3, column=0, sticky="ew", padx=PAD, pady=(8, PAD))

        self._open_folder_btn = ttk.Button(
            btn_row, text="Open Output Folder",
            command=self._on_open_folder)
        self._open_folder_btn.pack(side="left")

        ttk.Button(
            btn_row,
            text="▶   Run Another File",
            style="Accent.TButton",
            command=self._on_run_another,
        ).pack(side="right")

    # ── Helpers ────────────────────────────────────────────────────────────────

    @staticmethod
    def _mk_label(parent, text: str, top_pad: int = 0):
        ttk.Label(parent, text=text,
                  foreground=FG_MUTED, font=F_LABEL,
                  background=BG).pack(anchor="w", pady=(top_pad, 0))

    @staticmethod
    def _mk_err(parent) -> ttk.Label:
        lbl = ttk.Label(parent, text="",
                        foreground=RED, font=F_LABEL, background=BG)
        lbl.pack(anchor="w", pady=(2, 0))
        return lbl

    def _on_setup_canvas_configure(self, event):
        """Keep setup form width synced to canvas width."""
        self._setup_canvas.itemconfigure(self._setup_canvas_window, width=event.width)

    def _bind_setup_mousewheel(self, _event=None):
        self.bind_all("<MouseWheel>", self._on_setup_mousewheel)
        self.bind_all("<Button-4>", self._on_setup_mousewheel)  # Linux up
        self.bind_all("<Button-5>", self._on_setup_mousewheel)  # Linux down

    def _unbind_setup_mousewheel(self, _event=None):
        self.unbind_all("<MouseWheel>")
        self.unbind_all("<Button-4>")
        self.unbind_all("<Button-5>")

    def _on_setup_mousewheel(self, event):
        if getattr(event, "num", None) == 4:
            self._setup_canvas.yview_scroll(-1, "units")
            return
        if getattr(event, "num", None) == 5:
            self._setup_canvas.yview_scroll(1, "units")
            return
        delta = getattr(event, "delta", 0)
        if delta:
            self._setup_canvas.yview_scroll(int(-delta / 120), "units")

    # ── Panel navigation ───────────────────────────────────────────────────────

    def _show_panel(self, frame: ttk.Frame):
        frame.lift()

    # ── Initial state ──────────────────────────────────────────────────────────

    def _load_initial_state(self):
        settings = load_settings(self._script_dir)

        # Output folder
        saved_out = settings.get("last_output_dir", "")
        if saved_out and Path(saved_out).is_dir():
            self._folder_var.set(saved_out)
        else:
            desktop = Path.home() / "Desktop"
            if desktop.is_dir():
                self._folder_var.set(str(desktop))

        # Sheet names (main uses old key for backward compatibility)
        saved_sheet = settings.get("last_sheet_name", "")
        if saved_sheet:
            self._sheet_var.set(saved_sheet)
        saved_main_sheet = settings.get("last_main_sheet_name", "")
        if saved_main_sheet:
            self._sheet_var.set(saved_main_sheet)
        saved_vacant_sheet = settings.get("last_vacant_sheet_name", "")
        if saved_vacant_sheet:
            self._vacant_sheet_var.set(saved_vacant_sheet)
        saved_exit_sheet = settings.get("last_exit_sheet_name", "")
        if saved_exit_sheet:
            self._exit_sheet_var.set(saved_exit_sheet)
        saved_services_sheet = settings.get("last_services_sheet_name", "")
        if saved_services_sheet:
            self._services_sheet_var.set(saved_services_sheet)

        # Keyring
        if _HAS_KEYRING:
            try:
                pw = _keyring.get_password(_KR_SERVICE, _KR_USER)
                if pw:
                    self._pw_var.set(pw)
                    self._save_pw_var.set(True)
                    self._pw_status_lbl.config(
                        text="Password loaded from saved credentials")
            except Exception:
                pass

        self._validate()

    # ── Validation ─────────────────────────────────────────────────────────────

    def _on_field_changed(self, *_):
        self._validate()

    def _get_date_str(self) -> str:
        if _HAS_DATEPICKER:
            return self._date_entry.get()
        val = self._date_entry.get()
        return "" if val == "YYYY-MM-DD" else val

    def _validate(self) -> bool:
        # During early widget construction, trace callbacks can fire before
        # all error labels/buttons exist.
        if not all(
            hasattr(self, name)
            for name in (
                "_file_err",
                "_date_err",
                "_sheet_err",
                "_folder_err",
                "_run_btn",
            )
        ):
            return False

        ok = True

        # File(s)
        fp = self._file_var.get().strip()
        is_batch = len(self._batch_files) > 1
        if not is_batch and fp and not Path(fp).is_file():
            self._file_err.config(text="File not found.")
            ok = False
        else:
            self._file_err.config(text="")
        if (not is_batch) and (not fp):
            ok = False

        # Date
        if is_batch:
            valid_count = 0
            row_count = len(self._batch_date_rows)
            for row in self._batch_date_rows:
                date_str = row["date_var"].get().strip()
                if not date_str:
                    continue
                try:
                    dt.datetime.strptime(date_str, "%Y-%m-%d")
                    valid_count += 1
                except ValueError:
                    pass

            if row_count < 2 or valid_count != row_count:
                self._date_err.config(text="All batch files need valid YYYY-MM-DD dates.")
                ok = False
            else:
                self._date_err.config(text="")
        else:
            date_str = self._get_date_str()
            if date_str:
                try:
                    dt.datetime.strptime(date_str, "%Y-%m-%d")
                    self._date_err.config(text="")
                except ValueError:
                    self._date_err.config(text="Must be YYYY-MM-DD.")
                    ok = False
            else:
                self._date_err.config(text="")
                ok = False

        # Sheets
        if (
            not self._sheet_var.get().strip()
            or not self._vacant_sheet_var.get().strip()
            or not self._exit_sheet_var.get().strip()
            or not self._services_sheet_var.get().strip()
        ):
            self._sheet_err.config(text="All worksheet names are required.")
            ok = False
        else:
            self._sheet_err.config(text="")

        # Folder
        folder = self._folder_var.get().strip()
        if folder:
            if Path(folder).is_file():
                self._folder_err.config(text="Path points to a file, not a folder.")
                ok = False
            else:
                self._folder_err.config(text="")
        else:
            self._folder_err.config(text="")
            ok = False

        self._run_btn.config(state="normal" if ok else "disabled")
        return ok

    # ── Date placeholder (plain-Entry fallback) ────────────────────────────────

    def _on_date_focus_in(self, _=None):
        if not _HAS_DATEPICKER and self._date_entry.get() == "YYYY-MM-DD":
            self._date_entry.delete(0, tk.END)
            self._date_entry.configure(foreground=FG)

    def _on_date_focus_out(self, _=None):
        if not _HAS_DATEPICKER and not self._date_entry.get():
            self._date_entry.insert(0, "YYYY-MM-DD")
            self._date_entry.configure(foreground=FG_MUTED)

    # ── Browse ─────────────────────────────────────────────────────────────────

    def _on_browse_file(self):
        self._batch_files = []
        self._batch_info_lbl.config(text="")
        self._set_date_mode(False)
        cur = self._file_var.get()
        init = str(Path(cur).parent) if cur else str(Path.home())
        path = filedialog.askopenfilename(
            title="Select ICMS Snapshot",
            initialdir=init,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._file_var.set(path)
            self._auto_fill_from_file(path)

    def _on_browse_files(self):
        cur = self._file_var.get()
        init = str(Path(cur).parent) if cur else str(Path.home())
        paths = filedialog.askopenfilenames(
            title="Select Multiple ICMS Snapshots",
            initialdir=init,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not paths:
            return
        if len(paths) == 1:
            self._batch_files = []
            self._batch_info_lbl.config(text="")
            self._set_date_mode(False)
            self._file_var.set(paths[0])
            self._auto_fill_from_file(paths[0])
            return
        self._batch_files = [Path(p) for p in paths]
        first = self._batch_files[0].name
        last = self._batch_files[-1].name
        self._file_var.set(str(self._batch_files[0]))
        self._batch_info_lbl.config(
            text=f"Batch selected: {len(self._batch_files)} files ({first} ... {last})"
        )
        self._set_date_mode(True)
        self._render_batch_date_rows(self._batch_files)
        self._validate()

    def _set_date_input_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self._date_entry.configure(state=state)

    def _clear_batch_date_rows(self):
        for row in self._batch_date_rows:
            container = row.get("container")
            if container is not None:
                container.destroy()
        self._batch_date_rows = []

    def _render_batch_date_rows(self, files: List[Path]):
        self._clear_batch_date_rows()
        for path in files:
            row = ttk.Frame(self._batch_date_list)
            row.pack(fill="x", pady=(0, 6))

            file_var = tk.StringVar(value=path.name)
            date_var = tk.StringVar(value=_parse_date_from_filename(path.name) or "")
            date_var.trace_add("write", self._on_field_changed)

            file_entry = ttk.Entry(
                row,
                textvariable=file_var,
                font=F_LABEL,
            )
            file_entry.pack(side="left", fill="x", expand=True)
            m = re.search(r"\d{2}\.\d{2}\.\d{4}", path.name)
            jump_idx = m.start() if m else 0
            file_entry.bind(
                "<FocusIn>",
                lambda e, idx=jump_idx: self._focus_filename_entry_at(e.widget, idx),
            )
            ttk.Entry(
                row,
                textvariable=date_var,
                font=F_LABEL,
                width=12,
            ).pack(side="left", padx=(6, 0))

            self._batch_date_rows.append(
                {
                    "path": path,
                    "file_var": file_var,
                    "date_var": date_var,
                    "container": row,
                }
            )

    @staticmethod
    def _focus_filename_entry_at(widget: tk.Widget, idx: int):
        entry = widget
        try:
            entry.icursor(idx)
            entry.xview(idx)
        except Exception:
            pass

    def _set_date_mode(self, is_batch: bool):
        if is_batch:
            self._set_date_input_enabled(False)
            if self._date_lbl.winfo_ismapped():
                self._date_lbl.pack_forget()
            if self._date_row.winfo_ismapped():
                self._date_row.pack_forget()
            if not self._batch_date_frame.winfo_ismapped():
                self._batch_date_frame.pack(fill="x", pady=(6, 0))
        else:
            self._set_date_input_enabled(True)
            if not self._date_lbl.winfo_ismapped():
                self._date_lbl.pack(anchor="w", pady=(VSEP, 0), before=self._batch_date_frame)
            if not self._date_row.winfo_ismapped():
                self._date_row.pack(anchor="w", pady=(3, 0), before=self._batch_date_frame)
            if self._batch_date_frame.winfo_ismapped():
                self._batch_date_frame.pack_forget()
            self._clear_batch_date_rows()

    def _collect_batch_dates(self) -> List[Tuple[Path, str]]:
        out: List[Tuple[Path, str]] = []
        for row in self._batch_date_rows:
            path = row["path"]
            snap = row["date_var"].get().strip()
            out.append((path, snap))
        return out

    def _load_integrated_file_keys(self, output_dir: Path) -> set:
        """Return {(source_file_name, snapshot_date)} already integrated."""
        keys: set = set()
        history_path = output_dir / "icms_snapshot_history.csv"
        if history_path.exists():
            try:
                hist = pd.read_csv(history_path, dtype=str, encoding="utf-8").fillna("")
                if {"source_file_name", "snapshot_date"}.issubset(hist.columns):
                    for _, r in hist[["source_file_name", "snapshot_date"]].drop_duplicates().iterrows():
                        src = str(r["source_file_name"]).strip()
                        snap = str(r["snapshot_date"]).strip()
                        if src and snap:
                            keys.add((src, snap))
            except Exception:
                pass

        ingest_path = output_dir / "icms_ingest_log.csv"
        if ingest_path.exists():
            try:
                ing = pd.read_csv(ingest_path, dtype=str, encoding="utf-8").fillna("")
                cols = set(ing.columns)
                if {"source_file_name", "snapshot_date", "run_status"}.issubset(cols):
                    ok = ing[ing["run_status"].str.lower() == "success"]
                    for _, r in ok[["source_file_name", "snapshot_date"]].drop_duplicates().iterrows():
                        src = str(r["source_file_name"]).strip()
                        snap = str(r["snapshot_date"]).strip()
                        if src and snap:
                            keys.add((src, snap))
            except Exception:
                pass
        return keys

    def _find_duplicate_candidates(
        self,
        candidates: List[Tuple[Path, str]],
        integrated_keys: set,
    ) -> List[Tuple[Path, str]]:
        dups: List[Tuple[Path, str]] = []
        for p, snap in candidates:
            key = (p.name.strip(), str(snap).strip())
            if key in integrated_keys:
                dups.append((p, snap))
        return dups

    def _services_table_already_integrated(
        self,
        path: Path,
        snapshot_date: str,
        services_sheet_name: str,
        output_dir: Path,
        overlap_threshold: float = 0.98,
    ) -> Tuple[bool, Dict[str, float]]:
        """Return True when incoming Services Log table appears already loaded."""
        metrics: Dict[str, float] = {
            "incoming_event_key_count": 0,
            "existing_overlap_rate": 0.0,
            "already_integrated_warning": 0.0,
        }
        try:
            wb, _ = core.open_workbook(path, password=self._pw_var.get())
        except Exception:
            return False, metrics

        try:
            services_df = core.read_services_log(wb, sheet_name=services_sheet_name)
        finally:
            wb.close()

        overlap = core.summarize_services_overlap(services_df, output_dir)
        incoming_count = int(overlap.get("incoming_event_key_count", 0) or 0)
        overlap_rate = float(overlap.get("existing_overlap_rate", 0.0) or 0.0)
        metrics["incoming_event_key_count"] = incoming_count
        metrics["existing_overlap_rate"] = overlap_rate

        if incoming_count <= 0:
            return False, metrics

        # Strong signal #1: same source file + snapshot date already in services history.
        path_hist = output_dir / "icms_services_log_history.csv"
        same_file_date = False
        if path_hist.exists():
            try:
                hist = pd.read_csv(path_hist, dtype=str, encoding="utf-8").fillna("")
                if {"source_file_name", "snapshot_date"}.issubset(hist.columns):
                    same_file_date = bool(
                        (
                            (hist["source_file_name"].astype(str).str.strip() == path.name)
                            & (hist["snapshot_date"].astype(str).str.strip() == str(snapshot_date).strip())
                        ).any()
                    )
            except Exception:
                same_file_date = False

        already = bool(same_file_date or overlap_rate >= overlap_threshold)
        metrics["already_integrated_warning"] = 1.0 if already else 0.0
        return already, metrics

    def _find_services_duplicate_candidates(
        self,
        candidates: List[Tuple[Path, str]],
        services_sheet_name: str,
        output_dir: Path,
    ) -> List[Tuple[Path, str]]:
        out: List[Tuple[Path, str]] = []
        for p, snap in candidates:
            already, _ = self._services_table_already_integrated(
                p, snap, services_sheet_name, output_dir
            )
            if already:
                out.append((p, snap))
        return out

    def _auto_fill_from_file(self, filepath: str):
        parsed = _parse_date_from_filename(Path(filepath).name)
        if parsed:
            self._set_date(parsed)
        else:
            if not _HAS_DATEPICKER:
                self._date_entry.focus_set()

    def _set_date(self, date_str: str):
        """Set the date widget to *date_str* (YYYY-MM-DD)."""
        if _HAS_DATEPICKER:
            try:
                y, m, d = (int(x) for x in date_str.split("-"))
                self._date_entry.set_date(dt.date(y, m, d))
            except Exception:
                pass
        else:
            self._date_entry.configure(foreground=FG)
            self._date_entry.delete(0, tk.END)
            self._date_entry.insert(0, date_str)

    def _on_browse_folder(self):
        init = self._folder_var.get() or str(Path.home() / "Desktop")
        path = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=init,
            mustexist=False,
        )
        if path:
            self._folder_var.set(path)

    # ── Run ────────────────────────────────────────────────────────────────────

    def _on_run(self):
        if not self._validate():
            return

        input_path  = Path(self._file_var.get()) if self._file_var.get().strip() else None
        snap_date   = self._get_date_str()
        sheet_name  = self._sheet_var.get().strip()
        vacant_sheet_name = self._vacant_sheet_var.get().strip()
        exit_sheet_name = self._exit_sheet_var.get().strip()
        services_sheet_name = self._services_sheet_var.get().strip()
        publish_dashboard = self._publish_dashboard_var.get()
        password    = self._pw_var.get()
        output_dir  = Path(self._folder_var.get())
        batch_files = list(self._batch_files)

        confirmed_batch: Optional[List[Tuple[Path, str]]] = None
        if len(batch_files) > 1:
            confirmed_batch = self._collect_batch_dates()
        else:
            if input_path is None:
                messagebox.showerror("Error", "Select an input file first.")
                return

        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Error",
                                 f"Could not create output directory:\n{exc}")
            return

        integrated_keys = self._load_integrated_file_keys(output_dir)
        self._last_skipped_duplicates = []
        self._last_skipped_services_duplicates = []

        self._current_output_dir = output_dir

        # Duplicate warning / exclusion handling
        if len(batch_files) > 1 and confirmed_batch is not None:
            dup_candidates = self._find_duplicate_candidates(confirmed_batch, integrated_keys)
            if dup_candidates:
                dup_dialog = _BatchDuplicateDialog(self, dup_candidates)
                self.wait_window(dup_dialog)
                if dup_dialog.action == "cancel":
                    return
                if dup_dialog.action == "exclude":
                    dup_set = {(p.name, s) for p, s in dup_candidates}
                    filtered = [(p, s) for (p, s) in confirmed_batch if (p.name, s) not in dup_set]
                    self._last_skipped_duplicates = dup_candidates
                    if not filtered:
                        messagebox.showinfo(
                            "Nothing to ingest",
                            "All selected files were already integrated and excluded.",
                        )
                        return
                    confirmed_batch = filtered
                # include => keep confirmed_batch unchanged
        elif input_path is not None:
            single_tuple = [(input_path, snap_date)]
            dup_candidates = self._find_duplicate_candidates(single_tuple, integrated_keys)
            if dup_candidates:
                proceed = messagebox.askokcancel(
                    "File appears already integrated",
                    (
                        f"'{input_path.name}' with snapshot date {snap_date} appears to have "
                        "already been integrated into outputs in this folder.\n\n"
                        "Continue anyway?"
                    ),
                )
                if not proceed:
                    return

        # Services-table-level already-ingested warning / exclusion handling
        if len(batch_files) > 1 and confirmed_batch is not None:
            svc_dups = self._find_services_duplicate_candidates(
                confirmed_batch, services_sheet_name, output_dir
            )
            if svc_dups:
                svc_dialog = _BatchServicesDuplicateDialog(self, svc_dups)
                self.wait_window(svc_dialog)
                if svc_dialog.action == "cancel":
                    return
                if svc_dialog.action == "exclude":
                    dup_set = {(p.name, s) for p, s in svc_dups}
                    filtered = [(p, s) for (p, s) in confirmed_batch if (p.name, s) not in dup_set]
                    self._last_skipped_services_duplicates = svc_dups
                    if not filtered:
                        messagebox.showinfo(
                            "Nothing to ingest",
                            "All selected files were excluded because their Services Log table "
                            "appears already integrated.",
                        )
                        return
                    confirmed_batch = filtered
        elif input_path is not None:
            svc_already, _ = self._services_table_already_integrated(
                input_path, snap_date, services_sheet_name, output_dir
            )
            if svc_already:
                proceed = messagebox.askokcancel(
                    "Services Log appears already integrated",
                    (
                        f"The Services Log table in '{input_path.name}' appears already "
                        "integrated into this output folder.\n\n"
                        "Continue anyway?"
                    ),
                )
                if not proceed:
                    return

        # Reset Run panel
        self._progress["value"] = 0
        self._step_lbl.config(text="Starting…")
        self._log_text.config(state="normal")
        self._log_text.delete("1.0", tk.END)
        self._log_text.config(state="disabled")

        self._show_panel(self._run_frame)

        threading.Thread(
            target=self._etl_worker,
            args=(
                input_path,
                snap_date,
                confirmed_batch,
                sheet_name,
                vacant_sheet_name,
                exit_sheet_name,
                services_sheet_name,
                publish_dashboard,
                password,
                output_dir,
            ),
            daemon=True,
        ).start()

    # ── ETL worker (background thread) ────────────────────────────────────────

    def _etl_worker(
        self,
        input_path: Optional[Path],
        snapshot_date: str,
        batch_files: Optional[List[Tuple[Path, str]]],
        sheet_name: str,
        vacant_sheet_name: str,
        exit_sheet_name: str,
        services_sheet_name: str,
        publish_dashboard: bool,
        password: str,
        output_dir: Path,
    ):
        if batch_files:
            self._run_batch_worker(
                batch_files=batch_files,
                sheet_name=sheet_name,
                vacant_sheet_name=vacant_sheet_name,
                exit_sheet_name=exit_sheet_name,
                services_sheet_name=services_sheet_name,
                password=password,
                output_dir=output_dir,
                publish_latest_only=True,
            )
            return

        if input_path is None:
            self._gui_q.put({
                "type": "error",
                "msg": "No input file selected.",
                "tb": "",
            })
            return

        self._process_one_file(
            input_path=input_path,
            snapshot_date=snapshot_date,
            sheet_name=sheet_name,
            vacant_sheet_name=vacant_sheet_name,
            exit_sheet_name=exit_sheet_name,
            services_sheet_name=services_sheet_name,
            publish_dashboard=publish_dashboard,
            password=password,
            output_dir=output_dir,
            aliases=None,
            emit_done_event=True,
        )

    def _process_one_file(
        self,
        *,
        input_path: Path,
        snapshot_date: str,
        sheet_name: str,
        vacant_sheet_name: str,
        exit_sheet_name: str,
        services_sheet_name: str,
        publish_dashboard: bool,
        password: str,
        output_dir: Path,
        aliases: Optional[Dict[str, str]],
        emit_done_event: bool,
    ) -> dict:
        gq = self._gui_q
        rq = self._resp_q

        def log(msg: str, level: str = "info"):
            ts = dt.datetime.now().strftime("%H:%M:%S")
            gq.put({"type": "log", "msg": f"[{ts}]  {msg}", "level": level})

        def step(n: int):
            gq.put({"type": "step", "n": n, "label": _STEP_LABELS[n - 1]})

        try:
            source_file_name = input_path.name
            run_ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            notes: List[str] = []
            read_stats: Dict[str, object] = {}
            anomalies_diag: List[str] = []
            overlap_rate_diag = 0.0
            incoming_unique_key_count_diag = 0
            incoming_blank_champ_id_count_diag = 0
            incoming_blank_project_name_count_diag = 0
            incoming_duplicate_key_count_diag = 0
            services_metrics: Dict[str, float] = {
                "services_log_records_processed": 0,
                "services_log_history_total_rows": 0,
                "services_log_duplicate_records_dropped-diag": 0,
                "services_log_blank_service_id_count-diag": 0,
                "services_log_incoming_event_key_count-diag": 0,
                "services_log_existing_key_overlap_rate-diag": 0.0,
                "services_log_already_integrated_warning-diag": 0,
                "services_log_missing_sheet-diag": 0,
            }

            # ── 1. Open workbook ───────────────────────────────────────────
            step(1)
            log(f"Opening workbook: {input_path.name}")
            # Pass password as-is (empty string is valid; avoids getpass in GUI)
            wb, was_encrypted = core.open_workbook(input_path, password=password)
            if was_encrypted:
                log("File was password-protected — decrypted in memory")

            sheets = core.list_sheet_names(wb)
            if sheet_name not in sheets:
                wb.close()
                raise ValueError(
                    f"Sheet '{sheet_name}' not found.\n"
                    f"Available sheets: {', '.join(sheets)}"
                )

            # ── 2. Read sheets ─────────────────────────────────────────────
            step(2)
            log(f"Reading sheet: {sheet_name}")
            try:
                header_row_idx = core.find_header_row(
                    wb, sheet_name, core.ALL_APPROVED_FIELDS
                )
                df_raw = core.read_sheet(
                    wb, sheet_name, header_row=header_row_idx, stats=read_stats
                )
                vacant_df = pd.DataFrame()
                if vacant_sheet_name in sheets:
                    log(f"Reading sheet: {vacant_sheet_name}")
                    vacant_raw = core.read_sheet(wb, vacant_sheet_name)
                    vacant_keep_cols = [
                        "ICMS Project Name",
                        "ICMS Provider",
                        "PH Building Name",
                        "ICMS Project Type",
                        "# Days Vacant",
                    ]
                    vacant_avail = [c for c in vacant_keep_cols if c in vacant_raw.columns]
                    vacant_df = vacant_raw[vacant_avail].copy()
                else:
                    log(
                        f"WARNING: Sheet '{vacant_sheet_name}' not found; "
                        "vacant slot data will be skipped for this run.",
                        level="warn",
                    )
                    notes.append(f"Missing sheet: {vacant_sheet_name} (vacant data skipped)")

                exit_df = pd.DataFrame()
                if exit_sheet_name in sheets:
                    log(f"Reading sheet: {exit_sheet_name}")
                    exit_raw = core.read_sheet(wb, exit_sheet_name)
                    exit_keep_cols = [
                        "CHAMP ID",
                        "ICMS Project Name",
                        "ICMS Provider",
                        "PH Building Name",
                        "ICMS Project Type",
                        "ICMS Enrollment ID",
                        "ICMS Case Manager",
                        "First Name",
                        "Last Name",
                        "Date of Birth",
                        "HMIS ID",
                        "Medi-Cal CIN",
                        "Check-In Date",
                        "Check-Out Date",
                        "Move-In Date",
                        "Move-Out Date",
                        "Exit Reason",
                        "Exit Destination",
                        "Exit Location",
                        "Transferred to New Provider Within ICMS Program?",
                        "New ICMS Provider",
                    ]
                    exit_avail = [c for c in exit_keep_cols if c in exit_raw.columns]
                    exit_df = exit_raw[exit_avail].copy()
                else:
                    log(
                        f"WARNING: Sheet '{exit_sheet_name}' not found; "
                        "exit-log data will be skipped for this run.",
                        level="warn",
                    )
                    notes.append(f"Missing sheet: {exit_sheet_name} (exit data skipped)")

                services_df = pd.DataFrame()
                if services_sheet_name in sheets:
                    log(f"Reading sheet: {services_sheet_name}")
                    services_df = core.read_services_log(wb, sheet_name=services_sheet_name)
                else:
                    services_metrics["services_log_missing_sheet-diag"] = 1
                    log(
                        f"WARNING: Sheet '{services_sheet_name}' not found; "
                        "services-log data will be skipped for this run.",
                        level="warn",
                    )
                    notes.append(f"Missing sheet: {services_sheet_name} (services data skipped)")
            finally:
                wb.close()

            vacant_rows    = core.build_vacant_slot_rows(
                vacant_df, snapshot_date, source_file_name)
            input_row_count = len(df_raw)

            log(f"Participant rows read: {input_row_count}")
            log(f"Vacant slot records:   {len(vacant_df)}")
            log(f"Exit log records:      {len(exit_df)}")
            log(f"Services log records:  {len(services_df)}")

            if input_row_count == 0:
                raise ValueError("No data rows found in the selected sheet.")

            # Lightweight key-health diagnostics from raw participant rows.
            if "CHAMP ID" in df_raw.columns:
                incoming_blank_champ_id_count_diag = int(
                    df_raw["CHAMP ID"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .eq("")
                    .sum()
                )
            if "ICMS Project Name" in df_raw.columns:
                incoming_blank_project_name_count_diag = int(
                    df_raw["ICMS Project Name"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .eq("")
                    .sum()
                )

            # ── 3. Resolve fields ──────────────────────────────────────────
            step(3)
            log("Resolving field names…")
            aliases_map = aliases if aliases is not None else core.load_aliases(self._script_dir)
            wb_cols = [c for c in df_raw.columns if not c.startswith("_")]

            resolved, missing_req, missing_opt, aliases_added = (
                _gui_resolve_fields(wb_cols, aliases_map, self._script_dir, gq, rq)
            )

            if missing_req:
                notes.append(f"Missing required fields: {', '.join(missing_req)}")
                log(
                    f"WARNING: {len(missing_req)} required field(s) unresolved: "
                    f"{', '.join(missing_req)}",
                    level="warn",
                )
            log(
                f"Fields resolved: {len(resolved)} of "
                f"{len(core.ALL_APPROVED_FIELDS)}"
            )
            if aliases_added:
                log(f"{aliases_added} alias(es) saved to {core.ALIASES_FILENAME}")

            df = core.reduce_to_approved_fields(
                df_raw, resolved, snapshot_date, source_file_name, stats=read_stats)

            dropped_blank_key_rows_diag = int(
                read_stats.get("dropped_blank_key_rows-diag", 0) or 0
            )
            if dropped_blank_key_rows_diag > 0:
                msg = (
                    f"Dropped {dropped_blank_key_rows_diag} row(s) with blank "
                    "CHAMP ID or ICMS Project Name"
                )
                anomalies_diag.append(msg)
                log(f"DIAG WARNING: {msg}", level="warn")

            if all(k in df.columns for k in core.KEY_FIELDS):
                incoming_keys_df = (
                    df[core.KEY_FIELDS]
                    .fillna("")
                    .astype(str)
                    .apply(lambda s: s.str.strip())
                )
                incoming_unique_key_count_diag = int(
                    incoming_keys_df.drop_duplicates().shape[0]
                )
                incoming_duplicate_key_count_diag = int(
                    max(0, len(incoming_keys_df) - incoming_unique_key_count_diag)
                )

            # Duplicate detection
            dup_df         = core.find_duplicates(df)
            dup_flag_count = len(dup_df)
            if dup_flag_count > 0:
                n_pairs = (
                    dup_df.groupby(core.KEY_FIELDS, sort=False).ngroups
                    if all(k in dup_df.columns for k in core.KEY_FIELDS)
                    else 0
                )
                log(
                    f"WARNING: {n_pairs} duplicate CHAMP ID + Project pair(s) "
                    f"({dup_flag_count} flagged rows)",
                    level="warn",
                )
                notes.append(
                    f"Duplicate CHAMP ID / ICMS Project Name pairs: {n_pairs}")

            vacant_slot_count = len(vacant_rows)

            # ── 4. Dashboard flat file ─────────────────────────────────────
            step(4)
            dashboard_count = 0
            if publish_dashboard:
                log("Writing dashboard flat file…")
                core.update_dashboard_flatfile(df, output_dir, vacant_rows)
                dashboard_count = len(df) + vacant_slot_count
                log(
                    f"Dashboard flat file: {dashboard_count} rows "
                    f"({vacant_slot_count} vacant slot rows)"
                )
            else:
                log("Dashboard publish disabled for this run (historical mode).")
                notes.append("Dashboard flat file not updated (publish disabled)")

            # ── 5. Snapshot history ────────────────────────────────────────
            step(5)
            log("Updating snapshot history…")
            history_count = core.update_snapshot_history(
                df, snapshot_date, output_dir)
            log(f"Snapshot history: {history_count} total rows")
            # ── 6. Client/project master ───────────────────────────────────
            step(6)
            log("Updating client/project master…")
            (
                master_total, new_rows, returning_rows,
                not_present, exited_count, unknown_count,
            ) = core.update_client_project_master(
                df, snapshot_date, source_file_name, output_dir, exit_df)
            if (new_rows + returning_rows) > 0:
                overlap_rate_diag = round(
                    (returning_rows / (new_rows + returning_rows)), 4
                )
            log(
                f"Master: {master_total} rows  "
                f"({new_rows} new, {returning_rows} returning, "
                f"{not_present} absent, {exited_count} exited, "
                f"{unknown_count} unknown)"
            )

            key_norm_diag = int(read_stats.get("key_field_float_normalizations-diag", 0) or 0)
            if key_norm_diag > 0:
                msg = (
                    f"{key_norm_diag} key-field cell(s) were normalized from numeric "
                    "to integer string values"
                )
                anomalies_diag.append(msg)
                log(f"DIAG WARNING: {msg}", level="warn")

            if returning_rows == 0 and new_rows > 50:
                msg = (
                    f"0 returning keys with {new_rows} new keys; possible key-format, "
                    "sheet, or header mismatch"
                )
                anomalies_diag.append(msg)
                log(f"DIAG WARNING: {msg}", level="warn")

            if incoming_blank_champ_id_count_diag > 0:
                msg = (
                    f"{incoming_blank_champ_id_count_diag} incoming rows have blank CHAMP ID"
                )
                anomalies_diag.append(msg)
                log(f"DIAG WARNING: {msg}", level="warn")

            if incoming_blank_project_name_count_diag > 0:
                msg = (
                    f"{incoming_blank_project_name_count_diag} incoming rows have blank "
                    "ICMS Project Name"
                )
                anomalies_diag.append(msg)
                log(f"DIAG WARNING: {msg}", level="warn")

            if master_total > 0 and (new_rows + returning_rows) > 0 and overlap_rate_diag < 0.10:
                msg = (
                    f"Low incoming key overlap rate with existing master: "
                    f"{overlap_rate_diag:.4f}"
                )
                anomalies_diag.append(msg)
                log(f"DIAG WARNING: {msg}", level="warn")

            # ── 7. Write logs ──────────────────────────────────────────────
            step(7)

            overlap_summary = core.summarize_services_overlap(services_df, output_dir)
            incoming_event_keys_diag = int(overlap_summary.get("incoming_event_key_count", 0) or 0)
            existing_overlap_rate_diag = float(overlap_summary.get("existing_overlap_rate", 0.0) or 0.0)
            services_metrics["services_log_incoming_event_key_count-diag"] = incoming_event_keys_diag
            services_metrics["services_log_existing_key_overlap_rate-diag"] = existing_overlap_rate_diag
            if incoming_event_keys_diag > 0 and existing_overlap_rate_diag >= 0.98:
                services_metrics["services_log_already_integrated_warning-diag"] = 1
                msg = (
                    "Services Log table appears already integrated "
                    f"(overlap={existing_overlap_rate_diag:.4f})"
                )
                anomalies_diag.append(msg)
                log(f"DIAG WARNING: {msg}", level="warn")

            services_update = core.update_services_log_history(
                services_df, snapshot_date, source_file_name, output_dir
            )
            services_metrics.update(services_update)
            log(
                "Services log history: "
                f"{int(services_metrics.get('services_log_history_total_rows', 0))} total rows "
                f"({int(services_metrics.get('services_log_duplicate_records_dropped-diag', 0))} "
                "duplicates dropped)"
            )

            if dup_flag_count > 0:
                try:
                    warn_df = core.build_duplicate_warning_rows(
                        dup_df, run_ts, snapshot_date, source_file_name)
                    core.append_duplicate_warning(
                        warn_df, snapshot_date, output_dir)
                    log(f"Duplicate warning file: {dup_flag_count} rows written")
                except RuntimeError as exc:
                    log(f"WARNING: Could not write duplicate warning: {exc}",
                        level="warn")
                    notes.append(f"Duplicate warning write failed: {exc}")

            log_row = {
                "run_timestamp":                        run_ts,
                "snapshot_date":                        snapshot_date,
                "source_file_name":                     source_file_name,
                "sheet_name":                           sheet_name,
                "input_row_count":                      input_row_count,
                "vacant_slot_rows_injected":            vacant_slot_count,
                "exit_log_records_processed":           len(exit_df),
                "current_dashboard_row_count":          dashboard_count,
                "history_total_row_count_after_update": history_count,
                "master_row_count_after_update":        master_total,
                "new_client_project_rows":              new_rows,
                "returning_client_project_rows":        returning_rows,
                "not_present_in_latest_snapshot_count": not_present,
                "master_exited_count":                  exited_count,
                "master_unknown_count":                 unknown_count,
                "required_fields_missing_count":        len(missing_req),
                "optional_fields_missing_count":        len(missing_opt),
                "field_aliases_added_count":            aliases_added,
                "duplicate_client_project_count":       dup_flag_count,
                "services_log_records_processed": int(
                    services_metrics.get("services_log_records_processed", 0)
                ),
                "services_log_history_total_rows": int(
                    services_metrics.get("services_log_history_total_rows", 0)
                ),
                "services_log_duplicate_records_dropped-diag": int(
                    services_metrics.get("services_log_duplicate_records_dropped-diag", 0)
                ),
                "services_log_missing_sheet-diag": int(
                    services_metrics.get("services_log_missing_sheet-diag", 0)
                ),
                "services_log_blank_service_id_count-diag": int(
                    services_metrics.get("services_log_blank_service_id_count-diag", 0)
                ),
                "services_log_incoming_event_key_count-diag": int(
                    services_metrics.get("services_log_incoming_event_key_count-diag", 0)
                ),
                "services_log_existing_key_overlap_rate-diag": float(
                    services_metrics.get("services_log_existing_key_overlap_rate-diag", 0.0)
                ),
                "services_log_already_integrated_warning-diag": int(
                    services_metrics.get("services_log_already_integrated_warning-diag", 0)
                ),
                "key_field_float_normalizations-diag": int(
                    read_stats.get("key_field_float_normalizations-diag", 0) or 0
                ),
                "incoming_unique_key_count-diag": incoming_unique_key_count_diag,
                "incoming_blank_champ_id_count-diag": incoming_blank_champ_id_count_diag,
                "incoming_blank_project_name_count-diag": incoming_blank_project_name_count_diag,
                "incoming_duplicate_key_count-diag": incoming_duplicate_key_count_diag,
                "incoming_key_overlap_rate_with_existing_master-diag": overlap_rate_diag,
                "header_row_detected-diag": int(
                    read_stats.get("header_row_detected-diag", header_row_idx)
                ),
                "participant_column_count-diag": int(
                    read_stats.get("participant_column_count-diag", 0) or 0
                ),
                "dropped_blank_key_rows-diag": dropped_blank_key_rows_diag,
                "missing_required_fields-diag": ", ".join(missing_req),
                "missing_optional_fields-diag": ", ".join(missing_opt),
                "data_anomalies-diag": "; ".join(anomalies_diag),
                "password_protected_input":             was_encrypted,
                "run_status":                           "success",
                "notes": "; ".join(notes) if notes else "",
            }
            try:
                core.append_ingest_log(log_row, output_dir)
                log("Ingest log updated")
            except RuntimeError as exc:
                log(f"WARNING: Could not write ingest log: {exc}", level="warn")

            log("✓  ETL complete!")

            result = {
                "success":          True,
                "source_file_name": source_file_name,
                "sheet_name":       sheet_name,
                "vacant_sheet_name": vacant_sheet_name,
                "exit_sheet_name":  exit_sheet_name,
                "services_sheet_name": services_sheet_name,
                "snapshot_date":    snapshot_date,
                "dashboard_count":  dashboard_count,
                "dashboard_published": publish_dashboard,
                "vacant_slot_count": vacant_slot_count,
                "exit_log_count":   len(exit_df),
                "services_log_count": int(
                    services_metrics.get("services_log_records_processed", 0)
                ),
                "services_log_history_count": int(
                    services_metrics.get("services_log_history_total_rows", 0)
                ),
                "services_log_dupes_dropped": int(
                    services_metrics.get("services_log_duplicate_records_dropped-diag", 0)
                ),
                "history_count":    history_count,
                "master_total":     master_total,
                "exited_count":     exited_count,
                "unknown_count":    unknown_count,
                "dup_flag_count":   dup_flag_count,
                "missing_req":      missing_req,
                "output_dir":       str(output_dir.resolve()),
                "was_encrypted":    was_encrypted,
                "notes": "; ".join(notes) if notes else "",
            }
            if emit_done_event:
                gq.put({"type": "done", "results": result})
            return result

        except Exception as exc:
            err_msg = str(exc).strip() or exc.__class__.__name__
            tb = traceback.format_exc()
            log(f"ERROR: {err_msg}", level="error")
            log(tb, level="error")
            result = {
                "success": False,
                "source_file_name": input_path.name,
                "snapshot_date": snapshot_date,
                "error": err_msg,
                "tb": tb,
            }
            if emit_done_event:
                gq.put({
                    "type": "error",
                    "msg": err_msg,
                    "tb": tb,
                })
            return result

    def _run_batch_worker(
        self,
        *,
        batch_files: List[Tuple[Path, str]],
        sheet_name: str,
        vacant_sheet_name: str,
        exit_sheet_name: str,
        services_sheet_name: str,
        password: str,
        output_dir: Path,
        publish_latest_only: bool,
    ):
        gq = self._gui_q
        aliases = core.load_aliases(self._script_dir)
        confirmed = sorted(batch_files, key=lambda x: x[1])
        total = len(confirmed)
        results: List[dict] = []

        latest_idx = total - 1
        for idx, (path, snap) in enumerate(confirmed):
            publish_dashboard = (idx == latest_idx) if publish_latest_only else True
            gq.put({
                "type": "log",
                "msg": f"[{dt.datetime.now().strftime('%H:%M:%S')}]  "
                       f"Batch file {idx + 1} of {total}: {path.name} ({snap})",
                "level": "info",
            })
            result = self._process_one_file(
                input_path=path,
                snapshot_date=snap,
                sheet_name=sheet_name,
                vacant_sheet_name=vacant_sheet_name,
                exit_sheet_name=exit_sheet_name,
                services_sheet_name=services_sheet_name,
                publish_dashboard=publish_dashboard,
                password=password,
                output_dir=output_dir,
                aliases=aliases,
                emit_done_event=False,
            )
            results.append(result)

        success_count = sum(1 for r in results if r.get("success"))
        fail_count = total - success_count
        failed_results = [r for r in results if not r.get("success")]
        first_error = ""
        if failed_results:
            first_error = str(failed_results[0].get("error", "")).strip()
        if not first_error and fail_count > 0:
            first_error = "One or more batch files failed."
        newest_published = next(
            (
                r.get("source_file_name", "")
                for r in reversed(results)
                if r.get("success") and r.get("dashboard_published")
            ),
            "",
        )
        final_master = next(
            (r.get("master_total", 0) for r in reversed(results) if r.get("success")),
            0,
        )

        gq.put({
            "type": "done",
            "results": {
                "success": success_count > 0,
                "is_batch": True,
                "batch_total": total,
                "batch_success": success_count,
                "batch_failed": fail_count,
                "batch_dashboard_source": newest_published,
                "batch_results": results,
                "batch_failed_files": [
                    {
                        "source_file_name": r.get("source_file_name", ""),
                        "snapshot_date": r.get("snapshot_date", ""),
                        "error": str(r.get("error", "")).strip() or "Unspecified error",
                    }
                    for r in failed_results
                ],
                "master_total": final_master,
                "error": first_error if fail_count > 0 else "",
                "skipped_duplicates": [
                    {"source_file_name": p.name, "snapshot_date": s}
                    for p, s in self._last_skipped_duplicates
                ],
                "skipped_services_duplicates": [
                    {"source_file_name": p.name, "snapshot_date": s}
                    for p, s in self._last_skipped_services_duplicates
                ],
                "sheet_name": sheet_name,
                "vacant_sheet_name": vacant_sheet_name,
                "exit_sheet_name": exit_sheet_name,
                "services_sheet_name": services_sheet_name,
                "output_dir": str(output_dir.resolve()),
            },
        })

    # ── Queue polling (main thread) ────────────────────────────────────────────

    def _poll_queue(self):
        try:
            while True:
                event = self._gui_q.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _handle_event(self, event: dict):
        etype = event["type"]

        if etype == "step":
            self._progress["value"] = event["n"]
            self._step_lbl.config(text=event["label"])

        elif etype == "log":
            level = event.get("level", "info")
            self._append_log(event["msg"], level)

        elif etype == "alias_needed":
            self._show_alias_dialog(event["canonical"], event["columns"])

        elif etype == "done":
            self._on_etl_done(event["results"])

        elif etype == "error":
            self._append_log(f"ERROR:  {event['msg']}", "error")
            if event.get("tb"):
                self._append_log(event["tb"], "error")
            self._on_etl_error(event["msg"])

    def _append_log(self, msg: str, level: str = "info"):
        self._log_text.config(state="normal")
        tag = level if level in ("error", "warn") else "info"
        self._log_text.insert(tk.END, msg + "\n", tag)
        self._log_text.see(tk.END)
        self._log_text.config(state="disabled")

    def _show_alias_dialog(self, canonical: str, columns: List[str]):
        _FieldAliasDialog(self, canonical, columns, self._resp_q)

    # ── ETL completion handlers ────────────────────────────────────────────────

    def _on_etl_done(self, results: dict):
        self._etl_results = results

        # Save password if requested
        if _HAS_KEYRING and self._save_pw_var.get():
            try:
                _keyring.set_password(_KR_SERVICE, _KR_USER, self._pw_var.get())
            except Exception:
                pass

        # Persist settings
        input_path = Path(self._file_var.get()) if self._file_var.get().strip() else None
        last_input_dir = ""
        if results.get("is_batch"):
            if self._batch_files:
                last_input_dir = str(self._batch_files[0].parent.resolve())
            elif input_path is not None:
                last_input_dir = str(input_path.parent.resolve())
        elif input_path is not None:
            last_input_dir = str(input_path.parent.resolve())

        save_settings(
            {
                "last_input_dir":    last_input_dir,
                "last_output_dir":   results["output_dir"],
                "last_sheet_name":   results["sheet_name"],
                "last_main_sheet_name": results["sheet_name"],
                "last_vacant_sheet_name": results.get("vacant_sheet_name", "Vacant Slots"),
                "last_exit_sheet_name": results.get("exit_sheet_name", "Exit Log"),
                "last_services_sheet_name": results.get("services_sheet_name", "Services Log"),
            },
            self._script_dir,
        )

        self._show_results(results)
        self._show_panel(self._results_frame)

    def _on_etl_error(self, msg: str):
        self._etl_results = {"success": False, "error": msg}
        self._show_results(self._etl_results)
        self._show_panel(self._results_frame)

    # ── Results panel population ───────────────────────────────────────────────

    def _show_results(self, results: dict):
        success = results.get("success", False)

        if success:
            self._status_lbl.config(text="✓   Run complete",
                                    foreground=GREEN, font=F_HEADER)
            self._open_folder_btn.config(state="normal")
        else:
            self._status_lbl.config(text="✗   Run failed",
                                    foreground=RED, font=F_HEADER)
            self._open_folder_btn.config(state="disabled")

        t = self._sum_text
        t.config(state="normal")
        t.delete("1.0", tk.END)

        if success:
            def row(label: str, value: str, warn: bool = False):
                t.insert(tk.END, f"  {label:<26}", "lbl")
                t.insert(tk.END, value + "\n", "warn" if warn else "val")

            if results.get("is_batch"):
                t.insert(tk.END, "\n")
                row("Mode", "Batch")
                row("Files processed", str(results.get("batch_total", 0)))
                row("Succeeded", str(results.get("batch_success", 0)))
                row("Failed", str(results.get("batch_failed", 0)), warn=results.get("batch_failed", 0) > 0)
                row("Participant sheet", results.get("sheet_name", ""))
                row("Vacant slots sheet", results.get("vacant_sheet_name", ""))
                row("Exit log sheet", results.get("exit_sheet_name", ""))
                row("Services log sheet", results.get("services_sheet_name", ""))
                skipped = results.get("skipped_duplicates", [])
                if skipped:
                    row("Skipped duplicates", str(len(skipped)), warn=True)
                svc_skipped = results.get("skipped_services_duplicates", [])
                if svc_skipped:
                    row("Skipped services duplicates", str(len(svc_skipped)), warn=True)
                src = results.get("batch_dashboard_source", "")
                if src:
                    row("Dashboard source file", src)
                else:
                    row("Dashboard source file", "None published", warn=True)
                row("Master file rows", str(results.get("master_total", 0)))

                t.insert(tk.END, "\n  Per-file outcomes\n", "lbl")
                for r in results.get("batch_results", []):
                    status = "OK" if r.get("success") else "FAILED"
                    warn = not r.get("success")
                    snap = r.get("snapshot_date", "")
                    name = r.get("source_file_name", "")
                    row("  " + name, f"{snap}  [{status}]", warn=warn)
                failed_files = results.get("batch_failed_files", [])
                if failed_files:
                    t.insert(tk.END, "\n  Failure reasons\n", "lbl")
                    for item in failed_files:
                        name = item.get("source_file_name", "")
                        snap = item.get("snapshot_date", "")
                        err = item.get("error", "Unspecified error")
                        t.insert(tk.END, f"  {name} ({snap})\n", "warn")
                        t.insert(tk.END, f"    {err}\n", "error")
                if skipped:
                    t.insert(tk.END, "\n  Skipped as already integrated\n", "lbl")
                    for item in skipped:
                        row(
                            "  " + item.get("source_file_name", ""),
                            item.get("snapshot_date", ""),
                            warn=True,
                        )
            else:
                t.insert(tk.END, "\n")
                row("Input file",     results.get("source_file_name", ""))
                row("Participant sheet", results.get("sheet_name", ""))
                row("Vacant slots sheet", results.get("vacant_sheet_name", ""))
                row("Exit log sheet", results.get("exit_sheet_name", ""))
                row("Services log sheet", results.get("services_sheet_name", ""))
                row("Snapshot date",  results.get("snapshot_date", ""))

                dc = results.get("dashboard_count", 0)
                vc = results.get("vacant_slot_count", 0)
                if results.get("dashboard_published", True):
                    row("Dashboard rows", f"{dc}  ({vc} vacant slot rows)")
                else:
                    row("Dashboard rows", "Not published this run", warn=True)

                row("Exit log records",    str(results.get("exit_log_count", 0)))
                row(
                    "Services log records",
                    f"{results.get('services_log_count', 0)} "
                    f"(history {results.get('services_log_history_count', 0)}, "
                    f"dropped {results.get('services_log_dupes_dropped', 0)})",
                )
                row("Snapshot history",    f"{results.get('history_count', 0)} total rows")

                mt  = results.get("master_total", 0)
                ex  = results.get("exited_count", 0)
                un  = results.get("unknown_count", 0)
                row("Master file rows",
                    f"{mt}  ({ex} exited, {un} unknown)")

                dup = results.get("dup_flag_count", 0)
                row("Duplicate warnings",
                    str(dup) if dup == 0 else f"{dup}  ⚠",
                    warn=(dup > 0))

                missing = results.get("missing_req", [])
                if missing:
                    row("Missing required",
                        ", ".join(missing), warn=True)

            t.insert(tk.END, "\n")
            t.insert(tk.END, f"  {'Output folder':<26}", "lbl")

            out_dir = results.get("output_dir", "")
            t.insert(tk.END, out_dir, "link")
            t.tag_unbind("link", "<Button-1>")
            t.tag_bind("link", "<Button-1>",
                       lambda _: self._on_open_folder())
            t.tag_bind("link", "<Enter>",
                       lambda _: t.config(cursor="hand2"))
            t.tag_bind("link", "<Leave>",
                       lambda _: t.config(cursor="arrow"))
            t.insert(tk.END, "\n")

        else:
            t.insert(tk.END, "\nRun failed.\n\n", "error")
            if results.get("is_batch"):
                t.insert(
                    tk.END,
                    f"Batch files: {results.get('batch_total', 0)}  "
                    f"({results.get('batch_success', 0)} succeeded, "
                    f"{results.get('batch_failed', 0)} failed)\n\n",
                    "error",
                )
                failed_files = results.get("batch_failed_files", [])
                if failed_files:
                    t.insert(tk.END, "Failed files\n", "lbl")
                    for item in failed_files:
                        name = item.get("source_file_name", "")
                        snap = item.get("snapshot_date", "")
                        err = item.get("error", "Unspecified error")
                        t.insert(tk.END, f"  {name} ({snap})\n", "warn")
                        t.insert(tk.END, f"    {err}\n", "error")
                else:
                    t.insert(tk.END, results.get("error", "Unknown error"), "error")
                    t.insert(tk.END, "\n")
            else:
                t.insert(tk.END, results.get("error", "Unknown error"), "error")
                t.insert(tk.END, "\n")

        t.config(state="disabled")

    # ── Button handlers ────────────────────────────────────────────────────────

    def _on_open_folder(self):
        out = (
            self._etl_results.get("output_dir")
            or (str(self._current_output_dir.resolve())
                if self._current_output_dir else "")
        )
        if out and Path(out).is_dir():
            try:
                os.startfile(out)
            except Exception as exc:
                messagebox.showerror("Error", f"Could not open folder:\n{exc}")

    def _on_run_another(self):
        # Preserve: password, save-pw flag, output folder, sheet name
        saved_pw      = self._pw_var.get()
        saved_pw_flag = self._save_pw_var.get()
        saved_folder  = self._folder_var.get()
        saved_sheet   = self._sheet_var.get()
        saved_vacant_sheet = self._vacant_sheet_var.get()
        saved_exit_sheet = self._exit_sheet_var.get()
        saved_services_sheet = self._services_sheet_var.get()
        saved_publish = self._publish_dashboard_var.get()

        # Clear file
        self._file_var.set("")
        self._batch_files = []
        self._batch_info_lbl.config(text="")
        self._set_date_mode(False)

        # Reset date
        if not _HAS_DATEPICKER:
            self._date_entry.delete(0, tk.END)
            self._date_entry.insert(0, "YYYY-MM-DD")
            self._date_entry.configure(foreground=FG_MUTED)

        # Clear error labels
        self._file_err.config(text="")
        self._date_err.config(text="")
        self._folder_err.config(text="")

        # Restore preserved values
        self._pw_var.set(saved_pw)
        self._save_pw_var.set(saved_pw_flag)
        self._folder_var.set(saved_folder)
        self._sheet_var.set(saved_sheet)
        self._vacant_sheet_var.set(saved_vacant_sheet)
        self._exit_sheet_var.set(saved_exit_sheet)
        self._services_sheet_var.set(saved_services_sheet)
        self._publish_dashboard_var.set(saved_publish)

        # Reset Run panel state
        self._progress["value"] = 0
        self._step_lbl.config(text="Initialising…")
        self._log_text.config(state="normal")
        self._log_text.delete("1.0", tk.END)
        self._log_text.config(state="disabled")

        self._etl_results = {}
        self._validate()
        self._show_panel(self._setup_frame)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = ICMSApp()
    app.mainloop()
