"""
icms_gui.py
~~~~~~~~~~~
Tkinter GUI wrapper for the ICMS weekly snapshot ETL pipeline.

All import logic lives in icms_etl_core.run_snapshot_files; field and sheet
mappings live in icms_catalogue.py. The GUI only collects choices, shows
progress, and summarizes results.

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
from typing import Dict, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import icms_etl_core as core
from update_icms_snapshot_files import load_settings, save_settings
from version import __version__

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

_TOTAL_STEPS = len(core.STEP_LABELS)


def _date_note(info: Dict[str, Optional[str]]) -> str:
    """One-line description of where a default snapshot date came from."""
    if info.get("mismatch"):
        return (f"Workbook Generated On {info['generated_on']}; "
                f"filename says {info['filename_date']}. Using Generated On.")
    if info.get("source") == "Generated On":
        return "From workbook Generated On"
    if info.get("source") == "filename":
        return "From filename (Generated On not readable)"
    return "No date found; enter one"


# ── Already-imported dialog ───────────────────────────────────────────────────

class _ReprocessDialog(tk.Toplevel):
    """One choice for every already-imported file in the selection."""

    def __init__(self, parent: tk.Tk, duplicates: List[core.FileSpec], total: int):
        super().__init__(parent)
        self.title("Files Already Imported")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self._duplicates = duplicates
        self._total = total
        self.action: str = "cancel"  # reprocess | skip | cancel
        self._build()
        self._center(parent)

    def _build(self):
        outer = ttk.Frame(self, padding=PAD)
        outer.pack(fill="both", expand=True)

        n = len(self._duplicates)
        ttk.Label(
            outer,
            text=(
                f"{n} of {self._total} selected file(s) were already imported into\n"
                "this output folder (same file name and snapshot date).\n\n"
                "Reprocessing replaces their stored rows with freshly read values,\n"
                "which repairs earlier imports. Outputs are backed up first."
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
            height=min(10, max(3, n)),
            bg=INPUT_BG,
            fg=FG,
            relief="flat",
            borderwidth=1,
        )
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for f in self._duplicates:
            lb.insert(tk.END, f"{f.snapshot_date}  |  {Path(f.path).name}")

        btn = ttk.Frame(outer)
        btn.pack(fill="x", pady=(12, 0))
        ttk.Button(
            btn,
            text="Reprocess all selected",
            style="Accent.TButton",
            command=self._on_reprocess,
        ).pack(side="left")
        if n < self._total:
            ttk.Button(
                btn,
                text="Skip already-imported",
                command=self._on_skip,
            ).pack(side="left", padx=(8, 0))
        ttk.Button(
            btn,
            text="Cancel",
            command=self._on_cancel,
        ).pack(side="right")

    def _center(self, parent: tk.Tk):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_reqwidth()
        dh = self.winfo_reqheight()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _on_reprocess(self):
        self.action = "reprocess"
        self.destroy()

    def _on_skip(self):
        self.action = "skip"
        self.destroy()

    def _on_cancel(self):
        self.action = "cancel"
        self.destroy()


# ── Main application ──────────────────────────────────────────────────────────

class ICMSApp(tk.Tk):
    """Single-window Tkinter application for the ICMS snapshot ETL pipeline."""

    def __init__(self):
        super().__init__()

        self.title(f"ICMS Snapshot ETL v{__version__}")
        self.geometry("560x820")
        self.minsize(520, 640)
        self.resizable(False, True)
        self.configure(bg=BG)

        self._script_dir: Path = Path(__file__).parent.resolve()
        self._gui_q: queue.Queue = queue.Queue()
        self._etl_results: dict = {}
        self._current_output_dir: Optional[Path] = None
        self._batch_files: List[Path] = []
        self._batch_date_rows: List[dict] = []
        self._single_date_info: Optional[dict] = None
        self._skipped_already_imported: List[core.FileSpec] = []
        self._last_input_dir: str = ""

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

        self._date_info_lbl = ttk.Label(
            self._date_row, text="", foreground=FG_MUTED, font=F_SMALL,
            background=BG, wraplength=330, justify="left",
        )
        self._date_info_lbl.pack(side="left", padx=(10, 0))

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
        ttk.Label(
            body,
            text=("Known vendor names (e.g. Service Log / Services Log) are "
                  "recognized automatically."),
            foreground=FG_MUTED, font=F_SMALL, background=BG,
        ).pack(anchor="w")

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
            text="Publish as current dashboard flat file",
            variable=self._publish_dashboard_var,
        ).pack(anchor="w", pady=(3, 0))
        ttk.Label(
            body,
            text=("In a batch, only the newest file is published. "
                  "Turn this off to leave the current dashboard as is."),
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

        self._copy_btn = ttk.Button(
            btn_row, text="Copy findings for director",
            command=self._on_copy_findings, state="disabled")
        self._copy_btn.pack(side="left", padx=(6, 0))

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

        # Remembered input-file browse location (folder only; no file to
        # prefill until the user picks one).
        saved_input_dir = settings.get("last_input_dir", "")
        if saved_input_dir and Path(saved_input_dir).is_dir():
            self._last_input_dir = saved_input_dir

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

    def _update_settings(self, **kwargs) -> None:
        """Merge *kwargs* into settings.json without clobbering other keys."""
        settings = load_settings(self._script_dir)
        settings.update(kwargs)
        save_settings(settings, self._script_dir)

    def _remember_input_dir(self, path: str) -> None:
        """Record *path*'s parent folder as the last-used input directory."""
        self._last_input_dir = str(Path(path).parent.resolve())
        self._update_settings(last_input_dir=self._last_input_dir)

    def _on_browse_file(self):
        self._batch_files = []
        self._batch_info_lbl.config(text="")
        self._set_date_mode(False)
        cur = self._file_var.get()
        init = str(Path(cur).parent) if cur else (self._last_input_dir or str(Path.home()))
        path = filedialog.askopenfilename(
            title="Select ICMS Snapshot",
            initialdir=init,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._file_var.set(path)
            self._auto_fill_from_file(path)
            self._remember_input_dir(path)

    def _on_browse_files(self):
        cur = self._file_var.get()
        init = str(Path(cur).parent) if cur else (self._last_input_dir or str(Path.home()))
        paths = filedialog.askopenfilenames(
            title="Select Multiple ICMS Snapshots",
            initialdir=init,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not paths:
            return
        self._remember_input_dir(paths[0])
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
            info = self._read_date_info(path)
            container = ttk.Frame(self._batch_date_list)
            container.pack(fill="x", pady=(0, 6))
            row = ttk.Frame(container)
            row.pack(fill="x")
            ttk.Label(
                container, text=_date_note(info), font=F_SMALL, background=BG,
                foreground=AMBER if info.get("mismatch") or not info.get("date") else FG_MUTED,
            ).pack(anchor="w")

            file_var = tk.StringVar(value=path.name)
            date_var = tk.StringVar(value=info.get("date") or "")
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
                    "container": container,
                    "info": info,
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

    def _spec_for(self, path: Path, snap: str, info: Optional[dict]) -> core.FileSpec:
        info = info or {}
        if snap and snap == info.get("date"):
            source = info.get("source") or ""
        else:
            source = "manual entry"
        return core.FileSpec(
            path=path,
            snapshot_date=snap,
            date_source=source,
            generated_on=info.get("generated_on"),
            filename_date=info.get("filename_date"),
        )

    def _collect_batch_specs(self) -> List[core.FileSpec]:
        return [
            self._spec_for(row["path"], row["date_var"].get().strip(), row.get("info"))
            for row in self._batch_date_rows
        ]

    def _read_date_info(self, path) -> dict:
        """Read Generated On (opens the workbook, so show a busy cursor)."""
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            return core.default_snapshot_date(path, self._pw_var.get())
        finally:
            self.config(cursor="")

    def _auto_fill_from_file(self, filepath: str):
        info = self._read_date_info(filepath)
        self._single_date_info = info
        if info.get("date"):
            self._set_date(info["date"])
        elif not _HAS_DATEPICKER:
            self._date_entry.focus_set()
        self._date_info_lbl.config(
            text=_date_note(info),
            foreground=AMBER if info.get("mismatch") or not info.get("date") else FG_MUTED,
        )

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
            self._update_settings(last_output_dir=path)

    def _sheet_names(self) -> Dict[str, str]:
        return {
            "participant": self._sheet_var.get().strip(),
            "vacant": self._vacant_sheet_var.get().strip(),
            "exit": self._exit_sheet_var.get().strip(),
            "services": self._services_sheet_var.get().strip(),
        }

    # ── Run ────────────────────────────────────────────────────────────────────

    def _on_run(self):
        if not self._validate():
            return

        if len(self._batch_files) > 1:
            specs = self._collect_batch_specs()
        else:
            fp = self._file_var.get().strip()
            if not fp:
                messagebox.showerror("Error", "Select an input file first.")
                return
            specs = [self._spec_for(Path(fp), self._get_date_str(), self._single_date_info)]

        output_dir = Path(self._folder_var.get())
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Error", f"Could not create output directory:\n{exc}")
            return

        self._skipped_already_imported = []
        reprocessing = False
        dups = core.already_imported(specs, output_dir)
        if dups:
            dialog = _ReprocessDialog(self, dups, len(specs))
            self.wait_window(dialog)
            if dialog.action == "cancel":
                return
            if dialog.action == "skip":
                dup_ids = {id(d) for d in dups}
                self._skipped_already_imported = dups
                specs = [s for s in specs if id(s) not in dup_ids]
            else:
                reprocessing = True

        while True:
            locked = core.check_outputs_writable(output_dir)
            if not locked:
                break
            names = "\n".join(
                f"  • {p.name if p.suffix else p}" for p in locked
            )
            retry = messagebox.askretrycancel(
                "Output files in use",
                "Cannot write to one or more output files. Close them in Excel "
                "or another program, then click Retry.\n\n"
                f"{names}",
            )
            if not retry:
                return

        self._current_output_dir = output_dir
        self._progress["value"] = 0
        self._step_lbl.config(text="Starting…")
        self._log_text.config(state="normal")
        self._log_text.delete("1.0", tk.END)
        self._log_text.config(state="disabled")
        self._show_panel(self._run_frame)

        threading.Thread(
            target=self._etl_worker,
            args=(
                specs,
                self._sheet_names(),
                self._publish_dashboard_var.get(),
                self._pw_var.get(),
                output_dir,
                reprocessing,
            ),
            daemon=True,
        ).start()

    # ── ETL worker (background thread) ────────────────────────────────────────

    def _etl_worker(
        self,
        specs: List[core.FileSpec],
        sheet_names: Dict[str, str],
        publish_dashboard: bool,
        password: str,
        output_dir: Path,
        reprocessing: bool,
    ):
        gq = self._gui_q

        def log(msg: str, level: str = "info"):
            ts = dt.datetime.now().strftime("%H:%M:%S")
            gq.put({"type": "log", "msg": f"[{ts}]  {msg}", "level": level})

        def step(n: int):
            gq.put({"type": "step", "n": n, "label": core.STEP_LABELS[n - 1] + "…"})

        try:
            results = core.run_snapshot_files(
                specs,
                output_dir,
                password=password,
                sheet_names=sheet_names,
                publish_dashboard=publish_dashboard,
                reprocessing=reprocessing,
                log=log,
                step=step,
            )
            results["skipped_already_imported"] = [
                {"source_file_name": Path(s.path).name, "snapshot_date": s.snapshot_date}
                for s in self._skipped_already_imported
            ]
            gq.put({"type": "done", "results": results})
        except Exception as exc:
            gq.put({
                "type": "error",
                "msg": str(exc).strip() or exc.__class__.__name__,
                "tb": traceback.format_exc(),
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

    # ── ETL completion handlers ────────────────────────────────────────────────

    def _on_etl_done(self, results: dict):
        self._etl_results = results

        # Save password if requested
        if _HAS_KEYRING and self._save_pw_var.get():
            try:
                _keyring.set_password(_KR_SERVICE, _KR_USER, self._pw_var.get())
            except Exception:
                pass

        input_path = Path(self._file_var.get()) if self._file_var.get().strip() else None
        last_input_dir = ""
        if self._batch_files:
            last_input_dir = str(self._batch_files[0].parent.resolve())
        elif input_path is not None:
            last_input_dir = str(input_path.parent.resolve())

        sheets = self._sheet_names()
        self._update_settings(
            last_input_dir=last_input_dir,
            last_output_dir=results["output_dir"],
            last_sheet_name=sheets["participant"],
            last_main_sheet_name=sheets["participant"],
            last_vacant_sheet_name=sheets["vacant"],
            last_exit_sheet_name=sheets["exit"],
            last_services_sheet_name=sheets["services"],
        )

        self._show_results(results)
        self._show_panel(self._results_frame)

    def _on_etl_error(self, msg: str):
        self._etl_results = {"success": False, "error": msg}
        self._show_results(self._etl_results)
        self._show_panel(self._results_frame)

    def _on_copy_findings(self):
        text = self._etl_results.get("findings_text", "")
        if not text:
            return
        header = (f"ICMS Snapshot ETL v{__version__} structural findings "
                  f"({dt.date.today().isoformat()})\n\n")
        self.clipboard_clear()
        self.clipboard_append(header + text)
        self._copy_btn.config(text="Copied ✓")
        self.after(2000, lambda: self._copy_btn.config(text="Copy findings for director"))

    # ── Results panel population ───────────────────────────────────────────────

    def _show_results(self, results: dict):
        success = results.get("success", False)
        incomplete = bool(results.get("skipped_categories")) or results.get("batch_failed", 0) > 0

        if success and incomplete:
            self._status_lbl.config(text="⚠   Run complete with gaps",
                                    foreground=AMBER, font=F_HEADER)
        elif success:
            self._status_lbl.config(text="✓   Run complete",
                                    foreground=GREEN, font=F_HEADER)
        else:
            self._status_lbl.config(text="✗   Run failed",
                                    foreground=RED, font=F_HEADER)
        self._open_folder_btn.config(state="normal" if success else "disabled")
        self._copy_btn.config(state="normal" if results.get("findings_text") else "disabled")

        t = self._sum_text
        t.config(state="normal")
        t.delete("1.0", tk.END)

        def row(label: str, value: str, warn: bool = False):
            t.insert(tk.END, f"  {label:<26}", "lbl")
            t.insert(tk.END, value + "\n", "warn" if warn else "val")

        per_file = results.get("batch_results", [])
        if not per_file and not success:
            t.insert(tk.END, "\nRun failed.\n\n", "error")
            t.insert(tk.END, results.get("error", "Unknown error") + "\n", "error")
            t.config(state="disabled")
            return

        t.insert(tk.END, "\n")
        if results.get("is_batch"):
            row("Mode", "Batch")
            row("Files processed", str(results.get("batch_total", 0)))
            row("Succeeded", str(results.get("batch_success", 0)))
            row("Failed", str(results.get("batch_failed", 0)),
                warn=results.get("batch_failed", 0) > 0)
        elif per_file:
            r = per_file[0]
            row("Input file", r.get("source_file_name", ""))
            row("Snapshot date", f"{r.get('snapshot_date', '')}"
                + (f"  ({r['snapshot_date_source']})" if r.get("snapshot_date_source") else ""))
            if r.get("success"):
                row("Participant rows", str(r.get("participant_count", 0)))
                row("Exit log records", str(r.get("exit_log_count", 0)))
                row("Services log records",
                    f"{r.get('services_log_count', 0)} "
                    f"(history {r.get('services_log_history_count', 0)})")
                row("Snapshot history", f"{r.get('history_count', 0)} total rows")
                dup = r.get("dup_flag_count", 0)
                row("Duplicate enrollment rows", str(dup) if dup == 0 else f"{dup}  ⚠", warn=dup > 0)

        if results.get("dashboard_published"):
            row("Dashboard", f"{results.get('dashboard_count', 0)} rows from "
                f"{results.get('dashboard_source', '')}")
        else:
            row("Dashboard", results.get("dashboard_note") or "Not updated", warn=True)
        row("Master enrollments",
            f"{results.get('master_total', 0)}  ({results.get('master_exited_count', 0)} exited, "
            f"{results.get('master_unknown_count', 0)} unknown)")
        if results.get("master_error"):
            row("Master rebuild", results["master_error"], warn=True)
        if results.get("backup_dir"):
            row("Backup of prior outputs", results["backup_dir"])

        skipped_tabs = results.get("skipped_categories", [])
        if skipped_tabs:
            row("Not imported", ", ".join(skipped_tabs) + " (see findings)", warn=True)

        if results.get("is_batch"):
            t.insert(tk.END, "\n  Per-file outcomes\n", "lbl")
            for r in per_file:
                status = "OK" if r.get("success") else "FAILED"
                extra = ""
                if r.get("success") and r.get("skipped_categories"):
                    extra = "  (missing: " + ", ".join(r["skipped_categories"]) + ")"
                row("  " + r.get("snapshot_date", ""),
                    f"{r.get('source_file_name', '')}  [{status}]{extra}",
                    warn=not r.get("success") or bool(extra))

        failed_files = results.get("batch_failed_files", [])
        if failed_files:
            t.insert(tk.END, "\n  Not imported (existing outputs unchanged)\n", "lbl")
            for item in failed_files:
                t.insert(tk.END, f"  {item.get('source_file_name', '')} "
                                 f"({item.get('snapshot_date', '')})\n", "warn")
                t.insert(tk.END, f"    {item.get('error') or 'Unspecified error'}\n", "error")

        skipped = results.get("skipped_already_imported", [])
        if skipped:
            t.insert(tk.END, "\n  Skipped as already imported\n", "lbl")
            for item in skipped:
                row("  " + item.get("snapshot_date", ""), item.get("source_file_name", ""), warn=True)

        findings = results.get("findings_text", "")
        t.insert(tk.END, "\n  Structural findings\n", "lbl")
        if findings:
            t.insert(tk.END, "  Nothing needs to be done during this run. Use “Copy findings for "
                             "director” to pass these on.\n", "lbl")
            for line in findings.splitlines():
                t.insert(tk.END, f"  {line}\n", "warn")
        else:
            t.insert(tk.END, "  None — every column and tab was recognized.\n", "val")

        if success:
            t.insert(tk.END, "\n")
            t.insert(tk.END, f"  {'Output folder':<26}", "lbl")
            t.insert(tk.END, results.get("output_dir", ""), "link")
            t.tag_unbind("link", "<Button-1>")
            t.tag_bind("link", "<Button-1>", lambda _: self._on_open_folder())
            t.tag_bind("link", "<Enter>", lambda _: t.config(cursor="hand2"))
            t.tag_bind("link", "<Leave>", lambda _: t.config(cursor="arrow"))
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
        self._single_date_info = None
        self._date_info_lbl.config(text="")

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
