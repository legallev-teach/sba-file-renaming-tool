"""
SBA File Renamer — Desktop (Tkinter)
Packages to .exe via:  pyinstaller --onefile --windowed desktop_app.py
"""

import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from concurrent.futures import ThreadPoolExecutor

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.engine import (
    CSEC_SUBJECTS, CSEC_SUBJECT_NAMES, FILE_TYPES, DOC_NUMBERS,
    extract_title_page_text, extract_candidate_info,
    load_master_list, process_job, RenameJob,
    validate_candidate_number, validate_moderation_code, build_filename,
    check_tesseract_available, name_matches, normalise_cxc_name,
    verify_candidate_against_master_list,
)

ACCENT   = "#1a5276"
BG       = "#f4f6f7"
WHITE    = "#ffffff"
SUCCESS  = "#1e8449"
ERROR    = "#922b21"
DISABLED = "#abb2b9"
AMBER    = "#b7770d"
BLUE     = "#1a5276"

# Per-file status constants
ST_SCANNING  = "scanning"   # ⟳ background scan in progress
ST_FOUND     = "found"      # ✓ number found by scan
ST_MANUAL    = "manual"     # ✎ user typed it in
ST_MISSING   = "missing"    # ? no number found, needs input
ST_SKIPPED   = "skipped"    # ↷ user skipped this file
ST_ERROR     = "error"      # ✗ scan failed with an error

STATUS_ICON = {
    ST_SCANNING: "⟳",
    ST_FOUND:    "✓",
    ST_MANUAL:   "✎",
    ST_MISSING:  "?",
    ST_SKIPPED:  "↷",
    ST_ERROR:    "✗",
}

STATUS_COLOR = {
    ST_SCANNING: DISABLED,
    ST_FOUND:    SUCCESS,
    ST_MANUAL:   BLUE,
    ST_MISSING:  AMBER,
    ST_SKIPPED:  DISABLED,
    ST_ERROR:    ERROR,
}


class FileEntry:
    """Holds per-file state."""
    def __init__(self, path: Path):
        self.path        = path
        self.cand_num    = ""
        self.cand_name   = ""   # name extracted from PDF cover page (may be empty)
        self.status      = ST_SCANNING
        self.scan_error: str = ""   # human-readable error from last scan attempt


class SBARenamerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SBA File Renamer")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(680, 620)

        # ── State ──
        self.entries: list[FileEntry] = []   # one per loaded file
        self.current_idx: int | None = None
        self.output_dir: Path | None = None
        self.master_list: dict[str, str] = {}
        self._suppress_cand_trace = False    # prevent feedback loops

        # ── Thread safety ──
        self._scan_queue: queue.Queue = queue.Queue()
        self._executor = ThreadPoolExecutor(max_workers=4, thread_name_prefix="sba_scan")
        self._generation = 0

        # ── OCR availability ──
        self._ocr_available = check_tesseract_available()

        self._build_ui()
        self._poll_scan_queue()

        if not self._ocr_available:
            self.status_var.set(
                "Note: Tesseract OCR not found — text-based PDF extraction only."
            )

    # ─────────────────────────────────────────────────────────────────────────
    # UI Construction
    # ─────────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = dict(padx=12, pady=6)

        # ── Header ──
        header = tk.Frame(self, bg=ACCENT)
        header.pack(fill="x")
        tk.Label(header, text="SBA File Renamer", bg=ACCENT, fg=WHITE,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=16, pady=10)
        tk.Button(header, text="ℹ", bg=ACCENT, fg="#aed6f1",
                  font=("Segoe UI", 11), relief="flat", bd=0,
                  command=self._show_about,
                  cursor="hand2").pack(side="right", padx=16)

        # ── Action buttons ──
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(side="bottom", fill="x", padx=16, pady=(0, 14))

        self.scan_btn = tk.Button(btn_frame, text="⟳ Scan Current File",
                                  command=self._scan_current,
                                  bg="#5d6d7e", fg=WHITE, relief="flat", padx=12, pady=6,
                                  font=("Segoe UI", 9))
        self.scan_btn.pack(side="left", padx=(0, 8))

        self.rename_btn = tk.Button(btn_frame, text="Rename & Copy All →",
                                    command=self._rename_all,
                                    bg=SUCCESS, fg=WHITE, relief="flat", padx=16, pady=6,
                                    font=("Segoe UI", 10, "bold"))
        self.rename_btn.pack(side="right")

        # ── Scrollable canvas wrapper ──
        canvas_frame = tk.Frame(self, bg=BG)
        canvas_frame.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(canvas_frame, bg=BG, highlightthickness=0)
        v_scroll = tk.Scrollbar(canvas_frame, orient="vertical",
                                command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        main = tk.Frame(self._canvas, bg=BG)
        self._canvas_window = self._canvas.create_window(
            (0, 0), window=main, anchor="nw")

        def _on_main_configure(event):
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))

        def _on_canvas_configure(event):
            self._canvas.itemconfig(self._canvas_window, width=event.width)

        main.bind("<Configure>", _on_main_configure)
        self._canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_mousewheel_linux(event):
            direction = -1 if event.num == 4 else 1
            self._canvas.yview_scroll(direction, "units")

        self._canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self._canvas.bind_all("<Button-4>", _on_mousewheel_linux)
        self._canvas.bind_all("<Button-5>", _on_mousewheel_linux)

        main = tk.Frame(main, bg=BG)
        main.pack(fill="both", expand=True, padx=16, pady=8)

        # ── Section 1: File list ──
        self._section(main, "1. Select SBA Files")
        fr_files = tk.Frame(main, bg=BG)
        fr_files.pack(fill="x", padx=12, pady=(4, 0))

        btn_row = tk.Frame(fr_files, bg=BG)
        btn_row.pack(fill="x")
        tk.Button(btn_row, text="Add Files…", command=self._browse_files,
                  bg=ACCENT, fg=WHITE, relief="flat", padx=10).pack(side="left")
        tk.Button(btn_row, text="Remove Selected", command=self._remove_selected,
                  bg="#7f8c8d", fg=WHITE, relief="flat", padx=10).pack(side="left", padx=6)
        tk.Button(btn_row, text="Clear All", command=self._clear_all,
                  bg=ERROR, fg=WHITE, relief="flat", padx=10).pack(side="left")

        list_frame = tk.Frame(fr_files, bg=BG)
        list_frame.pack(fill="x", pady=(6, 0))

        self.file_listbox = tk.Listbox(
            list_frame, height=4, font=("Consolas", 9), selectmode="browse",
            bg=WHITE, fg="#1a252f", selectbackground=ACCENT, selectforeground=WHITE,
            activestyle="none", relief="solid", bd=1,
        )
        scrollbar = tk.Scrollbar(list_frame, orient="vertical",
                                 command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side="left", fill="x", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        self.list_status_lbl = tk.Label(fr_files, text="No files loaded.",
                                        bg=BG, fg=DISABLED, font=("Segoe UI", 9))
        self.list_status_lbl.pack(anchor="w", pady=(2, 0))

        # ── Section 2: Output folder ──
        self._section(main, "2. Output Folder")
        fr_out = tk.Frame(main, bg=BG)
        fr_out.pack(fill="x", **pad)
        self.out_label = tk.Label(fr_out, text="Same folder as source files (default)",
                                  bg=BG, fg=DISABLED, font=("Segoe UI", 9), anchor="w")
        self.out_label.pack(side="left", fill="x", expand=True)
        tk.Button(fr_out, text="Browse…", command=self._browse_output,
                  bg=ACCENT, fg=WHITE, relief="flat", padx=10).pack(side="right")

        # ── Section 3: Master list ──
        self._section(main, "3. Master List (optional — candidate number lookup)")
        fr_ml = tk.Frame(main, bg=BG)
        fr_ml.pack(fill="x", **pad)
        self.ml_label = tk.Label(fr_ml, text="No master list loaded",
                                 bg=BG, fg=DISABLED, font=("Segoe UI", 9), anchor="w")
        self.ml_label.pack(side="left", fill="x", expand=True)
        tk.Button(fr_ml, text="Load…", command=self._load_master_list,
                  bg="#5d6d7e", fg=WHITE, relief="flat", padx=10).pack(side="right")

        # ── Section 4: Batch settings ──
        self._section(main, "4. Batch Settings  (apply to all files)")
        grid = tk.Frame(main, bg=BG)
        grid.pack(fill="x", padx=12, pady=4)

        tk.Label(grid, text="Subject:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=3)
        self.subj_var = tk.StringVar()
        self.subj_combo = ttk.Combobox(grid, textvariable=self.subj_var,
                                       values=CSEC_SUBJECT_NAMES, width=36, state="readonly")
        self.subj_combo.grid(row=0, column=1, columnspan=2, sticky="w", padx=8)
        self.subj_combo.bind("<<ComboboxSelected>>", lambda _: self._update_preview())

        tk.Label(grid, text="Moderation Code:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=3)
        self.mod_var = tk.StringVar()
        tk.Entry(grid, textvariable=self.mod_var, width=12,
                 font=("Consolas", 10), state="readonly").grid(
                     row=1, column=1, sticky="w", padx=8)

        tk.Label(grid, text="File Type:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=3)
        self.ftype_var = tk.StringVar(value="SBA")
        ftype_frame = tk.Frame(grid, bg=BG)
        ftype_frame.grid(row=2, column=1, columnspan=2, sticky="w", padx=8)
        for ft in FILE_TYPES:
            tk.Radiobutton(ftype_frame, text=ft, variable=self.ftype_var, value=ft,
                           bg=BG, font=("Segoe UI", 9),
                           command=self._update_preview).pack(side="left", padx=6)

        tk.Label(grid, text="Document #:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=3, column=0, sticky="w", pady=3)
        self.docnum_var = tk.IntVar(value=1)
        dn_frame = tk.Frame(grid, bg=BG)
        dn_frame.grid(row=3, column=1, columnspan=2, sticky="w", padx=8)
        for n in DOC_NUMBERS:
            tk.Radiobutton(dn_frame, text=str(n), variable=self.docnum_var, value=n,
                           bg=BG, font=("Segoe UI", 9),
                           command=self._update_preview).pack(side="left", padx=6)

        self.subj_var.trace_add("write", self._on_subject_change)

        # ── Section 5: Current file / candidate number ──
        self._section(main, "5. Candidate Details  (per file)")
        cand_frame = tk.Frame(main, bg=BG)
        cand_frame.pack(fill="x", padx=12, pady=4)

        tk.Label(cand_frame, text="Current file:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        self.current_file_lbl = tk.Label(cand_frame, text="—", bg=BG,
                                         fg="#5d6d7e", font=("Segoe UI", 9, "italic"),
                                         anchor="w")
        self.current_file_lbl.grid(row=0, column=1, columnspan=3, sticky="w", padx=8)

        tk.Label(cand_frame, text="Candidate Number:", bg=BG,
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=3)
        self.cand_var = tk.StringVar()
        self.cand_var.trace_add("write", self._on_cand_typed)
        self.cand_entry = tk.Entry(cand_frame, textvariable=self.cand_var,
                                   width=16, font=("Consolas", 10))
        self.cand_entry.grid(row=1, column=1, sticky="w", padx=8)

        self.cand_name_lbl = tk.Label(cand_frame, text="", bg=BG, fg="#5d6d7e",
                                      font=("Segoe UI", 9, "italic"))
        self.cand_name_lbl.grid(row=1, column=2, sticky="w")

        self.skip_btn = tk.Button(cand_frame, text="↷ Skip File",
                                  command=self._toggle_skip,
                                  bg="#7f8c8d", fg=WHITE, relief="flat", padx=8,
                                  font=("Segoe UI", 9))
        self.skip_btn.grid(row=1, column=3, padx=8)

        # ── Master list warning banner (shown when name doesn't match) ──
        # Full-width amber banner — visible without scrolling, impossible to miss.
        self.verify_warn_lbl = tk.Label(
            cand_frame, text="", bg="#fef9e7", fg="#7d6608",
            font=("Segoe UI", 9, "bold"), anchor="w", wraplength=520,
            relief="solid", bd=1, padx=8, pady=5,
        )
        self.verify_warn_lbl.grid(row=2, column=0, columnspan=4, sticky="ew",
                                  pady=(4, 2), padx=(0, 8))
        self.verify_warn_lbl.grid_remove()   # hidden until needed

        # ── Preview ──
        self._section(main, "Preview")
        self.preview_var = tk.StringVar(value="—")
        tk.Label(main, textvariable=self.preview_var, bg="#eaf2ff",
                 font=("Consolas", 10), anchor="w", relief="flat",
                 padx=10, pady=6).pack(fill="x", padx=12)

        # ── Status ──
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(main, textvariable=self.status_var, bg=BG, fg="#5d6d7e",
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", padx=12, pady=(4, 0))

    def _section(self, parent, text):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", padx=12, pady=(6, 0))
        tk.Label(f, text=text.upper(), bg=BG, fg=ACCENT,
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        tk.Frame(f, bg="#d5d8dc", height=1).pack(side="left", fill="x",
                                                   expand=True, padx=(8, 0))

    # ─────────────────────────────────────────────────────────────────────────
    # File list management
    # ─────────────────────────────────────────────────────────────────────────

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select SBA PDF file(s)",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not paths:
            return

        added = 0
        existing_paths = {e.path for e in self.entries}
        for p in paths:
            path = Path(p)
            if path.suffix.lower() != ".pdf":
                continue
            if path in existing_paths:
                continue
            entry = FileEntry(path)
            self.entries.append(entry)
            existing_paths.add(path)
            added += 1

        self._refresh_listbox()

        if added:
            new_entries = self.entries[-added:]
            for entry in new_entries:
                self._submit_scan(entry)
            new_start_idx = len(self.entries) - added
            self.file_listbox.selection_clear(0, "end")
            self.file_listbox.selection_set(new_start_idx)
            self.file_listbox.see(new_start_idx)
            self._load_file_into_ui(new_start_idx)

        self._update_list_status()

    def _remove_selected(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.entries.pop(idx)
        self.current_idx = None
        self._refresh_listbox()
        self._update_list_status()

        if self.entries:
            new_idx = min(idx, len(self.entries) - 1)
            self.file_listbox.selection_set(new_idx)
            self._load_file_into_ui(new_idx)
        else:
            self._clear_candidate_ui()

    def _clear_all(self):
        if self.entries and not messagebox.askyesno(
                "Clear All", "Remove all files and reset the form?"):
            return
        self._generation += 1
        self.entries.clear()
        self.current_idx = None
        self.output_dir = None
        self.master_list = {}
        self.out_label.config(text="Same folder as source files (default)", fg=DISABLED)
        self.ml_label.config(text="No master list loaded", fg=DISABLED)
        self.subj_var.set("")
        self.mod_var.set("")
        self.ftype_var.set("SBA")
        self.docnum_var.set(1)
        self._refresh_listbox()
        self._clear_candidate_ui()
        self._update_list_status()
        self.status_var.set("Cleared.")

    def _refresh_listbox(self):
        self.file_listbox.delete(0, "end")
        for entry in self.entries:
            icon = STATUS_ICON.get(entry.status, "?")
            display = f" {icon}  {entry.path.name}"
            self.file_listbox.insert("end", display)
            color = STATUS_COLOR.get(entry.status, DISABLED)
            self.file_listbox.itemconfig("end", fg=color)

    def _update_list_status(self):
        total = len(self.entries)
        if total == 0:
            self.list_status_lbl.config(text="No files loaded.", fg=DISABLED)
            return
        ready    = sum(1 for e in self.entries if e.status in (ST_FOUND, ST_MANUAL))
        missing  = sum(1 for e in self.entries if e.status == ST_MISSING)
        skipped  = sum(1 for e in self.entries if e.status == ST_SKIPPED)
        scanning = sum(1 for e in self.entries if e.status == ST_SCANNING)
        errors   = sum(1 for e in self.entries if e.status == ST_ERROR)
        parts = [f"{total} file(s)"]
        if ready:    parts.append(f"{ready} ready")
        if missing:  parts.append(f"{missing} need number")
        if skipped:  parts.append(f"{skipped} skipped")
        if scanning: parts.append(f"{scanning} scanning…")
        if errors:   parts.append(f"{errors} error(s)")
        self.list_status_lbl.config(text="  ·  ".join(parts), fg="#5d6d7e")

    # ─────────────────────────────────────────────────────────────────────────
    # File selection / navigation
    # ─────────────────────────────────────────────────────────────────────────

    def _on_file_select(self, event=None):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        self._load_file_into_ui(sel[0])

    def _load_file_into_ui(self, idx: int):
        if idx < 0 or idx >= len(self.entries):
            return
        self.current_idx = idx
        entry = self.entries[idx]

        self.current_file_lbl.config(text=entry.path.name)

        self._suppress_cand_trace = True
        self.cand_var.set(entry.cand_num)
        self._suppress_cand_trace = False

        self._update_cand_name_label(entry.cand_num)
        self._run_verify_display(entry)
        self._update_skip_btn(entry)
        self._update_preview()

        # Always update the status bar to reflect the newly selected file,
        # so stale messages from other files don't linger.
        if entry.status == ST_ERROR and entry.scan_error:
            self.status_var.set(f"Scan error — {entry.path.name}: {entry.scan_error}")
        elif entry.status == ST_MISSING:
            self.status_var.set(f"{entry.path.name}: no number found — enter manually.")
        elif entry.status == ST_FOUND or entry.status == ST_MANUAL:
            self.status_var.set(f"{entry.path.name}: candidate {entry.cand_num}")
        elif entry.status == ST_SKIPPED:
            self.status_var.set(f"{entry.path.name}: skipped.")
        else:
            self.status_var.set(f"{entry.path.name}: scanning…")

    def _clear_candidate_ui(self):
        self.current_file_lbl.config(text="—")
        self._suppress_cand_trace = True
        self.cand_var.set("")
        self._suppress_cand_trace = False
        self.cand_name_lbl.config(text="")
        self.verify_warn_lbl.config(text="")
        self.preview_var.set("—")
        self.skip_btn.config(text="↷ Skip File", bg="#7f8c8d")

    # ─────────────────────────────────────────────────────────────────────────
    # Candidate number handling
    # ─────────────────────────────────────────────────────────────────────────

    def _on_cand_typed(self, *_):
        if self._suppress_cand_trace:
            return
        if self.current_idx is None:
            return
        entry = self.entries[self.current_idx]
        val = self.cand_var.get().strip()
        entry.cand_num = val
        if entry.status != ST_SKIPPED:
            ok, _ = validate_candidate_number(val)
            if ok:
                entry.status = ST_MANUAL
            else:
                entry.status = ST_MISSING
        self._refresh_listbox()
        self._update_cand_name_label(val)
        self._run_verify_display(entry)
        self._update_list_status()
        self._update_preview()

    def _update_cand_name_label(self, cand: str):
        """Show the master list name next to the candidate number field."""
        if cand in self.master_list:
            self.cand_name_lbl.config(text=f"✓ {self.master_list[cand]}", fg=SUCCESS)
        elif len(cand) == 10 and self.master_list:
            self.cand_name_lbl.config(text="✗ not in master list", fg=ERROR)
        else:
            self.cand_name_lbl.config(text="")

    def _run_verify_display(self, entry: FileEntry):
        """
        Run verify_candidate_against_master_list for the current entry and
        show/hide the amber warning banner.
        Only does anything meaningful when a master list is loaded.
        """
        if not self.master_list or not entry.cand_num:
            self.verify_warn_lbl.grid_remove()
            return

        result = verify_candidate_against_master_list(
            entry.cand_num,
            self.master_list,
            cover_page_name=entry.cand_name,
        )

        # Only surface name-mismatch warnings here — number-not-found is
        # already shown by the red "✗ not in master list" label next to the field.
        name_warnings = [w for w in result.warnings if "mismatch" in w.lower()]
        if name_warnings:
            self.verify_warn_lbl.config(text="⚠️  " + name_warnings[0].lstrip("⚠️ ").strip())
            self.verify_warn_lbl.grid()
        else:
            self.verify_warn_lbl.grid_remove()

    # ─────────────────────────────────────────────────────────────────────────
    # Skip
    # ─────────────────────────────────────────────────────────────────────────

    def _toggle_skip(self):
        if self.current_idx is None:
            return
        entry = self.entries[self.current_idx]
        if entry.status == ST_SKIPPED:
            ok, _ = validate_candidate_number(entry.cand_num)
            entry.status = ST_FOUND if ok else ST_MISSING
        else:
            entry.status = ST_SKIPPED
        self._update_skip_btn(entry)
        self._refresh_listbox()
        self._update_list_status()

    def _update_skip_btn(self, entry: FileEntry):
        if entry.status == ST_SKIPPED:
            self.skip_btn.config(text="↩ Unskip", bg=ACCENT)
        else:
            self.skip_btn.config(text="↷ Skip File", bg="#7f8c8d")

    # ─────────────────────────────────────────────────────────────────────────
    # Scanning — queue-based, thread-safe
    # ─────────────────────────────────────────────────────────────────────────

    def _poll_scan_queue(self):
        try:
            while True:
                result = self._scan_queue.get_nowait()
                self._apply_scan_result(result)
        except queue.Empty:
            pass
        self.after(100, self._poll_scan_queue)

    def _apply_scan_result(self, result: dict):
        """Apply a scan result dict to the UI. Always runs on main thread."""
        if result["generation"] != self._generation:
            return

        entry = result["entry"]
        if entry not in self.entries:
            return

        idx = self.entries.index(entry)

        entry.cand_num  = result["cand_num"]
        entry.cand_name = result.get("cand_name", "")   # name from cover page
        entry.status    = result["status"]
        entry.scan_error = result.get("error", "")

        icon    = STATUS_ICON[entry.status]
        color   = STATUS_COLOR[entry.status]
        display = f" {icon}  {entry.path.name}"
        self.file_listbox.delete(idx)
        self.file_listbox.insert(idx, display)
        self.file_listbox.itemconfig(idx, fg=color)

        if self.current_idx == idx:
            self.file_listbox.selection_set(idx)
            self._suppress_cand_trace = True
            self.cand_var.set(entry.cand_num)
            self._suppress_cand_trace = False
            self._update_cand_name_label(entry.cand_num)
            self._run_verify_display(entry)        # ← verification check
            self._update_preview()

            if entry.status == ST_MISSING:
                self.status_var.set(
                    f"{entry.path.name}: no number found — enter manually.")
            elif entry.status == ST_ERROR:
                self.status_var.set(
                    f"Scan error — {entry.path.name}: {entry.scan_error}")
            else:
                self.status_var.set(
                    f"{entry.path.name}: found {entry.cand_num}")

        self._update_list_status()

    def _submit_scan(self, entry: FileEntry):
        generation = self._generation
        self._executor.submit(self._run_scan, entry, generation)

    def _run_scan(self, entry: FileEntry, generation: int):
        """
        Worker function — runs in a pool thread.
        Posts result dict to _scan_queue; never touches UI directly.
        Now also extracts the candidate name from the cover page so the
        verify step has something to compare against the master list.
        """
        import re as _re

        try:
            text, extract_error = extract_title_page_text(entry.path)
            info = extract_candidate_info(text)
            num  = info.get("number", "")
            name = info.get("name", "")      # cover page name — may be ""

            # Fallback 1: filename
            if not num:
                m = _re.search(r'\b(\d{10})\b', entry.path.stem)
                if m:
                    num = m.group(1)

            # Fallback 2: master list — match extracted name or filename stem
            if not num and self.master_list:
                extracted_name = name.strip()
                stem_clean = entry.path.stem.replace("_", " ").replace("-", " ")

                for cand_num, stored_name in self.master_list.items():
                    original, reversed_ = normalise_cxc_name(stored_name)
                    if (extracted_name and name_matches(extracted_name, original, reversed_)) or \
                       name_matches(stem_clean, original, reversed_):
                        num = cand_num
                        break

            if num:
                self._scan_queue.put({
                    "generation": generation,
                    "entry":      entry,
                    "cand_num":   num,
                    "cand_name":  name,
                    "status":     ST_FOUND,
                    "error":      "",
                })
            elif extract_error:
                self._scan_queue.put({
                    "generation": generation,
                    "entry":      entry,
                    "cand_num":   "",
                    "cand_name":  name,
                    "status":     ST_ERROR,
                    "error":      extract_error,
                })
            else:
                self._scan_queue.put({
                    "generation": generation,
                    "entry":      entry,
                    "cand_num":   "",
                    "cand_name":  name,
                    "status":     ST_MISSING,
                    "error":      "",
                })

        except Exception as exc:
            self._scan_queue.put({
                "generation": generation,
                "entry":      entry,
                "cand_num":   "",
                "cand_name":  "",
                "status":     ST_ERROR,
                "error":      str(exc),
            })

    def _scan_current(self):
        if self.current_idx is None or not self.entries:
            messagebox.showwarning("No file", "Select a file from the list first.")
            return
        entry = self.entries[self.current_idx]
        entry.status = ST_SCANNING
        entry.scan_error = ""
        self._refresh_listbox()
        self.status_var.set(f"Scanning {entry.path.name}…")
        self._submit_scan(entry)

    # ─────────────────────────────────────────────────────────────────────────
    # Subject / preview
    # ─────────────────────────────────────────────────────────────────────────

    def _on_subject_change(self, *_):
        subj = self.subj_var.get()
        code = CSEC_SUBJECTS.get(subj, "")
        self.mod_var.set(code)
        self._update_preview()

    def _update_preview(self, *_):
        cand = self.cand_var.get().strip()
        mod  = self.mod_var.get().strip()
        ft   = self.ftype_var.get()
        dn   = self.docnum_var.get()

        cand_ok, _ = validate_candidate_number(cand)
        mod_ok,  _ = validate_moderation_code(mod)

        if cand_ok and mod_ok:
            name = build_filename(cand, mod, ft, dn, ".pdf")
            self.preview_var.set(name)
        else:
            self.preview_var.set("— select subject and ensure candidate number is 10 digits —")

    # ─────────────────────────────────────────────────────────────────────────
    # Output / master list
    # ─────────────────────────────────────────────────────────────────────────

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.output_dir = Path(d)
            self.out_label.config(text=str(self.output_dir), fg="#1a252f")

    def _load_master_list(self):
        path = filedialog.askopenfilename(
            title="Select master candidate list",
            filetypes=[("Excel/CSV/PDF", "*.xlsx *.xls *.csv *.pdf"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            self.master_list = load_master_list(path)
            self.ml_label.config(
                text=f"{Path(path).name} — {len(self.master_list)} candidates",
                fg=SUCCESS,
            )
            self.status_var.set(f"Master list loaded: {len(self.master_list)} entries.")
            # Re-scan missing/errored files now that master list is available
            for entry in self.entries:
                if entry.status in (ST_MISSING, ST_ERROR):
                    entry.status = ST_SCANNING
                    self._submit_scan(entry)
            self._refresh_listbox()
            # Refresh verify display for current file with new master list
            if self.current_idx is not None:
                self._run_verify_display(self.entries[self.current_idx])
        except ValueError as e:
            messagebox.showerror("Error loading master list", str(e))
        except Exception as e:
            messagebox.showerror("Error loading master list", f"Unexpected error: {e}")

    # ─────────────────────────────────────────────────────────────────────────
    # Rename All
    # ─────────────────────────────────────────────────────────────────────────

    def _rename_all(self):
        if not self.entries:
            messagebox.showwarning("No files", "Please add files first.")
            return

        mod = self.mod_var.get().strip()
        ft  = self.ftype_var.get()
        dn  = self.docnum_var.get()

        mod_ok, mod_err = validate_moderation_code(mod)
        if not mod_ok:
            messagebox.showerror("No subject", f"Please select a subject first. ({mod_err})")
            return

        out_dir = self.output_dir
        if out_dir:
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
                test_file = out_dir / ".sba_write_test"
                test_file.touch()
                test_file.unlink()
            except OSError as e:
                messagebox.showerror("Output folder error",
                                     f"Cannot write to output folder:\n{e}")
                return

        to_process = [
            e for e in self.entries
            if e.status != ST_SKIPPED
            and validate_candidate_number(e.cand_num)[0]
        ]
        skipped   = [e for e in self.entries if e.status == ST_SKIPPED]
        no_number = [
            e for e in self.entries
            if e.status not in (ST_SKIPPED,)
            and not validate_candidate_number(e.cand_num)[0]
        ]

        if not to_process:
            messagebox.showwarning(
                "Nothing to process",
                "No files have a valid 10-digit candidate number.\n"
                "Check the file list — files marked ? need a number entered manually."
            )
            return

        # Pre-flight: name collision check
        dest_names: dict[str, str] = {}
        collisions: list[str] = []
        for entry in to_process:
            extension = entry.path.suffix.lower() or ".pdf"
            new_name = build_filename(entry.cand_num, mod, ft, dn, extension)
            if new_name in dest_names:
                collisions.append(
                    f"  • {entry.path.name} and {dest_names[new_name]} → {new_name}"
                )
            else:
                dest_names[new_name] = entry.path.name
            check_dir = out_dir or entry.path.parent
            if (check_dir / new_name).exists():
                collisions.append(f"  • {new_name} already exists in output folder")

        if collisions:
            collision_text = "\n".join(collisions)
            if not messagebox.askyesno(
                "Name collisions detected",
                f"The following output files would conflict:\n\n{collision_text}\n\n"
                "Existing files will NOT be overwritten — those jobs will fail.\n"
                "Proceed anyway?"
            ):
                return

        # Pre-flight: master list verification warnings (collected across all files)
        if self.master_list:
            verify_issues: list[str] = []
            for entry in to_process:
                result = verify_candidate_against_master_list(
                    entry.cand_num, self.master_list, entry.cand_name
                )
                for w in result.warnings:
                    verify_issues.append(f"  • {entry.path.name}: {w.lstrip('⚠️ ')}")

            if verify_issues:
                issue_text = "\n".join(verify_issues[:10])  # cap at 10 lines
                if len(verify_issues) > 10:
                    issue_text += f"\n  … and {len(verify_issues) - 10} more"
                if not messagebox.askyesno(
                    "Master list warnings",
                    f"The following candidates have issues flagged by the master list:\n\n"
                    f"{issue_text}\n\n"
                    "These are warnings only — the rename will still work.\n"
                    "Proceed anyway?"
                ):
                    return

        if no_number:
            names = "\n".join(f"  • {e.path.name}" for e in no_number)
            if not messagebox.askyesno(
                "Missing numbers",
                f"{len(no_number)} file(s) have no candidate number and will be skipped:\n\n"
                f"{names}\n\nProceed with the {len(to_process)} ready file(s)?"
            ):
                return

        results = []
        for entry in to_process:
            job = RenameJob(
                source_path=entry.path,
                candidate_number=entry.cand_num,
                moderation_code=mod,
                file_type=ft,
                doc_number=dn,
                output_dir=self.output_dir,
            )
            results.append(process_job(job, overwrite=False))

        ok_results   = [r for r in results if r.ok]
        fail_results = [r for r in results if not r.ok]

        lines = []
        if ok_results:
            lines.append(f"✓ {len(ok_results)} renamed successfully:\n")
            lines += [f"  {r.new_name}" for r in ok_results]
        if fail_results:
            lines.append(f"\n✗ {len(fail_results)} failed:\n")
            lines += [f"  {r.source.name}: {r.error}" for r in fail_results]
        if skipped:
            lines.append(f"\n↷ {len(skipped)} skipped by user.")
        if no_number:
            lines.append(f"\n? {len(no_number)} had no candidate number — not processed.")

        title = "Done" if not fail_results else f"Done with {len(fail_results)} error(s)"
        messagebox.showinfo(title, "\n".join(lines))
        self.status_var.set(
            f"Renamed {len(ok_results)} file(s). "
            f"Output: {self.output_dir or 'same folder as source'}"
        )

    # ─────────────────────────────────────────────────────────────────────────
    # About
    # ─────────────────────────────────────────────────────────────────────────

    def _show_about(self):
        win = tk.Toplevel(self)
        win.title("About")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="SBA File Renamer", bg=BG,
                 font=("Segoe UI", 13, "bold"), fg=ACCENT).pack(padx=24, pady=(20, 4))
        tk.Label(win,
                 text="Originally built for\nCaribbean Union College Secondary School\n"
                      "Maracas, St. Joseph, Trinidad.",
                 bg=BG, font=("Segoe UI", 9), fg="#5d6d7e",
                 justify="center").pack(padx=24, pady=(0, 8))
        tk.Label(win,
                 text="Renames CSEC SBA files to CXC required format.\n"
                      "Free to use and share.",
                 bg=BG, font=("Segoe UI", 9), fg="#1a252f",
                 justify="center").pack(padx=24, pady=(0, 16))
        tk.Button(win, text="Close", command=win.destroy,
                  bg=ACCENT, fg=WHITE, relief="flat", padx=16, pady=4).pack(pady=(0, 16))


if __name__ == "__main__":
    app = SBARenamerApp()
    app.mainloop()
